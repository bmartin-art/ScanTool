/* ==========================================================================
   Duct Tracker — Cloudflare Worker
   1) HTTP API over D1 (unchanged): GET/PUT/DELETE /jobs[/:name]  (needs X-Access-Token)
   2) Scheduled cron: pulls new/changed CSVs from a OneDrive/SharePoint folder
      via Microsoft Graph, parses them with the same logic as the app, and
      upserts jobs into D1 — preserving scan progress on re-exports.
   3) Completion email: on PUT, if an order just reached "Complete" and has a
      customer email, sends one notice through the sales Gmail account (Gmail API,
      OAuth refresh token) — lands in the sales Sent folder, replies go to sales.
      Diagnostics: GET /ingest-now (ingest report), GET /mail-test?to=you@x (send test).

   Bindings:  DB (D1)
   Secrets (wrangler secret put / dashboard):
     ACCESS_TOKEN   shared code the app sends (already set)
     TENANT_ID      Azure AD tenant (directory) ID
     CLIENT_ID      app registration (client) ID
     CLIENT_SECRET  a client secret value for that app registration
     GMAIL_CLIENT_ID      OAuth Web client ID for the sales-mail sender
     GMAIL_CLIENT_SECRET  that Web client's secret
     GMAIL_REFRESH_TOKEN  refresh token authorized for gmail.send as the sender
   Vars ([vars] in wrangler.toml or dashboard):
     ALLOW_ORIGIN   e.g. https://bmartin-art.github.io  (default "*")
     SHARE_URL      the OneDrive folder share link to watch
     GMAIL_SENDER   address to send completion emails from (default sales@brookstoneind.com)
   ========================================================================== */

const GRAPH = "https://graph.microsoft.com/v1.0";
const STOCK_GROUP = "STCKCAT";

export default {
  /* -------------------------- HTTP API -------------------------- */
  async fetch(req, env) {
    const origin = env.ALLOW_ORIGIN || "*";
    const cors = {
      "Access-Control-Allow-Origin": origin,
      "Access-Control-Allow-Methods": "GET,PUT,DELETE,OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Access-Token",
      "Access-Control-Max-Age": "86400",
    };
    if (req.method === "OPTIONS") return new Response(null, { headers: cors });

    const tok =
      req.headers.get("X-Access-Token") ||
      (req.headers.get("Authorization") || "").replace(/^Bearer\s+/i, "");
    if (!env.ACCESS_TOKEN || tok !== env.ACCESS_TOKEN)
      return json({ error: "unauthorized" }, 401, cors);

    const parts = new URL(req.url).pathname.split("/").filter(Boolean);
    const name = parts[1] ? decodeURIComponent(parts[1]) : null;

    try {
      if (parts[0] === "ingest-now") {          // manual trigger + diagnostics
        const report = await runIngest(env);
        return json(report, 200, cors);
      }
      if (parts[0] === "mail-test") {            // verify sales-mailbox sending works
        const to = new URL(req.url).searchParams.get("to");
        if (!to) return json({ ok: false, error: "add ?to=you@example.com" }, 400, cors);
        const r = await sendCompletionEmail(env, { orderNo: "TEST", title: "Mail test", customer: "", email: to, otype: "Pickup" });
        return json(r, 200, cors);
      }
      if (parts[0] !== "jobs") return json({ error: "not found" }, 404, cors);

      if (req.method === "GET" && !name) {
        const { results } = await env.DB
          .prepare("SELECT data FROM jobs ORDER BY created_at DESC").all();
        return json(results.map((r) => JSON.parse(r.data)), 200, cors);
      }
      if (req.method === "GET" && name) {
        const row = await env.DB.prepare("SELECT data FROM jobs WHERE name = ?").bind(name).first();
        return json(row ? JSON.parse(row.data) : null, 200, cors);
      }
      if (req.method === "PUT" && name) {
        const jobObj = JSON.parse(await req.text());
        // Completion email: fire exactly once, server-side. The STORED record's
        // emailSent flag is the source of truth, so multiple devices / repeat
        // saves can't trigger a second send.
        let mail = null;
        try {
          const prev = await env.DB.prepare("SELECT data FROM jobs WHERE name = ?").bind(name).first();
          const prevSent = prev ? JSON.parse(prev.data).emailSent : false;
          if (prevSent) {
            jobObj.emailSent = true; jobObj.emailedAt = prev ? JSON.parse(prev.data).emailedAt : Date.now();
          } else if (jobObj.status === "Complete" && jobObj.email && env.GMAIL_REFRESH_TOKEN) {
            mail = await sendCompletionEmail(env, jobObj);
            if (mail.ok) { jobObj.emailSent = true; jobObj.emailedAt = Date.now(); }
          }
        } catch (e) { mail = { ok: false, error: String((e && e.message) || e) }; }
        await upsertJob(env, jobObj);      // app writes are authoritative (full object)
        return json({ ok: true, mail }, 200, cors);
      }
      if (req.method === "DELETE" && name) {
        await env.DB.prepare("DELETE FROM jobs WHERE name = ?").bind(name).run();
        return json({ ok: true }, 200, cors);
      }
      return json({ error: "method not allowed" }, 405, cors);
    } catch (e) {
      return json({ error: String((e && e.message) || e) }, 500, cors);
    }
  },

  /* -------------------------- Cron ingester -------------------------- */
  async scheduled(event, env, ctx) {
    ctx.waitUntil(runIngest(env).then((r) => console.log("cron ingest:", JSON.stringify(r))));
  },
};

/* ===================== OneDrive / Graph ingest ===================== */
async function runIngest(env) {
  const report = { at: new Date().toISOString(), steps: [], files: [], results: [], errors: [] };
  if (!env.SHARE_URL) { report.errors.push("SHARE_URL variable is not set"); return report; }
  if (!env.TENANT_ID || !env.CLIENT_ID || !env.CLIENT_SECRET) {
    report.errors.push("Missing one of TENANT_ID / CLIENT_ID / CLIENT_SECRET secrets"); return report;
  }
  let token;
  try { token = await getGraphToken(env); report.steps.push("Graph token: OK"); }
  catch (e) { report.errors.push("Graph token failed: " + e.message); return report; }

  let children;
  try { children = await listFolder(token, env.SHARE_URL); report.steps.push("Folder read: " + children.length + " item(s)"); }
  catch (e) { report.errors.push("Folder list failed: " + e.message); return report; }

  const csvs = children.filter((c) => c.file && /\.csv$/i.test(c.name || ""));
  report.files = children.map((c) => c.name + (c.file ? "" : " (folder)"));
  if (!csvs.length) { report.errors.push("No .csv files found in that folder"); return report; }

  for (const f of csvs) {
    try {
      const etag = f.eTag || f.cTag || "";
      const seen = await env.DB.prepare("SELECT etag FROM ingested_files WHERE file_id = ?").bind(f.id).first();
      if (seen && seen.etag === etag) { report.results.push({ file: f.name, action: "skipped (already current)" }); continue; }
      const url = f["@microsoft.graph.downloadUrl"];
      if (!url) { report.errors.push(f.name + ": no download URL from Graph"); continue; }
      const text = await (await fetch(url)).text();
      const code = ((f.name || "").match(/\+(\d{1,4})/) || [])[1] || "";   // +#### match code appended by the programmer
      const jobName = (f.name || "").replace(/\.csv$/i, "").replace(/\+\d{1,4}/, "").replace(/_/g, " ").trim();
      const fresh = buildJobFromCSV(text, jobName);
      if (!fresh.items.length) { report.results.push({ file: f.name, action: "no scannable items — skipped" }); continue; }

      // Find the base record to merge into: first by real order number, else claim
      // the one open temp cover sheet holding this +code.
      let baseRow = await env.DB.prepare("SELECT data FROM jobs WHERE name = ?").bind(fresh.name).first();
      let base = baseRow ? JSON.parse(baseRow.data) : null;
      let claimedTemp = "";
      if (!base && code) {
        const all = (await env.DB.prepare("SELECT data FROM jobs").all()).results.map((r) => JSON.parse(r.data));
        const hits = all.filter((j) => j.cover && j.cover.matchCode === code && j.status !== "Complete" && /^TMP-/.test(j.name));
        if (hits.length === 1) { base = hits[0]; claimedTemp = base.name; }
        else if (hits.length > 1) report.errors.push(f.name + ": +" + code + " matches " + hits.length + " open sheets — imported as new");
      }

      const merged = mergePreserveScans(base, fresh);
      merged.orderNo = fresh.orderNo; merged.name = fresh.name;   // adopt the real order number
      if ((merged.status === "Received" || !merged.status) && merged.items.length) merged.status = "Programmed"; // CAMduct export = job programmed
      await upsertJob(env, merged);
      if (claimedTemp && claimedTemp !== merged.name) await env.DB.prepare("DELETE FROM jobs WHERE name = ?").bind(claimedTemp).run();
      await env.DB.prepare(
        "INSERT INTO ingested_files (file_id, etag, order_no, ingested_at) VALUES (?,?,?,?) " +
        "ON CONFLICT(file_id) DO UPDATE SET etag=excluded.etag, order_no=excluded.order_no, ingested_at=excluded.ingested_at"
      ).bind(f.id, etag, fresh.name, Date.now()).run();
      report.results.push({ file: f.name, action: claimedTemp ? "claimed +" + code : (base ? "merged (kept scans)" : "created"), order: fresh.name, items: fresh.items.length });
    } catch (e) { report.errors.push((f.name || "?") + ": " + e.message); }
  }
  return report;
}

async function getGraphToken(env) {
  const body = new URLSearchParams({
    client_id: env.CLIENT_ID,
    client_secret: env.CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });
  const res = await fetch(
    `https://login.microsoftonline.com/${env.TENANT_ID}/oauth2/v2.0/token`,
    { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body }
  );
  if (!res.ok) throw new Error("token " + res.status + " " + (await res.text()));
  return (await res.json()).access_token;
}

function escapeHtml(s) { return String(s == null ? "" : s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }
function hclean(s) { return String(s == null ? "" : s).replace(/[\r\n]+/g, " ").trim(); }   // no header injection
function b64url(str) {
  const bytes = new TextEncoder().encode(str);
  let bin = ""; for (const b of bytes) bin += String.fromCharCode(b);
  return btoa(bin).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

// Exchange the stored refresh token for a short-lived Gmail access token.
async function gmailToken(env) {
  const body = new URLSearchParams({
    client_id: env.GMAIL_CLIENT_ID,
    client_secret: env.GMAIL_CLIENT_SECRET,
    refresh_token: env.GMAIL_REFRESH_TOKEN,
    grant_type: "refresh_token",
  });
  const res = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body,
  });
  if (!res.ok) throw new Error("gmail token " + res.status + " " + (await res.text()));
  return (await res.json()).access_token;
}

// Send the "order complete" notice through the sales Gmail account itself, so it
// files in the sales Sent folder and replies come back to sales. Sends as
// GMAIL_SENDER using the OAuth refresh token authorized for that mailbox.
async function sendCompletionEmail(env, order) {
  if (!env.GMAIL_REFRESH_TOKEN || !env.GMAIL_CLIENT_ID || !env.GMAIL_CLIENT_SECRET)
    return { ok: false, error: "Gmail credentials not set" };
  if (!order.email) return { ok: false, error: "no customer email on the order" };
  const sender = env.GMAIL_SENDER || "sales@brookstoneind.com";
  let token;
  try { token = await gmailToken(env); } catch (e) { return { ok: false, error: "token: " + e.message }; }

  const contactH = escapeHtml((order.contact || "").trim());
  // Job name for the notice: the cover sheet's Ship To / Job wins over the
  // CAMduct-filename-derived title, so the customer sees the real job name.
  const jobName = ((order.cover && order.cover.job) || order.title || "the order").trim();
  const titleH = escapeHtml(jobName);
  const subject = hclean(`${((order.cover && order.cover.job) || order.title || "Your order").trim()} is complete`);
  const html =
    `<p>Hello${contactH ? " " + contactH : ""},</p>` +
    `<p>Your ${titleH} is complete.</p>` +
    `<p>Thanks,<br>Brookstone Industries<br>161 Zooks Mill Road, Ephrata, PA 17522<br>717-859-3340</p>`;
  const mime = [
    `From: Brookstone Industries <${sender}>`,
    `To: ${hclean(order.email)}`,
    `Subject: ${subject}`,
    "MIME-Version: 1.0",
    'Content-Type: text/html; charset="UTF-8"',
    "", html,
  ].join("\r\n");

  const res = await fetch("https://gmail.googleapis.com/gmail/v1/users/me/messages/send", {
    method: "POST",
    headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({ raw: b64url(mime) }),
  });
  if (res.ok) return { ok: true, to: order.email, subject };
  return { ok: false, error: "gmail send " + res.status + " " + (await res.text()) };
}

// Resolve a sharing URL to its folder, then list child files from the drive.
// Reading children via /drives/{driveId}/items/{itemId}/children reliably
// includes @microsoft.graph.downloadUrl (the /shares/.../children path does not).
async function listFolder(token, shareUrl) {
  const shareId = "u!" + btoa(shareUrl).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
  const h = { Authorization: "Bearer " + token };

  // 1) resolve the share to the actual folder driveItem (its drive + item id)
  const metaRes = await fetch(`${GRAPH}/shares/${shareId}/driveItem?$select=id,parentReference`, { headers: h });
  if (!metaRes.ok) throw new Error("share resolve " + metaRes.status + " " + (await metaRes.text()));
  const meta = await metaRes.json();
  const driveId = meta.parentReference && meta.parentReference.driveId;
  const itemId = meta.id;
  if (!driveId || !itemId) throw new Error("could not resolve drive/item from share");

  // 2) list children straight from the drive (download URLs come back here)
  let url = `${GRAPH}/drives/${driveId}/items/${itemId}/children?$top=200`;
  const out = [];
  while (url) {
    const res = await fetch(url, { headers: h });
    if (!res.ok) throw new Error("graph list " + res.status + " " + (await res.text()));
    const data = await res.json();
    out.push(...(data.value || []));
    url = data["@odata.nextLink"] || null;
  }
  return out;
}

/* ===================== merge: keep scan progress ===================== */
function mergePreserveScans(existing, fresh) {
  if (!existing) return fresh;
  const byKey = {};
  for (const it of existing.items) byKey[it.key] = it;
  for (const it of fresh.items) {
    const old = byKey[it.key];
    if (old && typeof old.onTruck === "number")
      it.onTruck = Math.min(old.onTruck, it.qty); // carry scans forward, cap at new qty
  }
  fresh.pallets = existing.pallets || [];
  fresh.palletSeq = existing.palletSeq || 0;
  fresh.location = existing.location || fresh.location || "";
  fresh.createdAt = existing.createdAt || fresh.createdAt;
  // preserve front-office fields entered in the app (CSV doesn't carry these)
  fresh.company = existing.company || fresh.company || "";
  fresh.email = existing.email || fresh.email || "";
  fresh.otype = existing.otype || fresh.otype || "";
  fresh.needBy = existing.needBy || fresh.needBy || "";
  fresh.priority = existing.priority || false;
  fresh.status = existing.status || fresh.status || "Received";
  fresh.scheduledDate = existing.scheduledDate || fresh.scheduledDate || "";
  fresh.cover = existing.cover || fresh.cover || {};
  fresh.emailSent = existing.emailSent || false;
  fresh.emailedAt = existing.emailedAt || null;
  if (existing.customer) fresh.customer = existing.customer;   // keep an edited customer name
  if (existing.title) fresh.title = existing.title;
  fresh.contact = existing.contact || fresh.contact || "";
  fresh.phone = existing.phone || fresh.phone || "";
  fresh.po = existing.po || fresh.po || "";
  if (existing.laborHours != null && fresh.laborHours == null) fresh.laborHours = existing.laborHours;
  fresh._merged = true;
  return fresh;
}

async function upsertJob(env, jobObj) {
  const body = JSON.stringify(jobObj);
  await env.DB.prepare(
    "INSERT INTO jobs (name, created_at, data, updated_at) VALUES (?,?,?,?) " +
    "ON CONFLICT(name) DO UPDATE SET data=excluded.data, updated_at=excluded.updated_at"
  ).bind(jobObj.name, jobObj.createdAt || Date.now(), body, Date.now()).run();
}

/* ===================== parser (ported from the app) ===================== */
function parseCSV(text) {
  const rows = [];
  for (const line of text.replace(/\r\n?/g, "\n").split("\n")) {
    if (line.trim() === "") continue;
    const f = []; let i = 0;
    while (i <= line.length) {
      let v = "";
      if (line[i] === '"') {
        i++;
        while (i < line.length) {
          if (line[i] === '"' && (line[i + 1] === "," || i + 1 >= line.length)) { i++; break; }
          v += line[i++];
        }
      } else { while (i < line.length && line[i] !== ",") v += line[i++]; }
      f.push(v);
      if (line[i] === ",") { i++; if (i > line.length) break; } else break;
    }
    rows.push(f);
  }
  return rows;
}
function col(H, names) {
  for (const n of names) { const i = H.findIndex((h) => h.trim().toLowerCase() === n); if (i >= 0) return i; }
  return -1;
}
function parseFtime(s) {                        // CAMduct "F time" H:MM:SS -> hours
  s = (s || "").trim(); if (!s) return 0;
  const p = s.split(":").map((n) => parseInt(n, 10));
  if (p.some(isNaN)) return 0;
  let sec = 0;
  if (p.length === 3) sec = p[0] * 3600 + p[1] * 60 + p[2];
  else if (p.length === 2) sec = p[0] * 60 + p[1];
  else sec = p[0];
  return sec / 3600;
}
function buildJobFromCSV(text, jobName) {
  const rows = parseCSV(text);
  if (!rows.length) throw new Error("empty");
  const H = rows[0].map((h) => h.trim().toLowerCase());
  const iGroup = col(H, ["epicor group", "group"]);
  const iFtime = col(H, ["f time", "ftime", "fab time", "time"]);
  const iB64 = col(H, ["item globally unique id (base64)", "base64", "item scan code (base64)"]);
  const iGuid = col(H, ["item globally unique id", "guid"]);
  const iQty = col(H, ["qty", "quantity"]);
  const iDesc = col(H, ["description", "desc"]);
  const iItem = col(H, ["item number", "item no", "itemnumber", "item"]);
  const iNotes = col(H, ["notes", "customer", "customer name", "ship to"]);
  const iOrder = col(H, ["order number", "order no", "order #", "order"]);
  if (iB64 < 0) throw new Error("no base64 column");

  const items = []; let seq = 0, customer = "", orderNo = "", fabHours = 0;
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r]; if (!row.length || row.every((c) => c === "")) continue;
    const group = (row[iGroup] || "").trim();
    const desc = (row[iDesc] || "").trim();
    const qty = parseInt((row[iQty] || "0").trim(), 10) || 0;
    const key = (row[iB64] || "").trim();
    const feach = iFtime >= 0 ? parseFtime(row[iFtime]) : 0;
    if (!customer && iNotes >= 0 && (row[iNotes] || "").trim()) customer = (row[iNotes] || "").trim();
    if (!orderNo && iOrder >= 0 && (row[iOrder] || "").trim()) orderNo = (row[iOrder] || "").trim();
    if (/labor/i.test(desc)) continue;
    if (qty <= 0) continue;
    const keys = key.split(",").map((s) => s.trim()).filter(Boolean);
    const nGuid = Math.max(1, keys.length);
    const feachPer = feach / nGuid;    // merged lines share one F time across their GUIDs -> per-piece time
    fabHours += feachPer * qty;
    seq++;
    const itemNoRaw = (iItem >= 0 ? (row[iItem] || "").trim() : "") || String(seq);
    const type = group.toUpperCase() === STOCK_GROUP ? "stock" : "fab";
    const guid = (row[iGuid] || "").trim();
    const nums = itemNoRaw.split(",").map((s) => s.trim()).filter(Boolean);
    if (type === "stock" || keys.length <= 1) {
      items.push({ key, guid, itemNo: itemNoRaw, group, desc: desc || "(no description)", qty, type, onTruck: 0, fabEach: feachPer });
    } else {
      const per = Math.floor(qty / keys.length), extra = qty - per * keys.length;
      keys.forEach((k, idx) => {
        const q = per + (idx < extra ? 1 : 0); if (q <= 0) return;
        items.push({ key: k, guid: "", itemNo: nums[idx] || (itemNoRaw + "." + (idx + 1)), group, desc: desc || "(no description)", qty: q, type, onTruck: 0, fabEach: feachPer });
      });
    }
  }
  const fab = items.filter((i) => i.type === "fab"), stock = items.filter((i) => i.type === "stock");
  return {
    name: orderNo || jobName, orderNo, customer, title: jobName, location: "",
    company: "", email: "", otype: "", needBy: "", priority: false, status: "Received", scheduledDate: "",
    laborHours: iFtime >= 0 ? Math.round(fabHours * 100) / 100 : undefined,
    cover: {},
    createdAt: Date.now(), items, pallets: [], palletSeq: 0,
    _summary: { fabLines: fab.length, fabPieces: fab.reduce((s, i) => s + i.qty, 0), stockLines: stock.length, fabHours: Math.round(fabHours * 100) / 100 },
  };
}

function json(obj, status, cors) {
  return new Response(JSON.stringify(obj), { status, headers: { ...cors, "Content-Type": "application/json" } });
}
