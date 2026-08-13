// JobTread integration proxy (Cloudflare Worker) — secure replacement for the
// Google Apps Script proxy, which was blocked by a Workspace admin policy that
// disallows anonymous execution of Apps Script web apps.
//
// ONE-TIME SETUP:
//   1. `wrangler secret put JOBTREAD_GRANT_KEY`   (paste the real JobTread API grant key)
//   2. `wrangler secret put JT_PROXY_SECRET`      (any random string; also goes into
//      index.html's JOBTREAD_PROXY_KEY constant — gates access to this proxy, the real
//      grant key never leaves this Worker)
//   3. Optional: `wrangler secret put JOBTREAD_ORGANIZATION_ID` to skip auto-resolution.
//   4. `wrangler deploy` — copy the resulting workers.dev URL into index.html's
//      JOBTREAD_PROXY_URL constant.
//
// Twilio SMS setup, for the "Send via Text" button:
//   1. In the Twilio Console, use an SMS-capable Twilio phone number (the existing internal
//      Twilio account already approved for another app can be reused).
//   2. `wrangler secret put TWILIO_ACCOUNT_SID`
//      `wrangler secret put TWILIO_AUTH_TOKEN`
//      `wrangler secret put TWILIO_FROM_NUMBER`     (E.164 format, e.g. +13525551234)
//
// Take-off PDF upload setup, for the super-admin "Upload Take-Off" button. Uses Cloudflare R2
// (kept fully PRIVATE — accessed only via this Worker's binding, never a public bucket URL):
//   1. Enable R2 for this Cloudflare account (dash.cloudflare.com > R2 > Enable) if not already.
//   2. `wrangler r2 bucket create konchan-schedule-takeoffs`
//   3. Add to wrangler.toml:
//        [[r2_buckets]]
//        binding = "TAKEOFFS_BUCKET"
//        bucket_name = "konchan-schedule-takeoffs"
//   4. `wrangler deploy`
//
// Endpoints:
//   GET  ?action=jtSearch&q=<text>&key=<secret>
//   GET  ?action=jtJobDetail&jobId=&accountId=&key=<secret>
//   POST ?action=sendSms&key=<secret>   body: {"to":"<digits w/ country code>","body":"<message text>"}
//   POST ?action=takeoffUpload&key=<secret>   body: {"fileName":"","contentType":"application/pdf","data":"<base64>"}
//   GET  /takeoff/<id>   — serves an uploaded PDF back out (no secret; meant to be opened
//                          directly, e.g. from a texted SMS link, by people with no login)

const JOBTREAD_PAVE_URL = "https://api.jobtread.com/pave";
const ALLOWED_ORIGINS = new Set([
  "https://konchan-schedule.web.app",
  "https://konchan-schedule.firebaseapp.com",
]);

function corsHeaders(origin) {
  const allow = ALLOWED_ORIGINS.has(origin) ? origin : "https://konchan-schedule.web.app";
  return {
    "Access-Control-Allow-Origin": allow,
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
  };
}

function jsonResponse(obj, origin) {
  return new Response(JSON.stringify(obj), {
    headers: { "Content-Type": "application/json", ...corsHeaders(origin) },
  });
}

function checkSecret(url, env) {
  const expected = env.JT_PROXY_SECRET;
  return !!expected && url.searchParams.get("key") === expected;
}

async function jtQuery(query, env) {
  const grantKey = env.JOBTREAD_GRANT_KEY;
  if (!grantKey) throw new Error("JOBTREAD_GRANT_KEY secret is not set");
  const res = await fetch(JOBTREAD_PAVE_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ query: { $: { grantKey, notify: false }, ...query } }),
  });
  const text = await res.text();
  if (!res.ok) {
    throw new Error(`JobTread request failed (${res.status}): ${text.slice(0, 300)}`);
  }
  const payload = JSON.parse(text || "{}");
  if (payload.errors && payload.errors.length) {
    const first = payload.errors[0];
    throw new Error(`JobTread query error: ${first.message || JSON.stringify(first)}`);
  }
  return payload;
}

function getByPath(obj, path) {
  return path.split(".").reduce((acc, key) => (acc && typeof acc === "object" ? acc[key] : undefined), obj);
}

async function resolveOrganizationId(env) {
  if (env.JOBTREAD_ORGANIZATION_ID) return env.JOBTREAD_ORGANIZATION_ID;
  const payload = await jtQuery(
    { currentGrant: { user: { memberships: { nodes: { id: {}, organization: { id: {}, name: {} } } } } } },
    env
  );
  const nodes = getByPath(payload, "currentGrant.user.memberships.nodes") || [];
  const first = nodes[0];
  if (!first) throw new Error("Unable to resolve JobTread organization from the grant");
  return getByPath(first, "organization.id");
}

// Searches jobs directly via organization.jobs (confirmed a valid Pave connection — JobTread's
// own docs show it, e.g. "Count Open Jobs" example — despite an earlier assumption in this file
// that Pave had no organization.jobs field). Uses a server-side `where`/`like` filter across job
// name, account name, and job address so JobTread itself does the matching in one request,
// instead of paginating everything down and filtering client-side.
//
// SUPERSEDES two earlier approaches that both missed real jobs:
// (1) deriving jobs from the org's 100 most-recently-created documents — missed jobs with no
//     recent (or any) attached document, e.g. "Diadem Wellington".
// (2) walking organization.accounts -> jobs — the nested `jobs` sub-connection silently caps at
//     a small default page size (~10) with no way to paginate it per-account, so accounts with
//     more jobs than that silently lost matches, e.g. "High Springs Plaza North Bldg" (11th job
//     on its account, from Dec 2024).
async function searchJobs(searchText, env) {
  const q = (searchText || "").trim();
  if (!q) return [];
  const organizationId = await resolveOrganizationId(env);
  const pattern = `%${q}%`;
  const payload = await jtQuery(
    {
      organization: {
        $: { id: organizationId },
        jobs: {
          $: {
            size: 20,
            where: {
              or: [
                ["name", "like", pattern],
                [["location", "account", "name"], "like", pattern],
                [["location", "address"], "like", pattern],
              ],
            },
          },
          nodes: {
            id: {},
            name: {},
            location: { address: {}, account: { id: {}, name: {} } },
          },
        },
      },
    },
    env
  );
  const nodes = getByPath(payload, "organization.jobs.nodes") || [];
  return nodes.map((job) => ({
    jobId: job.id,
    accountId: getByPath(job, "location.account.id") || "",
    name: job.name || "",
    accountName: getByPath(job, "location.account.name") || "",
    address: getByPath(job, "location.address") || "",
  }));
}

// Contact lookup. VERIFIED live: email/phone are NOT direct scalar fields on Contact — per
// JobTread's own API docs ("Find a Contact by Custom Field Value"), they're stored as
// customFieldValues where customField.type is "emailAddress" or "phoneNumber". A contact can
// have multiple phone numbers; all are returned (semicolon-joined) rather than picking just one.
async function fetchAccountContact(accountId, env) {
  if (!accountId) return { name: "", email: "", phone: "" };
  const payload = await jtQuery(
    {
      account: {
        $: { id: accountId },
        contacts: {
          nodes: {
            id: {},
            name: {},
            customFieldValues: { nodes: { value: {}, customField: { type: {} } } },
          },
        },
      },
    },
    env
  );
  const contacts = getByPath(payload, "account.contacts.nodes") || [];
  const first = contacts[0];
  if (!first) return { name: "", email: "", phone: "" };
  const values = getByPath(first, "customFieldValues.nodes") || [];
  const emails = values.filter((v) => getByPath(v, "customField.type") === "emailAddress").map((v) => v.value);
  const phones = values.filter((v) => getByPath(v, "customField.type") === "phoneNumber").map((v) => v.value);
  return { name: first.name || "", email: emails[0] || "", phone: phones.join("; ") };
}

// PDF take-off lookup. VERIFIED live: job.files.nodes exposes id/name/url/type (type is a MIME
// type like "application/pdf", not "contentType" as originally guessed). Returns ALL matching
// PDFs (a job can have several) — the caller/UI decides which ones to attach.
async function fetchJobFiles(jobId, env) {
  if (!jobId) return [];
  const payload = await jtQuery({ job: { $: { id: jobId }, files: { nodes: { id: {}, name: {}, url: {}, type: {} } } } }, env);
  const nodes = getByPath(payload, "job.files.nodes") || [];
  return nodes
    .map((n) => ({ id: n.id || "", name: n.name || "", url: n.url || "", contentType: n.type || "" }))
    .filter((f) => {
      const n = f.name.toLowerCase();
      return n.includes(".pdf") || (f.contentType || "").toLowerCase().includes("pdf") || n.includes("takeoff") || n.includes("take-off");
    })
    .sort((a, b) => {
      const score = (f) => (f.name.toLowerCase().includes("takeoff") || f.name.toLowerCase().includes("take-off") ? 0 : 1);
      return score(a) - score(b);
    });
}

// Bid line items + total. job.documents.nodes exposes id/name/type/createdAt (confirmed
// working elsewhere in this project); the Bid document's `type` enum value is
// "customerOrder" (JobTread's internal name for what's shown to users as "Customer Order
// (Bid)"). Picks the most-recently-created matching document, then fetches its full line-item
// breakdown + total via a separate document-by-id query (same pattern as ARTracker's own Bids
// integration). Returns null (rather than throwing) on any failure so a schema surprise here
// never breaks the contact/files half of jtJobDetail.
async function fetchJobBid(jobId, env) {
  if (!jobId) return null;
  try {
    const listPayload = await jtQuery(
      { job: { $: { id: jobId }, documents: { nodes: { id: {}, name: {}, type: {}, createdAt: {} } } } },
      env
    );
    const docs = getByPath(listPayload, "job.documents.nodes") || [];
    const bidDocs = docs.filter((d) => d.type === "customerOrder");
    if (!bidDocs.length) return null;
    bidDocs.sort((a, b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0));
    const bidDocId = bidDocs[0].id;

    const detailPayload = await jtQuery(
      {
        document: {
          $: { id: bidDocId },
          id: {},
          name: {},
          price: {},
          costItems: {
            nodes: { id: {}, name: {}, description: {}, quantity: {}, price: {}, costType: { name: {} } },
          },
        },
      },
      env
    );
    const doc = getByPath(detailPayload, "document") || {};
    const lineItems = (getByPath(doc, "costItems.nodes") || []).map((n) => ({
      name: n.name || "",
      description: n.description || "",
      quantity: n.quantity != null ? n.quantity : null,
      price: n.price != null ? Number(n.price) : 0,
      costType: getByPath(n, "costType.name") || "",
    }));
    return {
      id: doc.id || bidDocId,
      name: doc.name || "",
      total: Number(doc.price || 0),
      lineItems,
    };
  } catch (err) {
    return null;
  }
}

const TAKEOFF_PREFIX = "takeoffs/";
const TAKEOFF_MAX_BYTES = 15 * 1024 * 1024; // 15MB

// Uploads a take-off PDF into the (private) R2 bucket and returns a random UUID that
// GET /takeoff/<id> can later use to serve it back out.
async function uploadTakeoffPdf(fileName, contentType, base64Data, env) {
  if (!env.TAKEOFFS_BUCKET) throw new Error("Take-off storage is not configured (missing TAKEOFFS_BUCKET R2 binding).");
  if (contentType !== "application/pdf") throw new Error("Only PDF files are supported");
  if (!base64Data) throw new Error("Missing file data");
  const binary = atob(base64Data);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  if (!bytes.length) throw new Error("File is empty");
  if (bytes.length > TAKEOFF_MAX_BYTES) throw new Error("File is too large (15MB max)");
  const id = crypto.randomUUID();
  await env.TAKEOFFS_BUCKET.put(`${TAKEOFF_PREFIX}${id}.pdf`, bytes, { httpMetadata: { contentType: "application/pdf" } });
  return { id, name: fileName || "Take-Off.pdf" };
}

// Streams a previously-uploaded take-off PDF back out of the (private) R2 bucket.
async function fetchTakeoffPdf(id, env) {
  if (!env.TAKEOFFS_BUCKET) throw new Error("Take-off storage is not configured (missing TAKEOFFS_BUCKET R2 binding).");
  return env.TAKEOFFS_BUCKET.get(`${TAKEOFF_PREFIX}${id}.pdf`);
}

// Sends `bodyText` as a plain SMS via the Twilio REST API.
async function sendTwilioSms(to, bodyText, env) {
  const accountSid = env.TWILIO_ACCOUNT_SID;
  const authToken = env.TWILIO_AUTH_TOKEN;
  const fromNumber = env.TWILIO_FROM_NUMBER;
  if (!accountSid || !authToken || !fromNumber) {
    throw new Error("Twilio is not configured (missing TWILIO_ACCOUNT_SID / TWILIO_AUTH_TOKEN / TWILIO_FROM_NUMBER secrets).");
  }
  const toE164 = to.startsWith("+") ? to : `+${to}`;
  const params = new URLSearchParams({ To: toE164, From: fromNumber, Body: bodyText });
  const res = await fetch(`https://api.twilio.com/2010-04-01/Accounts/${accountSid}/Messages.json`, {
    method: "POST",
    headers: {
      Authorization: `Basic ${btoa(`${accountSid}:${authToken}`)}`,
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: params.toString(),
  });
  const text = await res.text();
  let payload;
  try {
    payload = JSON.parse(text || "{}");
  } catch {
    payload = { raw: text };
  }
  if (!res.ok) {
    const msg = (payload && payload.message) || text.slice(0, 300);
    throw new Error(`SMS send failed (${res.status}): ${msg}`);
  }
  return payload;
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const origin = request.headers.get("Origin") || "";

    if (request.method === "OPTIONS") {
      return new Response(null, { headers: corsHeaders(origin) });
    }

    // Take-off PDF serve route — deliberately outside the ?action=/key= gate below, since these
    // links are meant to be opened directly (e.g. clicked from a texted SMS) by people with no
    // login. The id is a random UUID with no directory-traversal risk.
    const takeoffMatch = url.pathname.match(/^\/takeoff\/([0-9a-fA-F-]{36})$/);
    if (takeoffMatch) {
      try {
        const object = await fetchTakeoffPdf(takeoffMatch[1], env);
        if (!object) return new Response("Not found", { status: 404 });
        return new Response(object.body, {
          headers: {
            "Content-Type": "application/pdf",
            "Content-Disposition": "inline",
            "Cache-Control": "private, max-age=86400",
          },
        });
      } catch (err) {
        return new Response(err.message || "Failed to load file", { status: 500 });
      }
    }

    const action = url.searchParams.get("action");
    const knownActions = new Set(["jtSearch", "jtJobDetail", "sendSms", "takeoffUpload"]);
    if (!knownActions.has(action)) {
      return jsonResponse({ ok: false, error: "Unknown action" }, origin);
    }
    if (!checkSecret(url, env)) {
      return jsonResponse({ ok: false, error: "Unauthorized" }, origin);
    }

    try {
      if (action === "jtSearch") {
        const jobs = await searchJobs(url.searchParams.get("q") || "", env);
        return jsonResponse({ ok: true, jobs }, origin);
      }
      if (action === "sendSms") {
        if (request.method !== "POST") {
          return jsonResponse({ ok: false, error: "sendSms requires POST" }, origin);
        }
        const body = await request.json();
        const to = String(body.to || "").replace(/[^\d+]/g, "");
        const bodyText = String(body.body || "").slice(0, 1600);
        if (!to) return jsonResponse({ ok: false, error: "Missing recipient phone number" }, origin);
        if (!bodyText) return jsonResponse({ ok: false, error: "Missing message body" }, origin);
        const result = await sendTwilioSms(to, bodyText, env);
        return jsonResponse({ ok: true, result }, origin);
      }
      if (action === "takeoffUpload") {
        if (request.method !== "POST") {
          return jsonResponse({ ok: false, error: "takeoffUpload requires POST" }, origin);
        }
        const body = await request.json();
        const fileName = typeof body.fileName === "string" ? body.fileName.slice(0, 200) : "";
        const contentType = typeof body.contentType === "string" ? body.contentType : "";
        const data = typeof body.data === "string" ? body.data : "";
        const { id, name } = await uploadTakeoffPdf(fileName, contentType, data, env);
        const takeoffUrl = `${url.origin}/takeoff/${id}`;
        return jsonResponse({ ok: true, id, url: takeoffUrl, name }, origin);
      }
      const [contact, files, bid] = await Promise.all([
        fetchAccountContact(url.searchParams.get("accountId") || "", env),
        fetchJobFiles(url.searchParams.get("jobId") || "", env),
        fetchJobBid(url.searchParams.get("jobId") || "", env),
      ]);
      return jsonResponse({ ok: true, contact, files, bid }, origin);
    } catch (err) {
      return jsonResponse({ ok: false, error: err.message || String(err) }, origin);
    }
  },
};
