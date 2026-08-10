const SHEET_NAME = "data";

function getDataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange("A1").setValue("{}");
  }
  return sheet;
}

// ===================================================================================
// JobTread integration (secure proxy)
//
// ONE-TIME SETUP (do this in the Apps Script editor — never put real secrets in this file):
//   1. Project Settings (gear icon, left sidebar) > Script Properties > Add script property:
//        JOBTREAD_GRANT_KEY       = <JobTread API grant key>
//        JOBTREAD_ORGANIZATION_ID = <JobTread organization id>   (optional; auto-resolved if omitted)
//        JT_PROXY_SECRET          = <any random string you choose, e.g. from a password generator>
//   2. Deploy > New deployment > select type "Web app" > Execute as: Me, Who has access: Anyone.
//   3. Copy the resulting /exec URL and paste it into index.html's JOBTREAD_PROXY_URL constant,
//      and paste the JT_PROXY_SECRET value into index.html's JOBTREAD_PROXY_KEY constant.
//
// The real JobTread grant key never leaves this script — the client only holds JT_PROXY_SECRET,
// which just gates access to this proxy's two read-only endpoints (same pattern as ARTracker's
// existing schedule-tracker write proxy: a scoped secret client-side, the real credential
// server-side only).
// ===================================================================================
const JOBTREAD_PAVE_URL = "https://api.jobtread.com/pave";

function jtProps_() {
  return PropertiesService.getScriptProperties();
}

function jtCheckSecret_(e) {
  const expected = jtProps_().getProperty("JT_PROXY_SECRET");
  return !!expected && e.parameter && e.parameter.key === expected;
}

function jtQuery_(query) {
  const grantKey = jtProps_().getProperty("JOBTREAD_GRANT_KEY");
  if (!grantKey) throw new Error("JOBTREAD_GRANT_KEY script property is not set");
  const res = UrlFetchApp.fetch(JOBTREAD_PAVE_URL, {
    method: "post",
    contentType: "application/json",
    muteHttpExceptions: true,
    payload: JSON.stringify({ query: Object.assign({ $: { grantKey: grantKey, notify: false } }, query) }),
  });
  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error("JobTread request failed (" + code + "): " + text.slice(0, 300));
  }
  const payload = JSON.parse(text || "{}");
  if (payload.errors && payload.errors.length) {
    const first = payload.errors[0];
    throw new Error("JobTread query error: " + (first.message || JSON.stringify(first)));
  }
  return payload;
}

function jtGetByPath_(obj, path) {
  return path.split(".").reduce(function (acc, key) {
    return acc && typeof acc === "object" ? acc[key] : undefined;
  }, obj);
}

function jtResolveOrganizationId_() {
  const cached = jtProps_().getProperty("JOBTREAD_ORGANIZATION_ID");
  if (cached) return cached;
  const payload = jtQuery_({
    currentGrant: { user: { memberships: { nodes: { id: {}, organization: { id: {}, name: {} } } } } },
  });
  const nodes = jtGetByPath_(payload, "currentGrant.user.memberships.nodes") || [];
  const first = nodes[0];
  if (!first) throw new Error("Unable to resolve JobTread organization from the grant");
  return jtGetByPath_(first, "organization.id");
}

// Searches recent JobTread documents and returns the distinct jobs whose name/account/address
// match the search text. JobTread's Pave schema doesn't expose organization.jobs directly, so
// this mirrors ARTracker's proven approach of deriving jobs from the document list instead.
function jtSearchJobs_(searchText) {
  const organizationId = jtResolveOrganizationId_();
  const needle = (searchText || "").trim().toLowerCase();
  const payload = jtQuery_({
    organization: {
      $: { id: organizationId },
      documents: {
        $: { size: 100, sortBy: [{ field: "createdAt", order: "desc" }] },
        nodes: {
          id: {},
          job: {
            id: {},
            name: {},
            number: {},
            location: {
              id: {},
              name: {},
              address: { formatted: {} },
              account: { id: {}, name: {} },
            },
          },
        },
      },
    },
  });
  const nodes = jtGetByPath_(payload, "organization.documents.nodes") || [];
  const seen = {};
  const results = [];
  nodes.forEach(function (node) {
    const jobId = jtGetByPath_(node, "job.id");
    if (!jobId || seen[jobId]) return;
    const jobName = jtGetByPath_(node, "job.name") || "";
    const accountName = jtGetByPath_(node, "job.location.account.name") || "";
    const address = jtGetByPath_(node, "job.location.address.formatted") || "";
    const haystack = (jobName + " " + accountName + " " + address).toLowerCase();
    if (needle && haystack.indexOf(needle) === -1) return;
    seen[jobId] = true;
    results.push({
      jobId: jobId,
      accountId: jtGetByPath_(node, "job.location.account.id") || "",
      name: jobName,
      accountName: accountName,
      address: address,
    });
  });
  return results.slice(0, 20);
}

// Best-effort contact lookup for a job's account. JobTread's exact contact schema varies by
// tenant, so several field shapes are probed in order and the first one with data wins (same
// defensive pattern ARTracker uses for the same problem).
function jtFetchAccountContact_(accountId) {
  if (!accountId) return { name: "", email: "", phone: "" };
  const variants = [
    { account: { $: { id: accountId }, contacts: { nodes: { id: {}, name: {}, emailAddress: {}, phoneNumber: {} } } } },
    { account: { $: { id: accountId }, aces: { nodes: { membership: { accountType: {}, contact: { name: {}, emailAddress: {}, phoneNumber: {} }, user: { name: {}, emailAddress: {}, phoneNumber: {} } } } } } },
    { account: { $: { id: accountId }, memberships: { nodes: { contact: { name: {}, emailAddress: {}, phoneNumber: {} }, user: { name: {}, emailAddress: {}, phoneNumber: {} } } } } },
  ];
  for (var i = 0; i < variants.length; i++) {
    try {
      const payload = jtQuery_(variants[i]);
      const directContacts = jtGetByPath_(payload, "account.contacts.nodes") || [];
      if (directContacts.length) {
        const c = directContacts[0];
        return { name: c.name || "", email: c.emailAddress || "", phone: c.phoneNumber || "" };
      }
      const aces = jtGetByPath_(payload, "account.aces.nodes") || [];
      for (var j = 0; j < aces.length; j++) {
        const m = aces[j].membership || {};
        const name = (m.contact && m.contact.name) || (m.user && m.user.name);
        if (name) {
          return {
            name: name,
            email: (m.contact && m.contact.emailAddress) || (m.user && m.user.emailAddress) || "",
            phone: (m.contact && m.contact.phoneNumber) || (m.user && m.user.phoneNumber) || "",
          };
        }
      }
      const memberships = jtGetByPath_(payload, "account.memberships.nodes") || [];
      for (var k = 0; k < memberships.length; k++) {
        const mem = memberships[k];
        const name2 = (mem.contact && mem.contact.name) || (mem.user && mem.user.name);
        if (name2) {
          return {
            name: name2,
            email: (mem.contact && mem.contact.emailAddress) || (mem.user && mem.user.emailAddress) || "",
            phone: (mem.contact && mem.contact.phoneNumber) || (mem.user && mem.user.phoneNumber) || "",
          };
        }
      }
    } catch (err) {
      // Ignore unsupported field shapes and try the next probe variant.
    }
  }
  return { name: "", email: "", phone: "" };
}

// Best-effort PDF take-off lookup for a job. JobTread's file/attachment schema hasn't been
// confirmed against the live tenant yet, so several likely shapes are probed; results are
// filtered down to PDF-like files, with anything named "takeoff"/"take-off" sorted first.
// NOTE: verify this against a real job once JOBTREAD_GRANT_KEY is configured, and adjust the
// field names below if the JobTread schema differs (see jtFetchJobFiles_ variants).
function jtFetchJobFiles_(jobId) {
  if (!jobId) return [];
  const variants = [
    { job: { $: { id: jobId }, files: { nodes: { id: {}, name: {}, url: {}, contentType: {} } } } },
    { job: { $: { id: jobId }, attachments: { nodes: { id: {}, name: {}, url: {}, contentType: {} } } } },
    { job: { $: { id: jobId }, documents: { nodes: { id: {}, name: {}, type: {}, url: {} } } } },
  ];
  var found = [];
  for (var i = 0; i < variants.length; i++) {
    try {
      const payload = jtQuery_(variants[i]);
      const nodes =
        jtGetByPath_(payload, "job.files.nodes") ||
        jtGetByPath_(payload, "job.attachments.nodes") ||
        jtGetByPath_(payload, "job.documents.nodes") ||
        [];
      if (nodes.length) {
        found = nodes;
        break;
      }
    } catch (err) {
      // Ignore unsupported field shapes and try the next probe variant.
    }
  }
  return found
    .map(function (n) {
      return { id: n.id || "", name: n.name || "", url: n.url || "", contentType: n.contentType || n.type || "" };
    })
    .filter(function (f) {
      const n = f.name.toLowerCase();
      return (
        n.indexOf(".pdf") !== -1 ||
        (f.contentType || "").toLowerCase().indexOf("pdf") !== -1 ||
        n.indexOf("takeoff") !== -1 ||
        n.indexOf("take-off") !== -1
      );
    })
    .sort(function (a, b) {
      function score(f) {
        const n = f.name.toLowerCase();
        return n.indexOf("takeoff") !== -1 || n.indexOf("take-off") !== -1 ? 0 : 1;
      }
      return score(a) - score(b);
    });
}

function jtJsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Handles all requests via GET.
// ?action=save&d=<urlencoded json>            → saves data (legacy sheet-backed store)
// ?action=jtSearch&q=<text>&key=<secret>       → search JobTread jobs by name/account/address
// ?action=jtJobDetail&jobId=&accountId=&key=   → fetch contact + candidate PDF take-off files
// no action param                              → returns current data (legacy sheet-backed store)
function doGet(e) {
  const action = e.parameter && e.parameter.action;

  if (action === "jtSearch" || action === "jtJobDetail") {
    if (!jtCheckSecret_(e)) {
      return jtJsonResponse_({ ok: false, error: "Unauthorized" });
    }
    try {
      if (action === "jtSearch") {
        return jtJsonResponse_({ ok: true, jobs: jtSearchJobs_(e.parameter.q || "") });
      }
      return jtJsonResponse_({
        ok: true,
        contact: jtFetchAccountContact_(e.parameter.accountId || ""),
        files: jtFetchJobFiles_(e.parameter.jobId || ""),
      });
    } catch (err) {
      return jtJsonResponse_({ ok: false, error: err.message || String(err) });
    }
  }

  const sheet = getDataSheet_();

  if (e.parameter && e.parameter.action === "save") {
    try {
      const body = JSON.parse(decodeURIComponent(e.parameter.d));
      sheet.getRange("A1").setValue(JSON.stringify(body));
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Default: load
  const raw = sheet.getRange("A1").getValue();
  const json = raw && raw.toString().trim() ? raw.toString() : "{}";
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// Keep doPost as a no-op fallback so old requests don't error
function doPost(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: "Use GET with action=save" }))
    .setMimeType(ContentService.MimeType.JSON);
}
