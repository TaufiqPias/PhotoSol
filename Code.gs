// Google Apps Script backend. Paste into script.google.com, then Deploy > Web app (Execute as: Me, Access: Anyone)
// Each admin key belongs to one user. Every user only sees and manages their own vaults.
// Format:  "secret key": { name: "User name", email: "where 'Order is complete' emails go" }
// Two keys with the same name share the same vaults (the first email found for that name is used).
const ADMIN_KEYS = {
  6039: { name: "Taufiq", email: "taufiq.xenox@gmail.com" },
  4815: {
    name: "Kingson Transport LLC",
    email: "robertp@kingsontransport.com",
  },
  1234: { name: "Demo Company LLC", email: "rahim@example.com" },
  // "7315": { name: "Sara", email: "sara@example.com" },
};
const LEGACY_OWNER = "Taufiq"; // vaults created before multi-user support belong to this user
const NOTIFY_EMAIL = ""; // fallback if a user has no email above. Leave '' to use your own Google account
const SIGNUP_EMAIL = "taufiq.xenox@gmail.com"; // sign-up requests are emailed here
const ROOT_NAME = "PhotoSol Vault"; // parent folder in your Drive
const VAULT_DAYS = 30; // vaults are moved to the Drive Trash this many days after creation
const DAY_MS = 24 * 60 * 60 * 1000;

function root_() {
  const i = DriveApp.getFoldersByName(ROOT_NAME);
  return i.hasNext() ? i.next() : DriveApp.createFolder(ROOT_NAME);
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(
    ContentService.MimeType.JSON,
  );
}
function entry_(v) {
  return typeof v === "string" ? { name: v, email: "" } : v; // still accepts the old "key": "Name" format
}
function user_(key) {
  key = String(key || "");
  if (!Object.prototype.hasOwnProperty.call(ADMIN_KEYS, key))
    throw new Error("Wrong admin key");
  return entry_(ADMIN_KEYS[key]).name;
}
function emailOf_(name) {
  for (const k in ADMIN_KEYS) {
    const e = entry_(ADMIN_KEYS[k]);
    if (e.name === name && e.email) return e.email;
  }
  return NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
}
// ---- submitted flag and photo counter live in Script Properties (fast, no Drive scan) ----
function props_() {
  return PropertiesService.getScriptProperties().getProperties();
}
function isSub_(fid) {
  return !!PropertiesService.getScriptProperties().getProperty("s_" + fid);
}
function ownerOf_(f) {
  return f.getDescription() || LEGACY_OWNER; // the owner's name is kept in the folder description
}
function own_(v, user) {
  if (ownerOf_(v.f) !== user)
    throw new Error("This vault belongs to another user");
}
function list_(user) {
  const it = root_().getFolders(),
    P = props_(),
    now = Date.now(),
    res = [],
    old = [];
  while (it.hasNext()) {
    const f = it.next(),
      c = f.getDateCreated().getTime();
    if (now - c >= VAULT_DAYS * DAY_MS) {
      old.push(f);
      continue;
    }
    if (ownerOf_(f) !== user) continue;
    const id = f.getId();
    res.push({
      id: id,
      name: f.getName(),
      url: f.getUrl(),
      created: c,
      photos: Number(P["n_" + id]) || 0,
      submitted: !!P["s_" + id],
    });
  }
  old.forEach(function (f) {
    f.setTrashed(true);
  });
  res.sort(function (a, b) {
    return b.created - a.created;
  });
  return res;
}
function order_(id) {
  const f = DriveApp.getFolderById(id);
  // Expired or trashed vaults stop working immediately (even before the hourly cleanup runs)
  if (f.isTrashed()) throw new Error("Vault expired");
  if (Date.now() - f.getDateCreated().getTime() >= VAULT_DAYS * DAY_MS) {
    f.setTrashed(true);
    throw new Error("Vault expired");
  }
  const i = f.getFilesByName("order.json");
  if (!i.hasNext()) throw new Error("Invalid vault");
  return { f: f, data: JSON.parse(i.next().getBlob().getDataAsString()) };
}
function nextName_(fid) {
  // psupload-01, psupload-02 ... per vault
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const P = PropertiesService.getScriptProperties(),
      k = "n_" + fid,
      n = (Number(P.getProperty(k)) || 0) + 1;
    P.setProperty(k, String(n));
    return "psupload-" + String(n).padStart(2, "0") + ".jpg";
  } finally {
    lock.releaseLock();
  }
}

// ---- Auto-delete: moves vault folders older than VAULT_DAYS to the Drive Trash ----
function cleanupExpired() {
  const it = root_().getFolders(),
    now = Date.now();
  while (it.hasNext()) {
    const f = it.next();
    if (now - f.getDateCreated().getTime() >= VAULT_DAYS * DAY_MS)
      f.setTrashed(true);
  }
}
// Run this ONCE manually (select setupTrigger > Run) to install the hourly cleanup trigger
function setupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(function (t) {
      return t.getHandlerFunction() === "cleanupExpired";
    })
    .forEach(function (t) {
      ScriptApp.deleteTrigger(t);
    });
  ScriptApp.newTrigger("cleanupExpired").timeBased().everyHours(1).create();
}

function doGet(e) {
  try {
    const v = order_(e.parameter.id);
    return out_({
      ok: true,
      order: v.data,
      created: v.f.getDateCreated().getTime(),
      submitted: isSub_(v.f.getId()),
    });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}
function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.action === "list") {
      const user = user_(b.key);
      return out_({ ok: true, user: user, vaults: list_(user) });
    }
    if (b.action === "signup") {
      const f = b.form || {};
      if (f.website) return out_({ ok: true }); // honeypot: bots fill this hidden field
      const g = function (k) {
        return String(f[k] || "")
          .replace(/[\r\n]+/g, " ")
          .trim()
          .slice(0, 300);
      };
      const d = {
        name: g("name"),
        email: g("email"),
        cemail: g("cemail"),
        company: g("company"),
        phone: g("phone"),
        whatsapp: g("whatsapp"),
        address: String(f.address || "")
          .trim()
          .slice(0, 600),
      };
      for (const k in d)
        if (!d[k]) throw new Error("Please fill in all fields.");
      const mail = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
      if (!mail.test(d.email) || !mail.test(d.cemail))
        throw new Error("Please enter valid email addresses.");
      // basic spam protection: one request per email per 10 min, 30 requests per hour overall
      const C = CacheService.getScriptCache(),
        ek = "su_" + d.email.toLowerCase().slice(0, 200),
        n = Number(C.get("su_n")) || 0;
      if (C.get(ek))
        throw new Error(
          "We already received your request. Please wait for our email.",
        );
      if (n >= 30)
        throw new Error("Too many requests right now. Please try again later.");
      C.put(ek, "1", 600);
      C.put("su_n", String(n + 1), 3600);
      MailApp.sendEmail({
        to: SIGNUP_EMAIL,
        replyTo: d.email,
        subject: "New PhotoSol sign-up - " + d.company,
        body:
          "A new sign-up request was submitted.\n\n" +
          "Full Name: " +
          d.name +
          "\n" +
          "Personal Email: " +
          d.email +
          "\n" +
          "Company Email for Notification: " +
          d.cemail +
          "\n" +
          "Company Name: " +
          d.company +
          "\n" +
          "Contact Number: " +
          d.phone +
          "\n" +
          "WhatsApp Number: " +
          d.whatsapp +
          "\n" +
          "Address (BD): " +
          d.address +
          "\n\n" +
          "To approve: add a key to ADMIN_KEYS in Code.gs with the name and notification email above, then email the key to the personal email.",
      });
      return out_({ ok: true });
    }
    if (b.action === "status") {
      // Lightweight poll: reads only Script Properties, no Drive calls
      user_(b.key);
      const P = props_(),
        st = {};
      (Array.isArray(b.ids) ? b.ids : []).slice(0, 300).forEach(function (id) {
        st[id] = {
          photos: Number(P["n_" + id]) || 0,
          submitted: !!P["s_" + id],
        };
      });
      return out_({ ok: true, status: st });
    }
    if (b.action === "create") {
      const user = user_(b.key);
      const o = b.order,
        name = (o.wo + " - " + o.address).replace(/[\\\/:*?"<>|]/g, "-");
      const f = root_().createFolder(name);
      f.setDescription(user);
      f.createFile("order.json", JSON.stringify(o), "application/json");
      return out_({
        ok: true,
        id: f.getId(),
        name: name,
        url: f.getUrl(),
        created: f.getDateCreated().getTime(),
      });
    }
    const v = order_(b.id),
      f = v.f;
    if (b.action === "delete") {
      own_(v, user_(b.key));
      f.setTrashed(true); // moves to the Drive Trash (recoverable)
      return out_({ ok: true });
    }
    if (b.action === "update") {
      own_(v, user_(b.key));
      const file = f.getFilesByName("order.json").next(),
        o = v.data;
      // editable fields
      if (b.fields) {
        ["due", "contractor", "lock", "code", "pay"].forEach(function (k) {
          if (k in b.fields) o[k] = String(b.fields[k] || "").trim();
        });
      }
      o.instrs = Array.isArray(b.instrs) ? b.instrs.map(String) : [];
      o.instr = String(b.instr || "");
      file.setContent(JSON.stringify(o));
      return out_({ ok: true });
    }
    if (b.action === "reopen") {
      own_(v, user_(b.key));
      PropertiesService.getScriptProperties().deleteProperty("s_" + f.getId());
      return out_({ ok: true });
    }
    if (b.action === "upload") {
      if (isSub_(f.getId()))
        throw new Error(
          "This order was already submitted. The link is suspended.",
        );
      const name = nextName_(f.getId());
      const file = f.createFile(
        Utilities.newBlob(
          Utilities.base64Decode(b.data),
          b.type || "image/jpeg",
          name,
        ),
      );
      if (b.desc) file.setDescription(String(b.desc).slice(0, 900));
      return out_({ ok: true, name: name });
    }
    if (b.action === "submit") {
      if (isSub_(f.getId())) return out_({ ok: true, already: true });
      PropertiesService.getScriptProperties().setProperty(
        "s_" + f.getId(),
        String(Date.now()),
      );
      const it = f.getFiles();
      let c = 0;
      while (it.hasNext()) {
        if (it.next().getName() !== "order.json") c++;
      }
      MailApp.sendEmail(
        emailOf_(ownerOf_(f)),
        "Order is complete - WO " + v.data.wo,
        "Order is complete.\n\nWork Order: " +
          v.data.wo +
          "\nAddress: " +
          v.data.address +
          "\nPhotos: " +
          c +
          "\nFolder: " +
          f.getUrl(),
      );
      return out_({ ok: true, count: c });
    }
    throw new Error("Unknown action");
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}
