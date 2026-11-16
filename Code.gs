// Google Apps Script backend. Paste into script.google.com, then Deploy > Web app (Execute as: Me, Access: Anyone)
// Each admin key belongs to one user. Every user only sees and manages their own vaults.
// Format:  "secret key": "User name".  Two keys with the same name share the same vaults.
const ADMIN_KEYS = {
  "6039": "Taufiq",
  // "4821": "Rahim",
  // "7315": "Sara",
};
const LEGACY_OWNER = "Taufiq"; // vaults created before multi-user support belong to this user
const NOTIFY_EMAIL = "taufiq.xenox@gmail.com"; // leave '' to email your own Google account
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
function user_(key) {
  key = String(key || "");
  if (!Object.prototype.hasOwnProperty.call(ADMIN_KEYS, key)) throw new Error("Wrong admin key");
  return ADMIN_KEYS[key];
}
function ownerOf_(f) {
  return f.getDescription() || LEGACY_OWNER; // the owner's name is kept in the folder description
}
function own_(v, user) {
  if (ownerOf_(v.f) !== user) throw new Error("This vault belongs to another user");
}
function list_(user) {
  const it = root_().getFolders(),
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
    res.push({ id: f.getId(), name: f.getName(), url: f.getUrl(), created: c });
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
    if (now - f.getDateCreated().getTime() >= VAULT_DAYS * DAY_MS) f.setTrashed(true);
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
    return out_({ ok: true, order: v.data, created: v.f.getDateCreated().getTime() });
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
    if (b.action === "upload") {
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
      const it = f.getFiles();
      let c = 0;
      while (it.hasNext()) {
        if (it.next().getName() !== "order.json") c++;
      }
      MailApp.sendEmail(
        NOTIFY_EMAIL || Session.getEffectiveUser().getEmail(),
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
