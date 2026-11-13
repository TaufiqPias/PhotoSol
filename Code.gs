// Google Apps Script backend. Paste into script.google.com, then Deploy > Web app (Execute as: Me, Access: Anyone)
const ADMIN_KEY = "6039"; // secret password for creating vaults
const NOTIFY_EMAIL = "taufiq.xenox@gmail.com"; // leave '' to email your own Google account
const ROOT_NAME = "PhotoSol Vault"; // parent folder in your Drive

function root_() {
  const i = DriveApp.getFoldersByName(ROOT_NAME);
  return i.hasNext() ? i.next() : DriveApp.createFolder(ROOT_NAME);
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(
    ContentService.MimeType.JSON,
  );
}
function order_(id) {
  const f = DriveApp.getFolderById(id),
    i = f.getFilesByName("order.json");
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
function doGet(e) {
  try {
    return out_({ ok: true, order: order_(e.parameter.id).data });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}
function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.action === "create") {
      if (b.key !== ADMIN_KEY) throw new Error("Wrong admin key");
      const o = b.order,
        name = (o.wo + " - " + o.address).replace(/[\\\/:*?"<>|]/g, "-");
      const f = root_().createFolder(name);
      f.createFile("order.json", JSON.stringify(o), "application/json");
      return out_({ ok: true, id: f.getId(), name: name, url: f.getUrl() });
    }
    const v = order_(b.id),
      f = v.f;
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
