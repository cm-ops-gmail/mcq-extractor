// Google Apps Script: receives a picture from the MCQ Scraper page and saves it in a Drive folder.
// Setup: script.google.com -> New project -> paste this -> set FOLDER_ID -> Deploy -> New deployment
//        -> type "Web app" -> Execute as: Me -> Who has access: Anyone -> Deploy -> copy the /exec URL.
// First time only: run doGet once from the editor (Run button) and approve the Drive permission.

const FOLDER_ID = "1I2TQS_xMsqDiFbCWno1lrNBxYrP9i_XC"; // https://drive.google.com/drive/folders/<this id>

function doGet() {
  return json({ ok: true, folder: DriveApp.getFolderById(FOLDER_ID).getName() });
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    const blob = Utilities.newBlob(Utilities.base64Decode(p.data), p.mime || "image/png", p.name || "image.png");
    const file = DriveApp.getFolderById(FOLDER_ID).createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) { /* folder sharing applies */ }
    return json({ ok: true, id: file.getId(), url: "https://drive.google.com/file/d/" + file.getId() + "/view" });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
