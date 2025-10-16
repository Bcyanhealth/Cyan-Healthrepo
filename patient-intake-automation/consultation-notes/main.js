/**
 * === PATIENT-ID ROUTING (Per-Form Version) ===
 * Attach this script to a single form's Responses Sheet.
 * You only change:
 *   1) ROOT_FOLDER_ID  → your shared patients root
 *   2) DEST_SUBFOLDER  → where this form's files go inside each patient folder
 *   3) Field titles    → if your form uses different column names
 *
 * Folder structure per submission:
 *   <ROOT>/<PatientID>__<FirstName>/<DEST_SUBFOLDER>/<YYYY-MM-DD_HHMMSS>/{responses.json, summary Doc}
 */

/** 1) REQUIRED: ID of the shared top-level folder that holds all patients */
const ROOT_FOLDER_ID = '1K0KevC96ex7wjxEQcp7aBpUM-XxCxNGa';

/** 2) REQUIRED: this form's destination inside each patient folder
 *    Examples per form:
 *    - Registration:         'Details'
 *    - Consultation Notes:   'Consultation Notes'
 *    - Payments:             'Payments'
 *    - Telemedicine Payment: 'Telemedicine Payments'
 */
const DEST_SUBFOLDER = 'Consultations-notes'; // <- set per form

/** 3) Match your form's exact field titles (headers in the responses Sheet) */
const PATIENT_ID_FIELDS = ['Patient ID'];  // unique key (you said your form has "Patient ID")
const FIRST_NAME_FIELDS = ['First Name'];  // readable label (you said your form has "First Name")


/** Sheet column headers to write back */
const COL_PATIENT_FOLDER_LINK = 'Patient Folder';
const COL_SAVED_FILE_LINK     = 'Saved File Link';

const TZ = Session.getScriptTimeZone() || 'Africa/Nairobi';

/* ===================== MAIN ===================== */
function onFormSubmit(e) {
  if (!e || !e.namedValues || !e.range) {
    throw new Error('Missing event data; ensure trigger is From spreadsheet → On form submit.');
  }

  const nv = e.namedValues;

  // 1) Identify Patient ID (required)
  const patientId = getFirstMatch(nv, PATIENT_ID_FIELDS);
  if (!patientId) {
    tryWriteErrorNote(e, 'Missing Patient ID');
    throw new Error('Missing Patient ID in submission.');
  }

  // 2) Get first name (or split from full name), for nicer folder naming
  const firstName = getFirstMatch(nv, FIRST_NAME_FIELDS)
                 || firstToken(getFirstMatch(nv, FULL_NAME_FIELDS))
                 || 'Unknown';

  // 3) Resolve/create master patient folder under ROOT
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const patientFolderName = `${toSafe(patientId)}__${toSafe(firstName)}`;
  const patientFolder = getOrCreateSubfolder(root, patientFolderName);

  // 4) Ensure this form's destination subfolder exists
  const destFolder = getOrCreateSubfolder(patientFolder, DEST_SUBFOLDER);

  // 5) Create a timestamped folder for this submission
  const stamp = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'_'HHmmss");
  const runFolder = destFolder.createFolder(stamp);



  // 7) Create a summary Doc
  const docTitle = `${DEST_SUBFOLDER.toUpperCase()}__${patientId}__${stamp}`;
  const docId = createSummaryDoc(docTitle, nv, runFolder);

  // 8) Write links back to the response row
  const savedFileUrl = DriveApp.getFileById(docId).getUrl();
  writeBackLinks(e, patientFolder.getUrl(), savedFileUrl);

  // Log for debugging
  console.log(JSON.stringify({
    destSubfolder: DEST_SUBFOLDER,
    patientId,
    patientFolder: patientFolder.getUrl(),
    savedFile: savedFileUrl
  }));
}

/* ===================== HELPERS ===================== */

function getFirstMatch(namedValues, keys) {
  for (const k of keys) {
    const v = pickValue(namedValues[k]);
    if (v) return v;
  }
  return '';
}

function firstToken(s) {
  if (!s) return '';
  const t = String(s).trim().split(/\s+/);
  return t[0] || '';
}

function pickValue(maybeArray) {
  if (Array.isArray(maybeArray) && maybeArray.length) {
    const s = String(maybeArray[0]).trim();
    return s || '';
  }
  if (typeof maybeArray === 'string') return maybeArray.trim();
  return '';
}

function toSafe(s) {
  return String(s).replace(/[<>:"/\\|?*\u0000-\u001F]/g, '_').slice(0, 120);
}

function getOrCreateSubfolder(parent, name) {
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

function createSummaryDoc(title, namedValues, destFolder) {
  const doc = DocumentApp.create(title);
  const file = DriveApp.getFileById(doc.getId());
  file.moveTo(destFolder);

  const body = doc.getBody();
  body.clear();
  body.appendParagraph('Submission Summary').setHeading(DocumentApp.ParagraphHeading.HEADING1);

  const table = [['Field', 'Value']];
  Object.keys(namedValues).forEach((q) => {
    const a = namedValues[q];
    table.push([q, Array.isArray(a) ? a.join(', ') : String(a || '')]);
  });
  body.appendTable(table);

  doc.saveAndClose();
  return doc.getId();
}

function writeBackLinks(e, patientFolderUrl, savedFileUrl) {
  const sheet = e.range.getSheet();
  const headerRow = 1;
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];

  let colPatient = headers.indexOf(COL_PATIENT_FOLDER_LINK) + 1;
  let colSaved   = headers.indexOf(COL_SAVED_FILE_LINK) + 1;

  let newCol = lastCol;
  if (colPatient === 0) {
    newCol += 1;
    sheet.getRange(headerRow, newCol).setValue(COL_PATIENT_FOLDER_LINK);
    colPatient = newCol;
  }
  if (colSaved === 0) {
    newCol += 1;
    sheet.getRange(headerRow, newCol).setValue(COL_SAVED_FILE_LINK);
    colSaved = newCol;
  }

  const row = e.range.getRow();
  sheet.getRange(row, colPatient).setValue(patientFolderUrl);
  sheet.getRange(row, colSaved).setValue(savedFileUrl);
}

function tryWriteErrorNote(e, message) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    const row = e.range.getRow();
    sheet.getRange(row, 1).setNote('Automation error: ' + message);
  } catch (_) {}
}

/* ================ ONE-TIME HELPER ================ */
/** Run this once (from the editor opened on the responses Sheet) if you prefer programmatic trigger creation */
function createTrigger() {
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('Open the responses Google Sheet, then run createTrigger().');
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
}
