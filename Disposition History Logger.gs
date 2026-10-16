const HISTORY_SHEET_NAME = "Disposition History";
const STATE_SHEET_NAME = "Disposition State";
const HISTORY_HEADERS = [
  "Changed At", "Source Sheet", "Source Row", "Student Name", "Mobile No.",
  "Lead Created", "Old Disposition", "New Disposition", "Previous Data", "Edit Type"
];
const STATE_HEADERS = ["Sheet ID", "Source Row", "Disposition", "Previously Present", "Added At"];

function setupDispositionHistory() {
  const ss = SpreadsheetApp.getActive();
  let history = ss.getSheetByName(HISTORY_SHEET_NAME);
  if (!history) history = ss.insertSheet(HISTORY_SHEET_NAME);
  ensureHeaders_(history, HISTORY_HEADERS);

  let state = ss.getSheetByName(STATE_SHEET_NAME);
  if (!state) state = ss.insertSheet(STATE_SHEET_NAME);
  ensureHeaders_(state, STATE_HEADERS);

  const known = readState_(state);
  const baseline = [];
  ss.getSheets().forEach(sheet => {
    if (sheet.getName() === HISTORY_SHEET_NAME || sheet.getName() === STATE_SHEET_NAME) return;
    const header = findHeaders_(sheet);
    if (!header || !header.disposition) return;
    const lastRow = sheet.getLastRow();
    if (lastRow <= header.row) return;
    const rows = sheet.getRange(header.row + 1, 1, lastRow - header.row, sheet.getLastColumn()).getDisplayValues();
    rows.forEach((row, offset) => {
      const sourceRow = header.row + 1 + offset;
      if (!hasLeadData_(row, header)) return;
      const key = `${sheet.getSheetId()}:${sourceRow}`;
      if (known.has(key)) return;
      baseline.push([sheet.getSheetId(), sourceRow, row[header.disposition - 1] || "", true, ""]);
      known.add(key);
    });
  });
  if (baseline.length) state.getRange(state.getLastRow() + 1, 1, baseline.length, STATE_HEADERS.length).setValues(baseline);

  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === "captureDispositionEdit")
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));
  ScriptApp.newTrigger("captureDispositionEdit").forSpreadsheet(ss).onEdit().create();
  return `Disposition history is ready. Baseline rows: ${baseline.length}.`;
}

function captureDispositionEdit(e) {
  if (!e || !e.range) return;
  const ss = e.source;
  const sheet = e.range.getSheet();
  if (sheet.getName() === HISTORY_SHEET_NAME || sheet.getName() === STATE_SHEET_NAME) return;

  const header = findHeaders_(sheet);
  if (!header || !header.disposition) return;
  const edited = e.range;
  const dispositionColumn = header.disposition;
  const dispositionEdited = dispositionColumn >= edited.getColumn() &&
    dispositionColumn <= edited.getLastColumn();

  const history = ss.getSheetByName(HISTORY_SHEET_NAME);
  const state = ss.getSheetByName(STATE_SHEET_NAME);
  if (!history || !state) throw new Error("Run setupDispositionHistory() first.");

  const stateRows = readState_(state);
  const firstRow = Math.max(edited.getRow(), header.row + 1);
  const lastRow = edited.getLastRow();
  if (lastRow < firstRow) return;

  const sheetRows = sheet.getRange(firstRow, 1, lastRow - firstRow + 1, sheet.getLastColumn()).getDisplayValues();
  const stateUpdates = [];
  const historyRows = [];
  const timeZone = ss.getSpreadsheetTimeZone();

  sheetRows.forEach((row, offset) => {
    const sourceRow = firstRow + offset;
    if (!hasLeadData_(row, header)) return;
    const key = `${sheet.getSheetId()}:${sourceRow}`;
    const prior = stateRows.get(key);
    const disposition = row[dispositionColumn - 1] || "";
    const isNewLead = !prior;
    const changedDisposition = dispositionEdited &&
      (!prior || normalize_(prior.disposition) !== normalize_(disposition));
    if (!isNewLead && !changedDisposition) return;

    const leadCreated = header.leadCreated ? row[header.leadCreated - 1] || "" : "";
    const studentName = header.studentName ? row[header.studentName - 1] || "" : "";
    const mobile = header.mobile ? row[header.mobile - 1] || "" : "";
    const changedAt = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd HH:mm:ss");
    const previousData = !!prior && (!prior.addedAt || prior.addedAt.slice(0, 10) < changedAt.slice(0, 10));
    const historyRow = (oldDisposition, newDisposition, wasPreviousData, editType) => [
      changedAt, sheet.getName(), sourceRow, studentName, mobile, leadCreated,
      oldDisposition, newDisposition, wasPreviousData, editType
    ];
    if (isNewLead) historyRows.push(historyRow("", disposition, false, "LEAD_ADDED"));
    if (changedDisposition && (!isNewLead || disposition.trim())) {
      historyRows.push(historyRow(prior ? prior.disposition : "", disposition, previousData, "DISPOSITION"));
    }
    stateUpdates.push({key, sourceRow, disposition, present: true, addedAt: prior ? prior.addedAt : changedAt});
  });

  if (historyRows.length) {
    history.getRange(history.getLastRow() + 1, 1, historyRows.length, HISTORY_HEADERS.length).setValues(historyRows);
  }
  stateUpdates.forEach(update => {
    const existing = stateRows.get(update.key);
    const values = [[sheet.getSheetId(), update.sourceRow, update.disposition, update.present, update.addedAt]];
    if (existing) {
      state.getRange(existing.row, 1, 1, STATE_HEADERS.length).setValues(values);
    } else {
      const rowNumber = state.getLastRow() + 1;
      state.getRange(rowNumber, 1, 1, STATE_HEADERS.length).setValues(values);
      stateRows.set(update.key, {row: rowNumber, disposition: update.disposition, present: true, addedAt: update.addedAt});
    }
  });
}

function findHeaders_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (!lastColumn) return null;
  const rows = sheet.getRange(1, 1, Math.min(15, Math.max(sheet.getLastRow(), 1)), lastColumn).getDisplayValues();
  for (let r = 0; r < rows.length; r++) {
    const headers = rows[r].map(value => normalize_(value));
    const disposition = headers.findIndex(value => value === "disposition");
    const studentName = headers.findIndex(value => /^(studentname|fullname|name)$/.test(value));
    const mobile = headers.findIndex(value => /mobile|phone|contact/.test(value));
    if (disposition >= 0 && (studentName >= 0 || mobile >= 0)) {
      const leadCreated = headers.findIndex(value => /createdtime|createddate|leaddate|enquirydate/.test(value));
      return {
        row: r + 1,
        disposition: disposition + 1,
        studentName: studentName + 1,
        mobile: mobile + 1,
        leadCreated: leadCreated + 1
      };
    }
  }
  return null;
}

function hasLeadData_(row, header) {
  return [header.studentName, header.mobile, header.leadCreated]
    .filter(Boolean)
    .some(column => String(row[column - 1] || "").trim() !== "");
}

function ensureHeaders_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return;
  }
  const current = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
  if (headers.some((header, index) => current[index] !== header)) {
    throw new Error(`Unexpected headers in "${sheet.getName()}". Do not rename the logger columns.`);
  }
}

function readState_(sheet) {
  const result = new Map();
  if (sheet.getLastRow() < 2) return result;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, STATE_HEADERS.length).getValues();
  rows.forEach((row, index) => {
    const key = `${row[0]}:${row[1]}`;
    result.set(key, {row: index + 2, disposition: String(row[2] || ""), present: row[3] === true, addedAt: String(row[4] || "")});
  });
  return result;
}

function normalize_(value) {
  return String(value || "").replace(/[\u200B-\u200D\uFEFF]/g, "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const callback = String(params.callback || "");
  if (callback && !/^[A-Za-z_$][A-Za-z0-9_$]{0,80}$/.test(callback)) {
    return ContentService.createTextOutput("Invalid callback.")
      .setMimeType(ContentService.MimeType.TEXT);
  }

  let result;
  try {
    const id = String(params.id || "");
    if (!/^[A-Za-z0-9_-]{20,}$/.test(id)) throw new Error("Invalid spreadsheet ID.");
    const file = DriveApp.getFileById(id);
    if (file.getSharingAccess() !== DriveApp.Access.ANYONE_WITH_LINK ||
        file.getSharingPermission() !== DriveApp.Permission.VIEW) {
      throw new Error("Share this sheet as Anyone with the link — Viewer.");
    }
    const spreadsheet = SpreadsheetApp.openById(id);
    result = {
      ok: true,
      sheets: spreadsheet.getSheets().map(sheet => ({
        name: sheet.getName(),
        rows: sheet.getDataRange().getDisplayValues()
      }))
    };
  } catch (error) {
    result = {ok: false, error: error && error.message ? error.message : "Could not read this spreadsheet."};
  }

  const body = JSON.stringify(result);
  const output = callback ? `${callback}(${body});` : body;
  return ContentService.createTextOutput(output)
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}
