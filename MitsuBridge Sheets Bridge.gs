const ALLOWED_SPREADSHEET_IDS = new Set([
  "1fmcxmeOpeYgqfs92djrLSKdOuAe98PMMy5HRLMUkwp4",
  "1nHfD5zLLbDbys-Ei_ut2kP1oRf0xb0hYl6hB_ZwzQ0M",
  "1X4ZMrs29baR5zU-y6MgudLHdENZOCBIOsIE3mANvtDk"
]);

function doGet(e) {
  const callback = String(e?.parameter?.callback || "");
  const spreadsheetId = String(e?.parameter?.id || "");
  if (callback && !/^[A-Za-z_$][\w$]*$/.test(callback)) {
    return respond_({ok: false, error: "Invalid response callback."}, "");
  }
  if (!ALLOWED_SPREADSHEET_IDS.has(spreadsheetId)) {
    return respond_({ok: false, error: "This spreadsheet is not in the connector allowlist."}, callback);
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheets = spreadsheet.getSheets().map(sheet => ({
      name: sheet.getName(),
      rows: sheet.getDataRange().getDisplayValues()
    }));
    return respond_({ok: true, sheets}, callback);
  } catch (error) {
    console.error(`Could not read spreadsheet ${spreadsheetId}: ${error.message}`);
    return respond_({ok: false, error: "The connector could not read this spreadsheet. Check its permissions."}, callback);
  }
}

function respond_(data, callback) {
  const body = JSON.stringify(data);
  const output = callback ? `${callback}(${body});` : body;
  return ContentService.createTextOutput(output)
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}