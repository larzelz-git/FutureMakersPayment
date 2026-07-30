function writeTestSuccess(params) {
  var sourceConfig = Object.assign({}, getSourceConfig(), getRequestSourceConfig(params));
  var spreadsheet = SpreadsheetApp.openById(sourceConfig.spreadsheetId);
  var sheet = spreadsheet.getSheetByName('Test');

  if (!sheet) {
    return outputJsonp(params.callback, { ok: false, error: 'ไม่พบแท็บ Test' });
  }

  sheet.getRange('A1').setValue('Success');

  return outputJsonp(params.callback, {
    ok: true,
    sheetName: 'Test',
    cell: 'A1',
    value: 'Success'
  });
}
