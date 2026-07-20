function doGet(event) {
  var params = event && event.parameter ? event.parameter : {};

  if (params.action === 'saveSource') {
    return saveSourceConfig(params);
  }

  var sourceConfig = getSourceConfig();
  var spreadsheetId = sourceConfig.spreadsheetId;
  var spreadsheetUrl = sourceConfig.spreadsheetUrl;
  var sheetName = sourceConfig.sheetName;
  var pendingSheetName = sourceConfig.pendingSheetName;
  var paidSheetName = sourceConfig.paidSheetName;
  var rangeA1 = sourceConfig.range;
  var pendingRangeA1 = sourceConfig.pendingRange;
  var paidDetailRangeA1 = sourceConfig.paidDetailRange;
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);
  var pendingSheet = spreadsheet.getSheetByName(pendingSheetName);
  var paidSheet = spreadsheet.getSheetByName(paidSheetName);
  var values = sheet.getRange(rangeA1).getDisplayValues();
  var pendingValues = pendingSheet ? pendingSheet.getRange(pendingRangeA1).getDisplayValues() : [];
  var paidDetailRawValues = paidSheet ? paidSheet.getRange(paidDetailRangeA1).getDisplayValues() : [];
  var paidDetailHeader = paidDetailRawValues.length > 1 ? sheetRowToPaidDetail(paidDetailRawValues[1]) : [];
  var paidDetailValues = sheetRowsToPaidDetails(paidDetailRawValues);
  var snapshotDate = sheet.getRange('G1').getDisplayValue();

  if (!values || values.length < 2) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'No data found' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var header = values[0];
  var bodyRows = values.slice(1);
  var grandTotalRow = bodyRows.filter(function(row) {
    return row[0] === 'Grand Total';
  })[0] || [];
  var dataRows = bodyRows.filter(function(row) {
    return row[0] && row[0] !== 'Grand Total';
  });
  var pendingHeader = pendingValues[0] || [];
  var pendingRows = pendingValues.slice(1).filter(function(row) {
    return row.some(function(value) {
      return value !== '';
    }) && row[0] !== 'Grand Total';
  });

  var payload = {
    source: {
      spreadsheetId: spreadsheetId,
      spreadsheetUrl: spreadsheetUrl,
      sheetName: sheetName,
      pendingSheetName: pendingSheetName,
      paidSheetName: paidSheetName,
      range: rangeA1,
      pendingRange: pendingRangeA1,
      paidDetailRange: paidDetailRangeA1,
      snapshotDate: snapshotDate
    },
    header: header,
    rows: dataRows,
    pendingHeader: pendingHeader,
    pendingRows: pendingRows,
    paidDetailHeader: paidDetailHeader,
    paidDetailRows: paidDetailValues,
    grandTotal: grandTotalRow[2] || '',
    actualGrandTotal: grandTotalRow[4] || ''
  };

  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSourceConfig() {
  var defaults = getDefaultSourceConfig();
  var saved = PropertiesService.getScriptProperties().getProperty('sourceConfig');

  if (!saved) {
    return defaults;
  }

  try {
    var parsed = JSON.parse(saved);
    return Object.assign({}, defaults, parsed);
  } catch (error) {
    return defaults;
  }
}

function getDefaultSourceConfig() {
  var spreadsheetId = '1fJ6qvATbXqbveDBlqjvVE9Lz5jsRwdW_pXxdYdKheiY';

  return {
    spreadsheetId: spreadsheetId,
    spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/edit',
    sheetName: 'Summary',
    pendingSheetName: 'เตรียมจ่าย',
    paidSheetName: 'จ่ายแล้ว',
    range: 'A:F',
    pendingRange: 'A:G',
    paidDetailRange: 'G:V'
  };
}

function saveSourceConfig(params) {
  var defaults = getDefaultSourceConfig();
  var spreadsheetUrl = params.spreadsheetUrl || defaults.spreadsheetUrl;
  var spreadsheetId = params.spreadsheetId || extractSpreadsheetId(spreadsheetUrl) || defaults.spreadsheetId;
  var sourceConfig = {
    spreadsheetId: spreadsheetId,
    spreadsheetUrl: spreadsheetUrl,
    sheetName: params.sheetName || defaults.sheetName,
    pendingSheetName: params.pendingSheetName || defaults.pendingSheetName,
    paidSheetName: params.paidSheetName || defaults.paidSheetName,
    range: params.range || defaults.range,
    pendingRange: params.pendingRange || defaults.pendingRange,
    paidDetailRange: params.paidDetailRange || defaults.paidDetailRange
  };

  PropertiesService.getScriptProperties().setProperty('sourceConfig', JSON.stringify(sourceConfig));

  return outputJsonp(params.callback, {
    ok: true,
    source: sourceConfig
  });
}

function sheetRowsToPaidDetails(values) {
  return values
    .slice(2)
    .filter(function(row) {
      return row.some(function(value) {
        return value !== '';
      });
    })
    .map(function(row) {
      return [
        row[14] || '',
        row[15] || '',
        row[0] || '',
        row[1] || '',
        row[2] || ''
      ];
    });
}

function sheetRowToPaidDetail(row) {
  return [
    row[14] || '',
    row[15] || '',
    row[0] || '',
    row[1] || '',
    row[2] || ''
  ];
}

function extractSpreadsheetId(url) {
  var match = String(url || '').match(/\/spreadsheets\/d\/([^/]+)/);
  return match ? match[1] : '';
}

function outputJsonp(callbackName, payload) {
  var json = JSON.stringify(payload);

  if (callbackName) {
    return ContentService
      .createTextOutput(callbackName + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
