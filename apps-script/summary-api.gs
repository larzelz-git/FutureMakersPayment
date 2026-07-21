function doGet(event) {
  var params = event && event.parameter ? event.parameter : {};

  if (params.action === 'saveSource') {
    return saveSourceConfig(params);
  }

  var sourceConfig = Object.assign({}, getSourceConfig(), getRequestSourceConfig(params));
  var spreadsheetId = sourceConfig.spreadsheetId;
  var spreadsheetUrl = sourceConfig.spreadsheetUrl;
  var sheetName = sourceConfig.sheetName;
  var incomeSheetName = sourceConfig.incomeSheetName;
  var pendingSheetName = sourceConfig.pendingSheetName;
  var paidSheetName = sourceConfig.paidSheetName;
  var rangeA1 = sourceConfig.range;
  var incomeRangeA1 = sourceConfig.incomeRange;
  var pendingRangeA1 = sourceConfig.pendingRange;
  var paidDetailRangeA1 = sourceConfig.paidDetailRange;
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);
  var incomeSheet = spreadsheet.getSheetByName(incomeSheetName);
  var pendingSheet = spreadsheet.getSheetByName(pendingSheetName);
  var paidSheet = spreadsheet.getSheetByName(paidSheetName);
  var values = sheet.getRange(rangeA1).getDisplayValues();
  var incomeValues = incomeSheet ? incomeSheet.getRange(incomeRangeA1).getDisplayValues() : [];
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
      incomeSheetName: incomeSheetName,
      pendingSheetName: pendingSheetName,
      paidSheetName: paidSheetName,
      range: rangeA1,
      incomeRange: incomeRangeA1,
      pendingRange: pendingRangeA1,
      paidDetailRange: paidDetailRangeA1,
      snapshotDate: snapshotDate
    },
    header: header,
    rows: dataRows,
    pendingHeader: pendingHeader,
    pendingRows: pendingRows,
    incomeSummary: normalizeIncomeSummaryValues(incomeValues),
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
    sheetName: 'Summary รายจ่าย',
    incomeSheetName: 'Summary รายรับ',
    pendingSheetName: 'เตรียมจ่าย',
    paidSheetName: 'จ่ายแล้ว',
    range: 'A:F',
    incomeRange: 'A:Z',
    pendingRange: 'A:G',
    paidDetailRange: 'G:V'
  };
}

function getRequestSourceConfig(params) {
  var sourceConfig = {};
  var spreadsheetUrl = params.spreadsheetUrl || params.sourceUrl || params.url || '';
  var spreadsheetId = params.spreadsheetId || params.sourceId || params.id || '';

  if (spreadsheetUrl) {
    sourceConfig.spreadsheetUrl = spreadsheetUrl;
    sourceConfig.spreadsheetId = spreadsheetId || extractSpreadsheetId(spreadsheetUrl);
  } else if (spreadsheetId) {
    sourceConfig.spreadsheetId = spreadsheetId;
  }

  [
    'sheetName',
    'incomeSheetName',
    'pendingSheetName',
    'paidSheetName',
    'range',
    'incomeRange',
    'pendingRange',
    'paidDetailRange'
  ].forEach(function(key) {
    if (params[key]) {
      sourceConfig[key] = params[key];
    }
  });

  return sourceConfig;
}

function saveSourceConfig(params) {
  var defaults = getDefaultSourceConfig();
  var spreadsheetUrl = params.spreadsheetUrl || defaults.spreadsheetUrl;
  var spreadsheetId = params.spreadsheetId || extractSpreadsheetId(spreadsheetUrl) || defaults.spreadsheetId;
  var sourceConfig = {
    spreadsheetId: spreadsheetId,
    spreadsheetUrl: spreadsheetUrl,
    sheetName: params.sheetName || defaults.sheetName,
    incomeSheetName: params.incomeSheetName || defaults.incomeSheetName,
    pendingSheetName: params.pendingSheetName || defaults.pendingSheetName,
    paidSheetName: params.paidSheetName || defaults.paidSheetName,
    range: params.range || defaults.range,
    incomeRange: params.incomeRange || defaults.incomeRange,
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

function parseAmount(value) {
  if (value === null || value === undefined || value === '-') {
    return 0;
  }

  return Number(String(value).replace(/,/g, '')) || 0;
}

function findAmountNearLabel(row, labelIndex) {
  var index;
  var amount;

  for (index = labelIndex + 1; index < row.length; index += 1) {
    amount = parseAmount(row[index]);

    if (amount !== 0) {
      return amount;
    }
  }

  for (index = 0; index < row.length; index += 1) {
    if (index !== labelIndex) {
      amount = parseAmount(row[index]);

      if (amount !== 0) {
        return amount;
      }
    }
  }

  return 0;
}

function normalizeIncomeSummaryValues(values) {
  var summary = {
    salesMongo: 0,
    receivedTotal: 0,
    fee: 0,
    receivedWithFee: 0
  };

  (values || []).forEach(function(row) {
    row.forEach(function(cell, cellIndex) {
      var label = String(cell || '').replace(/\s+/g, ' ').trim();
      var normalizedLabel = label.toLowerCase();
      var amount = findAmountNearLabel(row, cellIndex);

      if (!amount) {
        return;
      }

      if (label.indexOf('ยอดขาย') !== -1 && normalizedLabel.indexOf('mongo') !== -1) {
        summary.salesMongo = amount;
      } else if (label.indexOf('รับจริงรวม') !== -1 && label.indexOf('Fee') !== -1) {
        summary.receivedWithFee = amount;
      } else if (label.indexOf('รับจริงรวม') !== -1 || label.indexOf('รับจริง') !== -1) {
        summary.receivedTotal = amount;
      } else if (label === 'Fee' || label.indexOf('ค่า Fee') !== -1 || label.indexOf('ค่าธรรมเนียม') !== -1) {
        summary.fee = amount;
      }
    });
  });

  if (!summary.receivedWithFee) {
    summary.receivedWithFee = summary.receivedTotal + summary.fee;
  }

  return summary;
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
