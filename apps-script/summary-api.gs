function doGet(event) {
  var params = event && event.parameter ? event.parameter : {};

  if (params.action === 'saveSource') {
    return saveSourceConfig(params);
  }

  if (params.action === 'savePendingApprovals') {
    return savePendingApprovals(params);
  }

  if (params.action === 'writeTestSuccess') {
    return writeTestSuccess(params);
  }

  var sourceConfig = Object.assign({}, getSourceConfig(), getRequestSourceConfig(params));
  var spreadsheetId = sourceConfig.spreadsheetId;
  var spreadsheetUrl = sourceConfig.spreadsheetUrl;
  var sheetName = sourceConfig.sheetName;
  var incomeSheetName = sourceConfig.incomeSheetName;
  var pendingSheetName = sourceConfig.pendingSheetName;
  var budgetSheetName = sourceConfig.budgetSheetName;
  var paidSheetName = sourceConfig.paidSheetName;
  var rangeA1 = sourceConfig.range;
  var incomeRangeA1 = sourceConfig.incomeRange;
  var pendingRangeA1 = sourceConfig.pendingRange;
  var budgetRangeA1 = sourceConfig.budgetRange;
  var paidDetailRangeA1 = sourceConfig.paidDetailRange;
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);
  var incomeSheet = spreadsheet.getSheetByName(incomeSheetName);
  var pendingSheet = spreadsheet.getSheetByName(pendingSheetName);
  var budgetSheet = spreadsheet.getSheetByName(budgetSheetName);
  var paidSheet = spreadsheet.getSheetByName(paidSheetName);
  var values = sheet.getRange(rangeA1).getDisplayValues();
  var incomeValues = incomeSheet ? incomeSheet.getRange(incomeRangeA1).getDisplayValues() : [];
  var pendingValues = pendingSheet ? pendingSheet.getRange(pendingRangeA1).getDisplayValues() : [];
  var budgetValues = budgetSheet ? budgetSheet.getRange(budgetRangeA1).getDisplayValues() : [];
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
  var budgetRows = budgetValues.slice(1).filter(function(row) {
    return row.some(function(value) {
      return value !== '';
    });
  });

  var payload = {
    source: {
      spreadsheetId: spreadsheetId,
      spreadsheetUrl: spreadsheetUrl,
      sheetName: sheetName,
      incomeSheetName: incomeSheetName,
      pendingSheetName: pendingSheetName,
      budgetSheetName: budgetSheetName,
      paidSheetName: paidSheetName,
      range: rangeA1,
      incomeRange: incomeRangeA1,
      pendingRange: pendingRangeA1,
      budgetRange: budgetRangeA1,
      paidDetailRange: paidDetailRangeA1,
      snapshotDate: snapshotDate
    },
    header: header,
    rows: dataRows,
    pendingHeader: pendingHeader,
    pendingRows: pendingRows,
    budgetRows: budgetRows,
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

function savePendingApprovals(params) {
  var sourceConfig = Object.assign({}, getSourceConfig(), getRequestSourceConfig(params));
  var spreadsheet = SpreadsheetApp.openById(sourceConfig.spreadsheetId);
  var sheet = spreadsheet.getSheetByName(sourceConfig.pendingSheetName || 'เตรียมจ่าย');

  if (!sheet) {
    return outputJsonp(params.callback, { ok: false, error: 'ไม่พบแท็บเตรียมจ่าย' });
  }

  var approvals = {};

  try {
    JSON.parse(params.approvals || '[]').forEach(function(item) {
      var prPo = String(item.prPo || '').trim().toUpperCase();

      if (prPo) {
        approvals[prPo] = item.approved === true;
      }
    });
  } catch (error) {
    return outputJsonp(params.callback, { ok: false, error: 'ข้อมูล Approve ไม่ถูกต้อง' });
  }

  var range = sheet.getDataRange();
  var values = range.getDisplayValues();
  var headerIndex = -1;
  var prPoColumn = -1;
  var approveColumn = -1;

  values.some(function(row, rowIndex) {
    var normalized = row.map(function(value) {
      return String(value || '').trim().toUpperCase();
    });
    var foundPrPo = normalized.indexOf('PR/PO');
    var foundApprove = normalized.indexOf('APPROVE');

    if (foundPrPo !== -1 && foundApprove !== -1) {
      headerIndex = rowIndex;
      prPoColumn = foundPrPo;
      approveColumn = foundApprove;
      return true;
    }

    return false;
  });

  if (headerIndex === -1) {
    return outputJsonp(params.callback, { ok: false, error: 'ไม่พบหัวตาราง PR/PO และ Approve ในแท็บเตรียมจ่าย' });
  }

  var output = values.slice(headerIndex + 1).map(function(row) {
    var prPo = String(row[prPoColumn] || '').trim().toUpperCase();
    return [prPo && approvals[prPo] === true ? 'approve' : ''];
  });

  if (output.length) {
    sheet.getRange(headerIndex + 2, approveColumn + 1, output.length, 1).setValues(output);
  }

  return outputJsonp(params.callback, {
    ok: true,
    updatedRows: output.length
  });
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
    budgetSheetName: 'Budget',
    paidSheetName: 'จ่ายแล้ว_Unfiltered',
    range: 'A:F',
    incomeRange: 'A:Z',
    pendingRange: 'A:G',
    budgetRange: 'A:Z',
    paidDetailRange: 'A:P'
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
    'budgetSheetName',
    'paidSheetName',
    'range',
    'incomeRange',
    'pendingRange',
    'budgetRange',
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
    budgetSheetName: params.budgetSheetName || defaults.budgetSheetName,
    paidSheetName: params.paidSheetName || defaults.paidSheetName,
    range: params.range || defaults.range,
    incomeRange: params.incomeRange || defaults.incomeRange,
    pendingRange: params.pendingRange || defaults.pendingRange,
    budgetRange: params.budgetRange || defaults.budgetRange,
    paidDetailRange: params.paidDetailRange || defaults.paidDetailRange
  };

  PropertiesService.getScriptProperties().setProperty('sourceConfig', JSON.stringify(sourceConfig));

  return outputJsonp(params.callback, {
    ok: true,
    source: sourceConfig
  });
}

function writeTestSuccess(params) {
  var spreadsheetId = params.spreadsheetId || getDefaultSourceConfig().spreadsheetId;
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName('Test');

  if (!sheet) {
    return outputJsonp(params.callback, { ok: false, error: 'ไม่พบแท็บ Test' });
  }

  sheet.getRange('A1').setValue('Success');
  SpreadsheetApp.flush();

  return outputJsonp(params.callback, {
    ok: true,
    sheetName: 'Test',
    cell: 'A1',
    value: 'Success'
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
    receivedWithFee: 0,
    channels: [],
    bankSummary: {
      latestDate: '',
      accounts: [],
      total: 0
    }
  };
  var channelRows = [];
  var bankAccounts = [];
  var totalRow = null;
  var inBankSection = false;

  (values || []).forEach(function(row) {
    var firstCell = String(row[0] || '').trim();

    if (!summary.bankSummary.latestDate && row[1]) {
      summary.bankSummary.latestDate = row[1];
    }

    if (firstCell === 'ข้อมูลธนาคารล่าสุด') {
      summary.bankSummary.latestDate = row[1] || '';
    }

    if (firstCell === 'รวมทั้งหมด') {
      totalRow = row;
    } else if (firstCell === 'สรุปยอดคงเหลือธนาคาร') {
      inBankSection = true;
    } else if (inBankSection && firstCell === 'ธนาคาร / บัญชี') {
      return;
    } else if (inBankSection && firstCell) {
      var bankBalance = parseAmount(row[3] || row[2] || row[1]);

      if (firstCell.indexOf('รวม') !== -1) {
        summary.bankSummary.total = bankBalance;
        inBankSection = false;
      } else {
        bankAccounts.push({
          account: firstCell,
          balance: bankBalance
        });
      }
    } else if (
      firstCell &&
      ['ข้อมูลธนาคารล่าสุด', 'ช่องทาง', 'ยอดขายรวม', 'เงินเข้าจริงสุทธิ'].indexOf(firstCell) === -1 &&
      firstCell.indexOf('หมายเหตุ') !== 0
    ) {
      var salesMongo = parseAmount(row[1]);
      var fee = parseAmount(row[3]);
      var receivedWithFee = parseAmount(row[4]);

      if (salesMongo > 0 || receivedWithFee > 0) {
        channelRows.push({
          channel: firstCell,
          salesMongo: salesMongo,
          feePercent: parseAmount(row[2]),
          fee: fee,
          receivedWithFee: receivedWithFee,
          netReceived: parseAmount(row[7]),
          gap: parseAmount(row[8]),
          receivedPercent: parseAmount(row[9])
        });
      }
    }

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

  if (totalRow) {
    summary.salesMongo = parseAmount(totalRow[1]) || summary.salesMongo;
    summary.fee = parseAmount(totalRow[3]) || summary.fee;
    summary.receivedWithFee = parseAmount(totalRow[4]) || summary.receivedWithFee;
    summary.receivedTotal = parseAmount(totalRow[7]) || summary.receivedTotal;
  }

  if (!summary.receivedWithFee) {
    summary.receivedWithFee = summary.receivedTotal + summary.fee;
  }

  summary.channels = channelRows;
  summary.bankSummary.accounts = bankAccounts;

  if (!summary.bankSummary.total) {
    summary.bankSummary.total = bankAccounts.reduce(function(sum, item) {
      return sum + item.balance;
    }, 0);
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
