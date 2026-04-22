var REFACTOR_FORMAT_SHEET_KEYS = [
  'snapshots'
];

function getRefactorFormatConfigs_() {
  return {
    snapshots: {
      fixedColumnCount: 19,
      dateColumn: REFACTOR_COLUMNS.snapshots.date,
      amountColumns: [
        REFACTOR_COLUMNS.snapshots.totalAssets,
        REFACTOR_COLUMNS.snapshots.netFlow,
        REFACTOR_COLUMNS.snapshots.dailyReturn,
        REFACTOR_COLUMNS.snapshots.cumulativeNetFlow,
        REFACTOR_COLUMNS.snapshots.broker,
        REFACTOR_COLUMNS.snapshots.alipay,
        REFACTOR_COLUMNS.snapshots.mybank,
        REFACTOR_COLUMNS.snapshots.equity,
        REFACTOR_COLUMNS.snapshots.debt,
        REFACTOR_COLUMNS.snapshots.commodity,
        REFACTOR_COLUMNS.snapshots.cash,
        REFACTOR_COLUMNS.snapshots.dailyFee,
        REFACTOR_COLUMNS.snapshots.totalProfit,
        REFACTOR_COLUMNS.snapshots.brokerProfit,
        REFACTOR_COLUMNS.snapshots.alipayFundProfit
      ],
      integerColumns: [REFACTOR_COLUMNS.snapshots.shares],
      rateColumns: [REFACTOR_COLUMNS.snapshots.recalculatedNav],
      percentColumns: [REFACTOR_COLUMNS.snapshots.feeRate],
      signedBackgroundAmountColumns: [
        REFACTOR_COLUMNS.snapshots.netFlow,
        REFACTOR_COLUMNS.snapshots.dailyReturn,
        REFACTOR_COLUMNS.snapshots.totalProfit,
        REFACTOR_COLUMNS.snapshots.brokerProfit,
        REFACTOR_COLUMNS.snapshots.alipayFundProfit
      ],
      compactColumns: {
        1: 88,
        2: 98,
        3: 98,
        4: 88,
        5: 98,
        6: 86,
        7: 78,
        8: 84,
        9: 84,
        10: 92,
        11: 84,
        12: 84,
        13: 72,
        14: 72,
        15: 72,
        16: 84,
        17: 92,
        18: 92,
        19: 112
      }
    }
  };
}

function refactorApplySheetFormatting_() {
  REFACTOR_FORMAT_SHEET_KEYS.forEach(function(sheetKey) {
    refactorFormatSheetByKey_(sheetKey);
  });
  logRefactor_('表格格式整理完成');
}

function refactorFormatSheetByKey_(sheetKey) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = REFACTOR_SHEET_NAMES[sheetKey];
  var sheet = ss.getSheetByName(sheetName);
  var formatConfigs = getRefactorFormatConfigs_();
  if (!sheet) {
    logRefactor_('格式整理跳过', { sheetKey: sheetKey, reason: 'missing_sheet' });
    return { sheetKey: sheetKey, updated: false };
  }

  var options = formatConfigs[sheetKey] || {};
  applyTableFormatting_(sheet, options);
  var result = {
    sheetKey: sheetKey,
    sheetName: sheetName,
    rows: sheet.getLastRow(),
    columns: sheet.getLastColumn(),
    updated: true
  };
  logRefactor_('单表格式整理完成', result);
  return result;
}

function applyTableFormatting_(sheet, options) {
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return;

  options = options || {};
  if (options.fixedColumnCount) {
    ensureSheetColumnCount_(sheet, options.fixedColumnCount);
  }
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var range = sheet.getRange(1, 1, lastRow, lastColumn);

  applyBaseSheetFormatting_(sheet, range, lastRow, lastColumn);
  applyNumberFormats_(sheet, options, lastRow, lastColumn);
  applyConditionalStyles_(sheet, options, lastRow);
  applyConfiguredColumnWidths_(sheet, lastColumn, options.compactColumns);
}

function applyBaseSheetFormatting_(sheet, range, lastRow, lastColumn) {
  range.setFontSize(10);
  range.setWrap(false);
  range.setVerticalAlignment('middle');
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 24);
  if (lastRow > 1) {
    sheet.setRowHeights(2, lastRow - 1, 20);
  }
  sheet.getRange(1, 1, 1, lastColumn)
    .setFontWeight('bold')
    .setBackground('#d9e8fb')
    .setHorizontalAlignment('center');
}

function applyNumberFormats_(sheet, options, lastRow, lastColumn) {
  if (lastRow <= 1) return;
  var columnFormats = {};
  setColumnFormats_(columnFormats, options.amountColumns, '#,##0.0', lastColumn);
  setColumnFormats_(columnFormats, options.integerColumns, '#,##0', lastColumn);
  setColumnFormats_(columnFormats, options.rateColumns, '0.0000', lastColumn);
  setColumnFormats_(columnFormats, options.percentColumns, '0.00%', lastColumn);
  setColumnFormats_(columnFormats, options.colorAmountColumns, '[Red]#,##0.0;[Color10]-#,##0.0;#,##0.0', lastColumn);
  setColumnFormats_(columnFormats, options.colorRateColumns, '[Red]0.0000;[Color10]-0.0000;0.0000', lastColumn);
  setColumnFormats_(columnFormats, options.colorPercentColumns, '[Red]0.00%;[Color10]-0.00%;0.00%', lastColumn);

  if (options.dateColumn && options.dateColumn <= lastColumn) {
    columnFormats[options.dateColumn] = 'yyyy-mm-dd hh:mm:ss';
  }

  var formats = [];
  for (var r = 1; r < lastRow; r++) {
    var rowFormats = [];
    for (var c = 1; c <= lastColumn; c++) {
      rowFormats.push(columnFormats[c] || 'General');
    }
    formats.push(rowFormats);
  }
  sheet.getRange(2, 1, lastRow - 1, lastColumn).setNumberFormats(formats);
}

function setColumnFormats_(columnFormats, columns, format, lastColumn) {
  if (!columns || !columns.length) return;
  columns.forEach(function(column) {
    if (column <= lastColumn) columnFormats[column] = format;
  });
}

function applyConditionalStyles_(sheet, options, lastRow) {
  if (lastRow <= 1) return;
  var columns = []
    .concat(options.signedBackgroundAmountColumns || [])
    .concat(options.signedBackgroundPercentColumns || []);

  if (!columns.length) {
    sheet.setConditionalFormatRules([]);
    return;
  }

  var rules = [];
  columns.forEach(function(column) {
    if (column > sheet.getLastColumn()) return;
    var range = sheet.getRange(2, column, lastRow - 1, 1);
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0)
        .setBackground('#f4c7c3')
        .setRanges([range])
        .build()
    );
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0)
        .setBackground('#b6d7a8')
        .setRanges([range])
        .build()
    );
  });
  sheet.setConditionalFormatRules(rules);
}

function applyConfiguredColumnWidths_(sheet, lastColumn, compactColumns) {
  if (lastColumn <= 0 || !compactColumns) return;
  for (var column = 1; column <= lastColumn; column++) {
    if (compactColumns[column]) {
      sheet.setColumnWidth(column, compactColumns[column]);
    }
  }
}
