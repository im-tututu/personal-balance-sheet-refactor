var REFACTOR_FORMAT_SHEET_KEYS = [
  'assets',
  'flows',
  'snapshots',
  'marketValueHistory',
  'validation',
  'assetSnapshots'
];

function getRefactorFormatConfigs_() {
  return {
    assets: {
      dateColumn: 5,
      amountColumns: [8, 10],
      percentColumns: [11]
    },
    flows: {
      dateColumn: REFACTOR_COLUMNS.flows.date,
      amountColumns: [REFACTOR_COLUMNS.flows.cashflow]
    },
    snapshots: {
      fixedColumnCount: 8,
      dateColumn: REFACTOR_COLUMNS.snapshots.date,
      amountColumns: [
        REFACTOR_COLUMNS.snapshots.totalAssets,
        REFACTOR_COLUMNS.snapshots.netFlow,
        REFACTOR_COLUMNS.snapshots.dailyReturn,
        REFACTOR_COLUMNS.snapshots.cumulativeNetFlow
      ],
      integerColumns: [REFACTOR_COLUMNS.snapshots.shares],
      rateColumns: [REFACTOR_COLUMNS.snapshots.nav],
      percentColumns: [REFACTOR_COLUMNS.snapshots.drawdown],
      signedBackgroundAmountColumns: [
        REFACTOR_COLUMNS.snapshots.netFlow,
        REFACTOR_COLUMNS.snapshots.dailyReturn
      ],
      compactColumns: {
        1: 88,
        2: 98,
        3: 98,
        4: 88,
        5: 72,
        6: 72,
        7: 98,
        8: 86
      }
    },
    marketValueHistory: {
      dateColumn: 1,
      amountColumns: [2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 13, 14],
      percentColumns: [10],
      compactColumns: {
        1: 88,
        2: 96,
        3: 84,
        4: 84,
        5: 92,
        6: 84,
        7: 84,
        8: 72,
        9: 72,
        10: 72,
        11: 84,
        12: 92,
        13: 92,
        14: 112
      }
    },
    validation: {
      dateColumn: 1,
      amountColumns: [2, 3, 4, 5, 6, 7, 10],
      rateColumns: [8],
      percentColumns: [9],
      colorAmountColumns: [4, 6, 7],
      compactColumns: {
        1: 88,
        2: 98,
        3: 98,
        4: 98,
        5: 98,
        6: 88,
        7: 88,
        8: 72,
        9: 72,
        10: 88
      }
    },
    assetSnapshots: {
      dateColumn: 1,
      compactColumns: {
        1: 88
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

  var options = cloneFormatOptions_(formatConfigs[sheetKey] || {});
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

function cloneFormatOptions_(options) {
  return JSON.parse(JSON.stringify(options || {}));
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
  resetNumberFormats_(sheet);
  applyDateSortingAndFormatting_(sheet, options.dateColumn, lastRow, lastColumn);
  applyColumnFormats_(sheet, options);
  resetFilter_(sheet, lastRow, lastColumn);
  resetConditionalStyles_(sheet, options);
  resetBanding_(sheet, lastRow, lastColumn);
  autoResizeReasonableColumns_(sheet, lastColumn, options.compactColumns);
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

function applyDateSortingAndFormatting_(sheet, dateColumn, lastRow, lastColumn) {
  if (!(lastRow > 1 && dateColumn)) return;
  sheet.getRange(2, 1, lastRow - 1, lastColumn).sort({
    column: dateColumn,
    ascending: false
  });
  sheet.getRange(2, dateColumn, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
}

function applyColumnFormats_(sheet, options) {
  applyNumberFormatToColumns_(sheet, options.amountColumns, '#,##0.0');
  applyNumberFormatToColumns_(sheet, options.integerColumns, '#,##0');
  applyNumberFormatToColumns_(sheet, options.rateColumns, '0.0000');
  applyNumberFormatToColumns_(sheet, options.percentColumns, '0.00%');
  applyNumberFormatToColumns_(sheet, options.colorAmountColumns, '[Red]#,##0.0;[Color10]-#,##0.0;#,##0.0');
  applyNumberFormatToColumns_(sheet, options.colorRateColumns, '[Red]0.0000;[Color10]-0.0000;0.0000');
  applyNumberFormatToColumns_(sheet, options.colorPercentColumns, '[Red]0.00%;[Color10]-0.00%;0.00%');
}

function resetFilter_(sheet, lastRow, lastColumn) {
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  if (lastRow >= 1 && lastColumn >= 1) {
    sheet.getRange(1, 1, lastRow, lastColumn).createFilter();
  }
}

function resetConditionalStyles_(sheet, options) {
  clearConditionalFormatRules_(sheet);
  applySignedBackgroundRules_(sheet, options.signedBackgroundAmountColumns, '#f4c7c3', '#b6d7a8');
  applySignedBackgroundRules_(sheet, options.signedBackgroundPercentColumns, '#f4c7c3', '#b6d7a8');
}

function resetBanding_(sheet, lastRow, lastColumn) {
  clearBandings_(sheet);
  if (lastRow > 1) {
    sheet.getRange(1, 1, lastRow, lastColumn).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  }
}

function applyNumberFormatToColumns_(sheet, columns, format) {
  if (!columns || !columns.length || sheet.getLastRow() <= 1) return;
  columns.forEach(function(column) {
    if (column > sheet.getLastColumn()) return;
    sheet.getRange(2, column, sheet.getLastRow() - 1, 1).setNumberFormat(format);
  });
}

function resetNumberFormats_(sheet) {
  if (sheet.getLastRow() <= 1 || sheet.getLastColumn() <= 0) return;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).setNumberFormat('General');
}

function clearBandings_(sheet) {
  var bandings = sheet.getBandings();
  bandings.forEach(function(banding) {
    banding.remove();
  });
}

function clearConditionalFormatRules_(sheet) {
  sheet.setConditionalFormatRules([]);
}

function applySignedBackgroundRules_(sheet, columns, positiveColor, negativeColor) {
  if (!columns || !columns.length || sheet.getLastRow() <= 1) return;

  var rules = sheet.getConditionalFormatRules();
  columns.forEach(function(column) {
    if (column > sheet.getLastColumn()) return;
    var range = sheet.getRange(2, column, sheet.getLastRow() - 1, 1);
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0)
        .setBackground(positiveColor)
        .setRanges([range])
        .build()
    );
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0)
        .setBackground(negativeColor)
        .setRanges([range])
        .build()
    );
  });
  sheet.setConditionalFormatRules(rules);
}

function autoResizeReasonableColumns_(sheet, lastColumn, compactColumns) {
  if (lastColumn <= 0) return;
  sheet.autoResizeColumns(1, lastColumn);
  for (var column = 1; column <= lastColumn; column++) {
    if (compactColumns && compactColumns[column]) {
      sheet.setColumnWidth(column, compactColumns[column]);
      continue;
    }
    var width = sheet.getColumnWidth(column);
    if (width > 160) {
      sheet.setColumnWidth(column, 160);
    } else if (width < 60) {
      sheet.setColumnWidth(column, 60);
    }
  }
}
