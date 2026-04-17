// ------------------------------
// 一次性初始化
// ------------------------------
// 只在新表刚建好、需要导入旧数据时运行。

function refactorSetupOnce() {
  refactorMigrateSourceData();
  refactorUpdateAssetPrices();
  refactorInitSnapshots();
  refactorUpdateSummaryFromSnapshots();
  refactorRefreshValidationSheet();
}

// 从旧表复制静态数据到新表，之后新表就不再依赖旧表。
function refactorMigrateSourceData() {
  var targetSs = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSs = SpreadsheetApp.openById(REFACTOR_SOURCE_SPREADSHEET_ID);

  migrateSheetValues_({
    sourceSpreadsheet: sourceSs,
    sourceSheetName: '资产',
    targetSpreadsheet: targetSs,
    targetSheetName: REFACTOR_SHEET_NAMES.assets
  });

  migrateSheetValues_({
    sourceSpreadsheet: sourceSs,
    sourceSheetName: '总流水',
    targetSpreadsheet: targetSs,
    targetSheetName: REFACTOR_SHEET_NAMES.flows
  });

  rebuildAssetDerivedColumns_(targetSs);
}

// 历史快照以旧表“市值记录”为准，再叠加投资流水得到净现金流合计。
function refactorInitSnapshots() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var snapshotSheet = ss.getSheetByName(REFACTOR_SHEET_NAMES.snapshots);
  if (!snapshotSheet) throw new Error('找不到工作表: ' + REFACTOR_SHEET_NAMES.snapshots);
  ensureSnapshotSheetLayout_(snapshotSheet);

  var existingLastRow = snapshotSheet.getLastRow();
  if (existingLastRow > 1) {
    snapshotSheet.getRange(2, 1, existingLastRow - 1, snapshotSheet.getMaxColumns()).clearContent();
  }

  var dailyFlowMap = buildDailyInvestmentFlowMap_(ss);
  var marketRows = getHistoricalMarketValueRows_();
  if (!marketRows.length) {
    refactorUpdateSummary_({
      latestDate: '',
      totalAssets: 0,
      netAssets: 0,
      dailyReturn: 0,
      xirr: '',
      maxDrawdown: 0,
      drawdownFrom: '',
      drawdownTo: ''
    });
    return;
  }

  var rows = [];
  var cumulativeFlow = 0;

  for (var i = 0; i < marketRows.length; i++) {
    var day = formatDateKey_(marketRows[i].date);
    var dayFlow = dailyFlowMap[day] || 0;
    cumulativeFlow += dayFlow;
    rows.push([
      marketRows[i].date,
      marketRows[i].totalAssets,
      '',
      '',
      dayFlow,
      '',
      '',
      '',
      cumulativeFlow
    ]);
  }

  finalizeSnapshotRows_(rows);
  snapshotSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  refactorUpdateSummaryFromSnapshots();
}

function migrateSheetValues_(options) {
  var sourceSheet = options.sourceSpreadsheet.getSheetByName(options.sourceSheetName);
  var targetSheet = options.targetSpreadsheet.getSheetByName(options.targetSheetName);
  if (!sourceSheet) throw new Error('找不到源工作表: ' + options.sourceSheetName);
  if (!targetSheet) throw new Error('找不到目标工作表: ' + options.targetSheetName);

  var lastRow = sourceSheet.getLastRow();
  var lastColumn = sourceSheet.getLastColumn();
  if (!lastRow || !lastColumn) return;

  var values = sourceSheet.getRange(1, 1, lastRow, Math.min(lastColumn, 17)).getValues();
  targetSheet.clearContents();
  targetSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
}

// 兼容旧入口，避免已经配好的手动脚本名失效。
function refactorBootstrap() {
  refactorSetupOnce();
}
