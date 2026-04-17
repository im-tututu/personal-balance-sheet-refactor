// ------------------------------
// 一次性初始化
// ------------------------------
// 只在新表刚建好、需要导入旧数据时运行。

function refactorSetupOnce() {
  refactorMigrateSourceData();
  refactorUpdateAssetPrices();
  refactorInitSnapshots();
  refactorUpdateSummaryFromSnapshots();
}

// 从旧表复制静态数据到新表，之后新表就不再依赖旧表。
function refactorMigrateSourceData() {
  var targetSs = SpreadsheetApp.getActiveSpreadsheet();
  var sourceUrl = getConfigValue_(targetSs, 'source_url');
  var sourceSs = SpreadsheetApp.openByUrl(sourceUrl);

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

// 用迁移后的投资流水反推历史净值轨迹，生成第一版快照表。
function refactorInitSnapshots() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var snapshotSheet = ss.getSheetByName(REFACTOR_SHEET_NAMES.snapshots);
  if (!snapshotSheet) throw new Error('找不到工作表: ' + REFACTOR_SHEET_NAMES.snapshots);

  var existingLastRow = snapshotSheet.getLastRow();
  if (existingLastRow > 1) {
    snapshotSheet.getRange(2, 1, existingLastRow - 1, snapshotSheet.getMaxColumns()).clearContent();
  }

  var dailyFlowMap = buildDailyInvestmentFlowMap_(ss);
  var flowDates = Object.keys(dailyFlowMap).sort();
  if (!flowDates.length) {
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
  var baseNetAssets = getCurrentNetAssets_(ss);
  var cumulativeFlow = 0;

  for (var i = 0; i < flowDates.length; i++) {
    var day = flowDates[i];
    cumulativeFlow += dailyFlowMap[day];
    rows.push([
      parseDateKey_(day),
      '',
      '',
      baseNetAssets - cumulativeFlow,
      dailyFlowMap[day],
      '',
      '',
      ''
    ]);
  }

  rows[rows.length - 1][1] = getCurrentTotalAssets_(ss);
  rows[rows.length - 1][2] = getCurrentTotalLiabilities_(ss);
  rows[rows.length - 1][3] = getCurrentNetAssets_(ss);

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
