// ------------------------------
// 一次性初始化
// ------------------------------
// 只在新表刚建好、需要导入旧数据时运行。

function refactorSetupOnce() {
  var todayKey = formatDateKey_(normalizeDate_(new Date()));
  logRefactor_('开始执行初始化任务', { date: todayKey });

  refactorMigrateSourceData();
  var snapshotResult = refactorInitSnapshots();

  var result = {
    date: todayKey,
    snapshotRows: snapshotResult ? snapshotResult.snapshotRows : 0
  };
  logRefactor_('初始化任务执行完成', result);
  return result;
}

// 从旧表复制静态数据到新表，之后新表就不再依赖旧表。
function refactorMigrateSourceData() {
  var targetSs = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSs = SpreadsheetApp.openById(REFACTOR_SOURCE_SPREADSHEET_ID);

  copySourceSheetAs_({
    sourceSpreadsheet: sourceSs,
    sourceSheetName: '资产',
    targetSpreadsheet: targetSs,
    targetSheetName: REFACTOR_SHEET_NAMES.assets
  });

  copySourceSheetAs_({
    sourceSpreadsheet: sourceSs,
    sourceSheetName: '总流水',
    targetSpreadsheet: targetSs,
    targetSheetName: REFACTOR_SHEET_NAMES.flows
  });
}

// 历史快照先导入旧表“市值记录”的原始历史数据，再计算份额和净值。
function refactorInitSnapshots() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var snapshotSheet = ss.getSheetByName(REFACTOR_SHEET_NAMES.snapshots);
  if (!snapshotSheet) throw new Error('找不到工作表: ' + REFACTOR_SHEET_NAMES.snapshots);

  var flowTimeline = buildInvestmentFlowTimeline_(ss);
  var undatedFlowTotal = getUndatedInvestmentFlowTotal_(ss);
  var marketRows = getHistoricalMarketValueRows_();
  var rows = buildSnapshotSeedRows_(marketRows, flowTimeline, undatedFlowTotal);

  importRawSnapshotRows_(snapshotSheet, rows);
  return recalculateSnapshotRows_(snapshotSheet);
}

function copySourceSheetAs_(options) {
  var sourceSheet = options.sourceSpreadsheet.getSheetByName(options.sourceSheetName);
  if (!sourceSheet) throw new Error('找不到源工作表: ' + options.sourceSheetName);

  var targetSheet = options.targetSpreadsheet.getSheetByName(options.targetSheetName);
  var temporarySheet = null;
  if (targetSheet) {
    if (options.targetSpreadsheet.getSheets().length === 1) {
      temporarySheet = options.targetSpreadsheet.insertSheet('_tmp_import_guard');
    }
    options.targetSpreadsheet.deleteSheet(targetSheet);
  }

  var copiedSheet = sourceSheet.copyTo(options.targetSpreadsheet);
  copiedSheet.setName(options.targetSheetName);
  if (temporarySheet) {
    options.targetSpreadsheet.deleteSheet(temporarySheet);
  }
}

// 兼容旧入口，避免已经配好的手动脚本名失效。
function refactorBootstrap() {
  refactorSetupOnce();
}
