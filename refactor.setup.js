// ------------------------------
// 一次性初始化
// ------------------------------
// 只在新表刚建好、需要导入旧数据时运行。

function refactorSetupOnce() {
  var todayKey = formatDateKey_(normalizeDate_(new Date()));
  logRefactor_('开始执行初始化任务', { date: todayKey });

  refactorMigrateSourceData();
  var archiveResult = refactorArchiveHistoricalMarketValue_();
  refactorUpdateAssetPrices();
  var assetSnapshotResult = refactorUpsertCurrentAssetSnapshot_();
  var marketHistoryResult = refactorUpdateMarketValueHistoryFromAssetSnapshots_([todayKey]);
  refactorInitSnapshots();
  refactorUpdateSummaryFromSnapshots();
  refactorRefreshValidationSheet();

  var result = {
    date: todayKey,
    insertedHistoricalDates: archiveResult ? archiveResult.insertedDates : 0,
    assetSnapshotRows: assetSnapshotResult ? assetSnapshotResult.snapshotRows : 0,
    marketHistoryDates: marketHistoryResult ? marketHistoryResult.upsertedDates : 0
  };
  logRefactor_('初始化任务执行完成', result);
  return result;
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

  var flowTimeline = buildInvestmentFlowTimeline_(ss);
  var openingCumulativeNetFlow = getSnapshotOpeningCumulativeFlow_(ss);
  var marketRows = getHistoricalMarketValueRows_();
  var rows = buildSnapshotSeedRows_(marketRows, flowTimeline);

  if (!rows.length) {
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

  finalizeSnapshotRows_(rows, openingCumulativeNetFlow);
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

  if (targetSheet.getMaxColumns() < lastColumn) {
    targetSheet.insertColumnsAfter(targetSheet.getMaxColumns(), lastColumn - targetSheet.getMaxColumns());
  }
  if (targetSheet.getMaxRows() < lastRow) {
    targetSheet.insertRowsAfter(targetSheet.getMaxRows(), lastRow - targetSheet.getMaxRows());
  }

  var targetWholeRange = targetSheet.getRange(1, 1, targetSheet.getMaxRows(), targetSheet.getMaxColumns());
  targetWholeRange.clearContent();
  targetWholeRange.clearFormat();
  targetWholeRange.clearDataValidations();
  targetWholeRange.clearNote();

  var sourceRange = sourceSheet.getRange(1, 1, lastRow, lastColumn);
  var targetRange = targetSheet.getRange(1, 1, lastRow, lastColumn);

  var values = sourceRange.getValues();
  var formulas = sourceRange.getFormulas();
  var numberFormats = sourceRange.getNumberFormats();
  var backgrounds = sourceRange.getBackgrounds();
  var fontColors = sourceRange.getFontColors();
  var fontWeights = sourceRange.getFontWeights();
  var horizontalAlignments = sourceRange.getHorizontalAlignments();
  var verticalAlignments = sourceRange.getVerticalAlignments();
  var wraps = sourceRange.getWraps();
  var notes = sourceRange.getNotes();

  var mixed = [];
  for (var r = 0; r < lastRow; r++) {
    var row = [];
    for (var c = 0; c < lastColumn; c++) {
      row.push(formulas[r][c] ? formulas[r][c] : values[r][c]);
    }
    mixed.push(row);
  }

  targetRange.setValues(mixed);
  targetRange.setNumberFormats(numberFormats);
  targetRange.setBackgrounds(backgrounds);
  targetRange.setFontColors(fontColors);
  targetRange.setFontWeights(fontWeights);
  targetRange.setHorizontalAlignments(horizontalAlignments);
  targetRange.setVerticalAlignments(verticalAlignments);
  targetRange.setWraps(wraps);
  targetRange.setNotes(notes);
}

// 兼容旧入口，避免已经配好的手动脚本名失效。
function refactorBootstrap() {
  refactorSetupOnce();
}
