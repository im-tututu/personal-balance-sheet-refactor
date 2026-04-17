// ------------------------------
// 每日任务
// ------------------------------
// 适合绑定时间触发器，每天跑一次。

function refactorDailyRun() {
  var todayKey = formatDateKey_(normalizeDate_(new Date()));
  logRefactor_('开始执行每日任务', { date: todayKey });

  refactorUpdateAssetPrices();
  var assetSnapshotResult = refactorUpsertCurrentAssetSnapshot_();
  var marketHistoryResult = refactorUpdateMarketValueHistoryFromAssetSnapshots_([todayKey]);
  refactorDailyUpdate();
  refactorUpdateSummaryFromSnapshots();
  refactorRefreshValidationSheet();

  var result = {
    date: todayKey,
    assetSnapshotRows: assetSnapshotResult ? assetSnapshotResult.snapshotRows : 0,
    marketHistoryDates: marketHistoryResult ? marketHistoryResult.upsertedDates : 0
  };
  logRefactor_('每日任务执行完成', result);
  return result;
}

// 每日只维护当天那一行快照；如果当天已经存在，则覆盖更新。
function refactorDailyUpdate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var snapshotSheet = ss.getSheetByName(REFACTOR_SHEET_NAMES.snapshots);
  if (!snapshotSheet) throw new Error('找不到工作表: ' + REFACTOR_SHEET_NAMES.snapshots);
  ensureSnapshotSheetLayout_(snapshotSheet);

  var today = normalizeDate_(new Date());
  var data = snapshotSheet.getLastRow() > 1
    ? snapshotSheet.getRange(
        2,
        1,
        snapshotSheet.getLastRow() - 1,
        REFACTOR_COLUMNS.snapshots.shares
      ).getValues()
    : [];

  var dailyFlowMap = buildDailyInvestmentFlowMap_(ss);
  var openingCumulativeNetFlow = getSnapshotOpeningCumulativeFlow_(ss);
  var totalAssets = getCurrentTotalAssets_(ss);
  var todayCumulativeNetFlow = getCumulativeInvestmentFlowAsOf_(dailyFlowMap, today);
  var todayKey = formatDateKey_(today);
  var existingTodayIndex = -1;

  for (var i = 0; i < data.length; i++) {
    var rowDate = normalizeDate_(data[i][0]);
    if (rowDate && formatDateKey_(rowDate) === todayKey) {
      existingTodayIndex = i;
      break;
    }
  }

  if (existingTodayIndex > -1) {
    data[existingTodayIndex][1] = totalAssets;
    data[existingTodayIndex][5] = todayCumulativeNetFlow;
    finalizeSnapshotRows_(data, openingCumulativeNetFlow);
    snapshotSheet.getRange(2, 1, data.length, REFACTOR_COLUMNS.snapshots.shares).setValues(data);
    refactorUpdateSummaryFromSnapshots();
    return;
  }

  data.push([
    today,
    totalAssets,
    '',
    '',
    '',
    todayCumulativeNetFlow,
    ''
  ]);

  finalizeSnapshotRows_(data, openingCumulativeNetFlow);
  snapshotSheet.getRange(2, 1, data.length, REFACTOR_COLUMNS.snapshots.shares).setValues(data);
  refactorUpdateSummaryFromSnapshots();
}


// 摘要区只读快照表结果，不再自己重复算一套净值轨迹。
function refactorUpdateSummaryFromSnapshots() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var snapshotSheet = ss.getSheetByName(REFACTOR_SHEET_NAMES.snapshots);
  if (!snapshotSheet || snapshotSheet.getLastRow() <= 1) return;

  var rows = snapshotSheet.getRange(2, 1, snapshotSheet.getLastRow() - 1, REFACTOR_COLUMNS.snapshots.shares).getValues()
    .filter(function(row) { return row[0]; })
    .sort(function(a, b) {
      return normalizeDate_(a[0]).getTime() - normalizeDate_(b[0]).getTime();
    });
  if (!rows.length) return;

  var latest = rows[rows.length - 1];
  var xirr = calculatePortfolioXirr_(ss, rows);
  var drawdownMeta = getMaxDrawdownMeta_(rows);

  refactorUpdateSummary_({
    latestDate: latest[0],
    totalAssets: latest[1] || 0,
    netAssets: '',
    dailyReturn: latest[3] || 0,
    xirr: xirr,
    maxDrawdown: drawdownMeta.maxDrawdown,
    drawdownFrom: drawdownMeta.fromDate,
    drawdownTo: drawdownMeta.toDate
  });
}

// 兼容旧入口，避免旧触发器失效。
function refactorRunAll() {
  refactorDailyRun();
}
