// ------------------------------
// 每日任务
// ------------------------------
// 适合绑定时间触发器，每天跑一次。

function refactorDailyRun() {
  refactorUpdateAssetPrices();
  refactorDailyUpdate();
  refactorUpdateSummaryFromSnapshots();
  refactorRefreshValidationSheet();
}

// 每日只维护当天那一行快照；如果当天已经存在，则覆盖更新。
function refactorDailyUpdate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var snapshotSheet = ss.getSheetByName(REFACTOR_SHEET_NAMES.snapshots);
  if (!snapshotSheet) throw new Error('找不到工作表: ' + REFACTOR_SHEET_NAMES.snapshots);
  ensureSnapshotSheetLayout_(snapshotSheet);

  var today = normalizeDate_(new Date());
  var todayKey = formatDateKey_(today);
  var data = snapshotSheet.getLastRow() > 1
    ? snapshotSheet.getRange(2, 1, snapshotSheet.getLastRow() - 1, REFACTOR_COLUMNS.snapshots.cumulativeNetFlow).getValues()
    : [];

  var dailyFlowMap = buildDailyInvestmentFlowMap_(ss);
  var todayFlow = dailyFlowMap[todayKey] || 0;
  var totalAssets = getCurrentTotalAssets_(ss);
  var previousCumulativeFlow = data.length
    ? (toNumber_(data[data.length - 1][REFACTOR_COLUMNS.snapshots.cumulativeNetFlow - 1]) || 0)
    : 0;
  var cumulativeNetFlow = previousCumulativeFlow + todayFlow;

  if (data.length) {
    var lastRow = data[data.length - 1];
    var lastDate = normalizeDate_(lastRow[0]);
    if (lastDate && formatDateKey_(lastDate) === todayKey) {
      lastRow[1] = totalAssets;
      lastRow[2] = '';
      lastRow[3] = '';
      lastRow[4] = todayFlow;
      lastRow[8] = cumulativeNetFlow;
      finalizeSnapshotRows_(data);
      snapshotSheet.getRange(2, 1, data.length, REFACTOR_COLUMNS.snapshots.cumulativeNetFlow).setValues(data);
      refactorUpdateSummaryFromSnapshots();
      return;
    }
  }

  data.push([
    today,
    totalAssets,
    '',
    '',
    todayFlow,
    '',
    '',
    '',
    cumulativeNetFlow
  ]);
  finalizeSnapshotRows_(data);
  snapshotSheet.getRange(2, 1, data.length, REFACTOR_COLUMNS.snapshots.cumulativeNetFlow).setValues(data);
  refactorUpdateSummaryFromSnapshots();
}

// 摘要区只读快照表结果，不再自己重复算一套净值轨迹。
function refactorUpdateSummaryFromSnapshots() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var snapshotSheet = ss.getSheetByName(REFACTOR_SHEET_NAMES.snapshots);
  if (!snapshotSheet || snapshotSheet.getLastRow() <= 1) return;

  var rows = snapshotSheet.getRange(2, 1, snapshotSheet.getLastRow() - 1, REFACTOR_COLUMNS.snapshots.cumulativeNetFlow).getValues()
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
    dailyReturn: latest[5] || 0,
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
