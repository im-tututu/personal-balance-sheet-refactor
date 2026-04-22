// ------------------------------
// 每日任务
// ------------------------------
// 适合绑定时间触发器，每天跑一次。

function refactorDailyRun() {
  var todayKey = formatDateKey_(normalizeDate_(new Date()));
  logRefactor_('开始执行每日任务', { date: todayKey });

  var updateResult = refactorDailyUpdate();

  var result = {
    date: todayKey,
    totalRows: updateResult ? updateResult.totalRows : 0
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
  var todayRow = buildCurrentSnapshotRow_(ss, today);
  return upsertSnapshotRows_(snapshotSheet, [todayRow], false);
}

// 兼容旧入口，避免旧触发器失效。
function refactorRunAll() {
  refactorDailyRun();
}
