/**
 * スプレッドシートを開いたときのメニューと、統合タブの自動更新。
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('タスク表')
    .addItem('完了タスクを完了タブへ移す', 'archiveCompletedTasksFromMenu')
    .addItem('移す対象を確認する（移動しない）', 'previewCompletedTasks')
    .addSeparator()
    .addItem('統合タブを作り直す', 'rebuildIntegrationFromMenu')
    .addItem('ダッシュボードを作り直す', 'setupDashboard')
    .addSeparator()
    .addItem('自動処理を設定（タブ追加時の統合更新・毎週月曜のアーカイブ）', 'installTriggers')
    .addItem('自動処理を解除', 'removeTriggers')
    .addToUi();

  // タブが増減・改名されていたら統合タブの数式を合わせる
  try {
    refreshIntegration_(false);
  } catch (e) {
    console.warn(e);
  }
}

function installTriggers() {
  removeTriggers();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('handleSpreadsheetChange').forSpreadsheet(ss).onChange().create();
  installWeeklyTrigger();
  ss.toast('タブの追加・削除・名前変更で統合タブを自動更新し、毎週月曜 7時台に完了タスクを移します。');
}

function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'handleSpreadsheetChange') ScriptApp.deleteTrigger(t);
  });
  removeWeeklyTrigger();
}

/** タブの追加・削除・名前変更（changeType は INSERT_GRID / REMOVE_GRID / OTHER）で統合タブを合わせる */
function handleSpreadsheetChange(e) {
  if (e && ['INSERT_GRID', 'REMOVE_GRID', 'OTHER'].indexOf(e.changeType) === -1) return;
  refreshIntegration_(false);
}
