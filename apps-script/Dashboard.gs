/**
 * ダッシュボードタブ：期日が近い未完了タスクの一覧
 *
 * TASK表 統合タブを元に、数式だけで動く（開くたび・編集のたびに最新になる）。
 * - E2 の「表示する日数」以内に期日が来るタスクと、期日を過ぎたタスクを期日順に出す
 * - H2 で担当を選ぶと、その人のタスクだけに絞れる（空欄なら全員）
 * - ステータスか状態が「完了」の行は出さない
 * - 期限切れ・今日・3日以内・7日以内で行に色を付ける
 */

var DASHBOARD_SHEET_NAME = 'ダッシュボード';
var DASHBOARD_DEFAULT_DAYS = 14;
var DASHBOARD_LIST_ROW = 8;   // 一覧の1行目（見出しはその1行上）

var DASHBOARD_COLORS = {
  overdue: '#f4cccc',
  today: '#fce5cd',
  within3: '#fff2cc',
  within7: '#d9ead3'
};

function setupDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  var days = DASHBOARD_DEFAULT_DAYS;
  var person = '';
  if (sheet) {
    // 作り直しても、設定した日数と担当は引き継ぐ
    days = Number(sheet.getRange('E2').getValue()) || DASHBOARD_DEFAULT_DAYS;
    person = sheet.getRange('H2').getValue();
    sheet.clear();
    sheet.clearConditionalFormatRules();
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
  } else {
    sheet = ss.insertSheet(DASHBOARD_SHEET_NAME, 0);
  }

  var src = "'" + INTEGRATED_SHEET_NAME + "'!";
  var L = DASHBOARD_LIST_ROW;

  // 1〜2行目：タイトルと設定
  sheet.getRange('A1').setValue('期日が近いタスク').setFontSize(16).setFontWeight('bold');
  sheet.getRange('A2:H2').setValues([['今日', '=TODAY()', '', '表示する日数', days, '', '担当', person]]);
  sheet.getRange('B2').setNumberFormat(DATE_FORMAT);
  sheet.getRange('A2:H2').setFontColor('#666666');
  sheet.getRange('E2').setBackground('#fff8e1').setFontColor('#000000')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(0, 365).build());
  sheet.getRange('H2').setBackground('#fff8e1').setFontColor('#000000')
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInRange(ss.getRange(src + 'E2:E'), true).setAllowInvalid(true).build());

  // 4〜5行目：件数
  var b = 'B' + L + ':B';
  var tiles = [
    ['期限切れ', '=COUNTIF(' + b + ',"<0")', DASHBOARD_COLORS.overdue],
    ['今日', '=COUNTIF(' + b + ',0)', DASHBOARD_COLORS.today],
    ['1〜3日後', '=COUNTIFS(' + b + ',">=1",' + b + ',"<=3")', DASHBOARD_COLORS.within3],
    ['4〜7日後', '=COUNTIFS(' + b + ',">=4",' + b + ',"<=7")', DASHBOARD_COLORS.within7],
    ['表示中の合計', '=COUNT(' + b + ')', '#eeeeee']
  ];
  tiles.forEach(function (t, i) {
    sheet.getRange(4, i + 1).setValue(t[0]).setFontColor('#666666');
    sheet.getRange(5, i + 1).setFormula(t[1]).setFontSize(20).setFontWeight('bold');
    sheet.getRange(4, i + 1, 2, 1).setBackground(t[2]).setHorizontalAlignment('center');
  });

  // 7行目〜：一覧
  var headers = ['期日', '残日数', 'カテゴリ', '項目', 'Task', '担当', '状態', '重要度', '状況詳細'];
  sheet.getRange(L - 1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#434343').setFontColor('#ffffff');

  // 統合タブの列：1カテゴリ 2項目 3Task 4状況詳細 5担当 6開始 7期日 8残日数 9状態 10重要度 11URL1 12URL2 13ステータス
  sheet.getRange(L, 1).setFormula(
    '=ARRAYFORMULA(IFERROR(LET(' +
    'tasks,' + src + 'A2:M,' +
    'due,CHOOSECOLS(tasks,7),' +
    'hits,FILTER(tasks,ISNUMBER(due),due<=TODAY()+$E$2,' +
    'CHOOSECOLS(tasks,13)<>"完了",CHOOSECOLS(tasks,9)<>"完了",' +
    '($H$2="")+(CHOOSECOLS(tasks,5)=$H$2)),' +
    'SORT(HSTACK(CHOOSECOLS(hits,7),CHOOSECOLS(hits,7)-TODAY(),CHOOSECOLS(hits,1,2,3,5,9,10,4)),1,TRUE,4,TRUE)' +
    '),"該当するタスクはありません"))');

  var rows = sheet.getMaxRows() - L + 1;
  sheet.getRange(L, 1, rows, 1).setNumberFormat(DATE_FORMAT);
  sheet.getRange(L, 2, rows, 1).setNumberFormat('0"日後";0"日超過";"今日"').setHorizontalAlignment('center');
  sheet.getRange(L, 9, rows, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  // 行の色分け
  var list = sheet.getRange(L, 1, rows, headers.length);
  var cond = function (expr, color) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($B' + L + '),' + expr + ')')
      .setBackground(color).setRanges([list]).build();
  };
  sheet.setConditionalFormatRules([
    cond('$B' + L + '<0', DASHBOARD_COLORS.overdue),
    cond('$B' + L + '=0', DASHBOARD_COLORS.today),
    cond('$B' + L + '<=3', DASHBOARD_COLORS.within3),
    cond('$B' + L + '<=7', DASHBOARD_COLORS.within7)
  ]);

  sheet.setFrozenRows(L - 1);
  [110, 80, 110, 110, 320, 70, 90, 90, 360].forEach(function (w, i) {
    sheet.setColumnWidth(i + 1, w);
  });
  sheet.activate();
}
