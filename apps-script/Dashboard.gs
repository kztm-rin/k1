/**
 * ダッシュボードタブ：期日が近い未完了タスクの一覧
 *
 * 各 TASK表 タブを直接つないだ数式で動く（開くたび・編集のたびに最新になる）。
 * - E2 の「表示する日数」以内に期日が来るタスクと、期日を過ぎたタスクを出す
 * - 担当ごとに見出し行（■ 担当名（件数））でまとめ、その中は期日順に並べる
 * - H2 で担当を選ぶと、その人のタスクだけに絞れる（空欄なら全員）
 * - ステータスか状態が「完了」の行は出さない
 * - 期限切れ・今日・3日以内・7日以内で行に色を付ける
 * - A列の「開く」で、元のタブのその行へ飛ぶ
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

// 一覧の列：A開く B期日 C残日数 D カテゴリ E項目 F Task G担当 H状態 I重要度 J状況詳細 K リンク先（非表示）
var DASHBOARD_HEADERS = ['', '期日', '残日数', 'カテゴリ', '項目', 'Task', '担当', '状態', '重要度', '状況詳細', 'リンク先'];
var DASHBOARD_WIDTHS = [50, 110, 80, 110, 110, 320, 70, 90, 90, 360, 60];

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
    sheet.showColumns(1, sheet.getMaxColumns());
  } else {
    sheet = ss.insertSheet(DASHBOARD_SHEET_NAME, 0);
  }

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
      .requireValueInRange(ss.getRange("'" + INTEGRATED_SHEET_NAME + "'!E2:E"), true).setAllowInvalid(true).build());

  // 4〜5行目：件数
  var c = 'C' + L + ':C';
  var tiles = [
    ['期限切れ', '=COUNTIF(' + c + ',"<0")', DASHBOARD_COLORS.overdue],
    ['今日', '=COUNTIF(' + c + ',0)', DASHBOARD_COLORS.today],
    ['1〜3日後', '=COUNTIFS(' + c + ',">=1",' + c + ',"<=3")', DASHBOARD_COLORS.within3],
    ['4〜7日後', '=COUNTIFS(' + c + ',">=4",' + c + ',"<=7")', DASHBOARD_COLORS.within7],
    ['表示中の合計', '=COUNT(' + c + ')', '#eeeeee']
  ];
  tiles.forEach(function (t, i) {
    sheet.getRange(4, i + 2).setValue(t[0]).setFontColor('#666666');
    sheet.getRange(5, i + 2).setFormula(t[1]).setFontSize(20).setFontWeight('bold');
    sheet.getRange(4, i + 2, 2, 1).setBackground(t[2]).setHorizontalAlignment('center');
  });

  // 7行目〜：一覧
  var width = DASHBOARD_HEADERS.length;
  sheet.getRange(L - 1, 1, 1, width).setValues([DASHBOARD_HEADERS])
    .setFontWeight('bold').setBackground('#434343').setFontColor('#ffffff');
  sheet.getRange(L, 1).setFormula(
    '=ARRAYFORMULA(IF(K' + L + ':K="","",HYPERLINK(K' + L + ':K,"開く")))');
  updateDashboardFormula_(ss, findSourceSheets_(ss));

  var rows = sheet.getMaxRows() - L + 1;
  sheet.getRange(L, 1, rows, 1).setHorizontalAlignment('center');
  sheet.getRange(L, 2, rows, 1).setNumberFormat(DATE_FORMAT);
  sheet.getRange(L, 3, rows, 1).setNumberFormat('0"日後";0"日超過";"今日"').setHorizontalAlignment('center');
  sheet.getRange(L, 10, rows, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  // 行の色分け
  var list = sheet.getRange(L, 1, rows, width - 1);
  var daysCell = '$C' + L;
  var cond = function (expr, color) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER(' + daysCell + '),' + daysCell + expr + ')')
      .setBackground(color).setRanges([list]).build();
  };
  var groupRow = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=LEFT($B' + L + ',1)="■"')
    .setBackground('#cfe2f3').setBold(true).setRanges([list]).build();
  sheet.setConditionalFormatRules([
    groupRow,
    cond('<0', DASHBOARD_COLORS.overdue),
    cond('=0', DASHBOARD_COLORS.today),
    cond('<=3', DASHBOARD_COLORS.within3),
    cond('<=7', DASHBOARD_COLORS.within7)
  ]);

  sheet.setFrozenRows(L - 1);
  DASHBOARD_WIDTHS.forEach(function (w, i) {
    sheet.setColumnWidth(i + 1, w);
  });
  sheet.hideColumns(width);   // リンク先の列
  sheet.activate();
}

/**
 * 一覧の数式（B列から K列まで広がる）を今のタブ構成に合わせる。
 * タブの追加・削除・名前変更のときは統合タブと一緒に呼ばれる。
 */
function updateDashboardFormula_(ss, sources) {
  var sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) return;
  var anchor = sheet.getRange(DASHBOARD_LIST_ROW, 2);
  var formula = buildDashboardFormula_(sources);
  if (normalizeFormula_(anchor.getFormula()) !== normalizeFormula_(formula)) anchor.setFormula(formula);
}

function buildDashboardFormula_(sources) {
  if (sources.length === 0) return '="該当するタスクはありません"';
  // 元の列：1カテゴリ 2項目 3Task 4状況詳細 5担当 6開始 7期日 8残日数 9状態 10重要度 11URL1 12URL2 13ステータス 14リンク先
  // 出力：期日 残日数 カテゴリ 項目 Task 担当 状態 重要度 状況詳細 リンク先
  return '=ARRAYFORMULA(IFERROR(LET(' +
    'tasks,' + buildSourceStack_(sources, true) + ',' +
    'due,CHOOSECOLS(tasks,7),' +
    'hits,FILTER(tasks,ISNUMBER(due),due<=TODAY()+$E$2,' +
    'CHOOSECOLS(tasks,13)<>"完了",CHOOSECOLS(tasks,9)<>"完了",' +
    '($H$2="")+(CHOOSECOLS(tasks,5)=$H$2)),' +
    // 担当ごとにまとめ、各担当の見出し行（■ 担当名（件数））の下に期日順で並べる
    'owner,IF(CHOOSECOLS(hits,5)="","（担当なし）",CHOOSECOLS(hits,5)),' +
    'items,HSTACK(CHOOSECOLS(hits,7),CHOOSECOLS(hits,7)-TODAY(),CHOOSECOLS(hits,1,2,3),owner,CHOOSECOLS(hits,9,10,4,14)),' +
    'people,SORT(UNIQUE(owner)),' +
    'grouped,REDUCE(IF(SEQUENCE(1,10),""),people,LAMBDA(acc,p,LET(' +
    'mine,SORT(FILTER(items,CHOOSECOLS(items,6)=p),1,TRUE,4,TRUE),' +
    'VSTACK(acc,HSTACK("■ "&p&"（"&ROWS(mine)&"件）",IF(SEQUENCE(1,9),"")),mine)))),' +
    'DROP(grouped,1)' +
    '),"該当するタスクはありません"))';
}
