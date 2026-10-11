/**
 * ダッシュボードタブの画面づくり。一覧の数式は DashboardFormula.gs
 * E2：表示する日数、H2：担当（全て／（担当なし）／各担当者。選択肢は非表示の L 列）
 */

var DASHBOARD_SHEET_NAME = 'ダッシュボード';
var DASHBOARD_DEFAULT_DAYS = 14;
var DASHBOARD_LIST_ROW = 8;   // 一覧の1行目（見出しはその1行上）
var DASHBOARD_ALL = '全て';
var DASHBOARD_NO_OWNER = '（担当なし）';
var DASHBOARD_OPTIONS_COL = 12;   // L列：担当の選択肢（非表示）

var DASHBOARD_COLORS = { overdue: '#f4cccc', today: '#fce5cd', within3: '#fff2cc', within7: '#d9ead3' };

// 一覧の列：A開く B期日 C残日数 Dカテゴリ E項目 F Task G担当 H状態 I重要度 J状況詳細 Kリンク先（非表示）
var DASHBOARD_HEADERS = ['', '期日', '残日数', 'カテゴリ', '項目', 'Task', '担当', '状態', '重要度', '状況詳細', 'リンク先'];
var DASHBOARD_WIDTHS = [50, 110, 80, 110, 110, 320, 70, 90, 90, 360, 60];

function setupDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  var days = DASHBOARD_DEFAULT_DAYS;
  var person = DASHBOARD_ALL;
  if (sheet) {
    days = Number(sheet.getRange('E2').getValue()) || DASHBOARD_DEFAULT_DAYS;
    person = sheet.getRange('H2').getValue() || DASHBOARD_ALL;
    sheet.clear();
    sheet.clearConditionalFormatRules();
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
  } else {
    sheet = ss.insertSheet(DASHBOARD_SHEET_NAME, 0);
  }
  if (sheet.getMaxColumns() < DASHBOARD_OPTIONS_COL) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), DASHBOARD_OPTIONS_COL - sheet.getMaxColumns());
  }
  sheet.showColumns(1, sheet.getMaxColumns());

  var L = DASHBOARD_LIST_ROW;

  // 1〜2行目：タイトルと設定
  sheet.getRange('A1').setValue('期日が近いタスク').setFontSize(16).setFontWeight('bold');
  sheet.getRange('A2:H2').setValues([['今日', '=TODAY()', '', '表示する日数', days, '', '担当', person]]);
  sheet.getRange('B2').setNumberFormat(DATE_FORMAT);
  sheet.getRange('A2:H2').setFontColor('#666666');
  sheet.getRange('E2').setBackground('#fff8e1').setFontColor('#000000')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(0, 365).build());
  // 担当の選択肢（L列）
  var owners = "'" + INTEGRATED_SHEET_NAME + "'!E2:E";
  sheet.getRange(1, DASHBOARD_OPTIONS_COL).setFormula(
    '=VSTACK("' + DASHBOARD_ALL + '","' + DASHBOARD_NO_OWNER + '",' +
    'IFERROR(SORT(UNIQUE(FILTER(' + owners + ',' + owners + '<>""))),""))');
  sheet.getRange('H2').setBackground('#fff8e1').setFontColor('#000000')
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInRange(sheet.getRange(1, DASHBOARD_OPTIONS_COL, sheet.getMaxRows(), 1), true)
      .setAllowInvalid(true).build());

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
  sheet.hideColumns(width);                   // K列：リンク先
  sheet.hideColumns(DASHBOARD_OPTIONS_COL);   // L列：担当の選択肢
  sheet.activate();
}
