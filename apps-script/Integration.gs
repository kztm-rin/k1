/**
 * TASK表 統合タブの自動化
 *
 * 「TASK表」で始まるプロジェクトタブを全部つないだ数式を、統合タブの A2 に1つだけ置く。
 * - 行の追加・削除・編集は数式がそのまま拾うので、手で範囲を直す必要はない
 * - カテゴリ列はタブ名から自動で入る（「TASK表 」を除いた名前。CATEGORY_ALIASES で言い換え可）
 * - 期日の入っている行だけを、期日の早い順に並べる（記載ルール 1．期日）
 * - 残日数は 期日 − 今日 で計算する
 * - 開始が空の行は、統合タイムラインに載るよう 開始 = 期日 として出す
 * タブの追加・削除・名前変更のときだけ、スクリプトが数式を書き直す。
 */

var INTEGRATED_SHEET_NAME = 'TASK表 統合';
var SOURCE_PREFIX = 'TASK表';

// 「TASK表」で始まっても統合に入れないタブ
var INTEGRATION_EXCLUDED = ['TASK表 統合', 'TASK表 Archive'];

// タブ名とカテゴリ名を変えたいとき（タブ名から「TASK表」を除いた名前 → カテゴリ）
var CATEGORY_ALIASES = {
  '代々木駅前': '代々木',
  '新規案件': '新規',
  '運営（既存媒体運営制作管理）': '運営'
};

var INTEGRATED_HEADERS = ['カテゴリ'].concat(TASK_HEADERS);
var DATE_FORMAT = 'yy"/"m"/"d"("ddd")"';

function rebuildIntegrationFromMenu() {
  var sources = refreshIntegration_(true);
  SpreadsheetApp.getUi().alert(
    '「' + INTEGRATED_SHEET_NAME + '」を作り直しました。\n\n対象のタブ：\n' +
    sources.map(function (s) { return '・' + s.name + '（' + s.category + '）'; }).join('\n'));
}

/**
 * 統合タブの数式を今のタブ構成に合わせる。
 * @param {boolean} force 数式が同じでも書き直す
 * @return {{name: string, category: string}[]} 統合に入れたタブ
 */
function refreshIntegration_(force) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(INTEGRATED_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(INTEGRATED_SHEET_NAME);

  var sources = findSourceSheets_(ss);
  var formula = buildIntegrationFormula_(sources);
  var anchor = sheet.getRange('A2');
  if (!force && normalizeFormula_(anchor.getFormula()) === normalizeFormula_(formula)) return sources;

  var lock = LockService.getDocumentLock();
  lock.waitLock(30 * 1000);
  try {
    sheet.getRange(1, 1, 1, INTEGRATED_HEADERS.length).setValues([INTEGRATED_HEADERS]);
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) sheet.getRange(2, 1, lastRow - 1, INTEGRATED_HEADERS.length).clearContent();
    anchor.setFormula(formula);

    var rows = sheet.getMaxRows() - 1;
    sheet.getRange(2, 6, rows, 2).setNumberFormat(DATE_FORMAT);   // 開始・期日
    sheet.getRange(2, 8, rows, 1).setNumberFormat('0');           // 残日数
  } finally {
    lock.releaseLock();
  }
  return sources;
}

function findSourceSheets_(ss) {
  var excluded = INTEGRATION_EXCLUDED.map(normalizeSheetName_);
  var sources = [];
  ss.getSheets().forEach(function (sheet) {
    if (sheet.getType() !== SpreadsheetApp.SheetType.GRID) return;
    var name = normalizeSheetName_(sheet.getName());
    if (name.indexOf(SOURCE_PREFIX) !== 0 || excluded.indexOf(name) !== -1) return;

    var headerRow = findHeaderRow_(sheet);
    if (!headerRow) return;

    var base = name.slice(SOURCE_PREFIX.length).trim();
    sources.push({
      sheet: sheet,
      name: sheet.getName(),
      category: CATEGORY_ALIASES[base] || base,
      firstDataRow: headerRow + 1
    });
  });
  return sources;
}

function buildIntegrationFormula_(sources) {
  if (sources.length === 0) return '=""';
  var parts = sources.map(function (s) {
    var ref = "'" + s.name.replace(/'/g, "''") + "'!A" + s.firstDataRow + ':L';
    return 'LET(tab,' + ref + ',HSTACK(IF(SEQUENCE(ROWS(tab)),"' + s.category.replace(/"/g, '""') + '"),tab))';
  });
  // 列：1カテゴリ 2項目 3Task 4状況詳細 5担当 6開始 7期日 8残日数 9状態 10重要度 11URL1 12URL2 13ステータス
  return '=ARRAYFORMULA(IFERROR(LET(' +
    'allrows,VSTACK(' + parts.join(',') + '),' +
    'dated,FILTER(allrows,ISNUMBER(CHOOSECOLS(allrows,7))),' +
    'sorted,SORT(dated,7,TRUE),' +
    'due,CHOOSECOLS(sorted,7),' +
    'startdate,CHOOSECOLS(sorted,6),' +
    'HSTACK(CHOOSECOLS(sorted,1,2,3,4,5),IF(startdate="",due,startdate),due,due-TODAY(),CHOOSECOLS(sorted,9,10,11,12,13))' +
    '),""))';
}

function normalizeSheetName_(name) {
  return String(name).replace(/[\s　]+/g, ' ').trim();
}

function normalizeFormula_(f) {
  return String(f).replace(/\s+/g, '');
}
