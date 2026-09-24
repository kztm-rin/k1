/**
 * TASK表：完了タスクのアーカイブ
 *
 * 各プロジェクトタブから「ステータス」が「完了」の行を完了タブへ移し、
 * 元のタブからは削除する（記載ルール 7．完了したタスクの扱い）。
 * 「状態」だけが「完了」の行は移さない。
 * 完了タブの「元タブ」列には移動元のタブ名を、そのタブへのリンク付きで入れる。
 *
 * 対象になるタブ：A列の見出しが「項目」、続く列が「Task」…「ステータス」の
 * 12列構成になっているタブ。全体タブ（カテゴリ列あり）や記載ルールタブ、
 * Google ToDo 連携タブ（リストID列あり）は構成が違うので自動的に対象外になる。
 */

var ARCHIVE_SHEET_NAME = '完了';

// 構成が同じでも移動の対象にしないタブ
var EXCLUDED_SHEETS = ['記載ルール', '全体', ARCHIVE_SHEET_NAME];

// プロジェクトタブの列構成（記載ルール 4．各列の書き方）
var TASK_HEADERS = ['項目', 'Task', '状況詳細', '担当', '開始', '期日', '残日数', '状態', '重要度', 'URL1', 'URL2', 'ステータス'];
var DONE = '完了';

// 見出し行を探す範囲（上から何行目まで）
var HEADER_SEARCH_ROWS = 5;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('タスク表')
    .addItem('完了タスクを完了タブへ移す', 'archiveCompletedTasksFromMenu')
    .addItem('移す対象を確認する（移動しない）', 'previewCompletedTasks')
    .addSeparator()
    .addItem('毎週月曜の自動アーカイブを設定', 'installWeeklyTrigger')
    .addItem('自動アーカイブを解除', 'removeWeeklyTrigger')
    .addToUi();
}

function archiveCompletedTasksFromMenu() {
  var ui = SpreadsheetApp.getUi();
  var targets = findCompletedRows_(SpreadsheetApp.getActiveSpreadsheet());
  var total = countRows_(targets);
  if (total === 0) {
    ui.alert('完了タスクはありません。');
    return;
  }
  var answer = ui.alert(
    '完了タスクのアーカイブ',
    summarize_(targets) + '\n\n計 ' + total + ' 行を「' + ARCHIVE_SHEET_NAME + '」タブへ移します。よろしいですか？',
    ui.ButtonSet.OK_CANCEL);
  if (answer !== ui.Button.OK) return;

  var moved = archiveCompletedTasks();
  ui.alert(summarize_(targets) + '\n\n計 ' + moved + ' 行を「' + ARCHIVE_SHEET_NAME + '」タブへ移しました。');
}

function previewCompletedTasks() {
  var targets = findCompletedRows_(SpreadsheetApp.getActiveSpreadsheet());
  var total = countRows_(targets);
  SpreadsheetApp.getUi().alert(total === 0
    ? '完了タスクはありません。'
    : summarize_(targets) + '\n\n計 ' + total + ' 行が移動の対象です。');
}

/**
 * 完了行を完了タブへ移す。トリガーからも呼ばれる。
 * @return {number} 移した行数
 */
function archiveCompletedTasks() {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30 * 1000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var targets = findCompletedRows_(ss);
    if (countRows_(targets) === 0) return 0;

    var archive = getOrCreateArchiveSheet_(ss);
    var archivedAt = new Date();
    var moved = 0;

    targets.forEach(function (t) {
      // 先に全行を完了タブへ写してから、元タブを下から消す
      t.rows.forEach(function (row) {
        copyRowToArchive_(t.sheet, row, t.width, archive, archivedAt);
      });
      deleteRows_(t.sheet, t.rows);
      moved += t.rows.length;
    });

    SpreadsheetApp.flush();
    return moved;
  } finally {
    lock.releaseLock();
  }
}

/**
 * @return {{sheet: Sheet, width: number, rows: number[]}[]} タブごとの完了行（1始まり、昇順）
 */
function findCompletedRows_(ss) {
  var result = [];
  ss.getSheets().forEach(function (sheet) {
    // グラフだけのシート（OBJECT）やデータコネクタのシートは行を扱えないので飛ばす
    if (sheet.getType() !== SpreadsheetApp.SheetType.GRID) return;
    if (EXCLUDED_SHEETS.indexOf(sheet.getName()) !== -1) return;

    var headerRow = findHeaderRow_(sheet);
    if (!headerRow) return;

    var lastRow = sheet.getLastRow();
    if (lastRow <= headerRow) return;

    var width = TASK_HEADERS.length;
    var values = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, width).getDisplayValues();
    var statusCol = TASK_HEADERS.indexOf('ステータス');

    var rows = [];
    values.forEach(function (v, i) {
      if (String(v[statusCol]).trim() === DONE) {
        rows.push(headerRow + 1 + i);
      }
    });
    if (rows.length) result.push({ sheet: sheet, width: width, rows: rows });
  });
  return result;
}

function findHeaderRow_(sheet) {
  var rows = Math.min(HEADER_SEARCH_ROWS, sheet.getLastRow());
  if (rows === 0 || sheet.getLastColumn() < TASK_HEADERS.length) return 0;
  var values = sheet.getRange(1, 1, rows, TASK_HEADERS.length).getDisplayValues();
  for (var i = 0; i < values.length; i++) {
    var matches = TASK_HEADERS.every(function (h, j) {
      return String(values[i][j]).trim() === h;
    });
    if (matches) return i + 1;
  }
  return 0;
}

function getOrCreateArchiveSheet_(ss) {
  var sheet = ss.getSheetByName(ARCHIVE_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(ARCHIVE_SHEET_NAME, ss.getSheets().length);
  if (sheet.getLastRow() === 0) {
    var headers = ['元タブ', 'アーカイブ日'].concat(TASK_HEADERS);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 1行を完了タブの末尾へ写す。書式とリンクを残すため copyTo を使い、
 * 残日数などの数式は移動先で値に置き換える（HYPERLINK は残す）。
 */
function copyRowToArchive_(source, row, width, archive, archivedAt) {
  var destRow = archive.getLastRow() + 1;
  archive.getRange(destRow, 1, 1, 2).setValues([[source.getName(), archivedAt]]);
  archive.getRange(destRow, 1).setRichTextValue(SpreadsheetApp.newRichTextValue()
    .setText(source.getName())
    .setLinkUrl('#gid=' + source.getSheetId())
    .build());
  archive.getRange(destRow, 2).setNumberFormat('yy/m/d');

  var src = source.getRange(row, 1, 1, width);
  var dest = archive.getRange(destRow, 3, 1, width);
  src.copyTo(dest, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  var formulas = src.getFormulas()[0];
  var values = src.getValues()[0];
  formulas.forEach(function (f, j) {
    if (f && !/^=\s*HYPERLINK\(/i.test(f)) dest.getCell(1, j + 1).setValue(values[j]);
  });
}

/** 行番号の昇順リストを受け取り、連続した塊ごとに下から削除する */
function deleteRows_(sheet, rows) {
  var i = rows.length - 1;
  while (i >= 0) {
    var end = rows[i];
    var start = end;
    while (i > 0 && rows[i - 1] === start - 1) {
      i--;
      start--;
    }
    sheet.deleteRows(start, end - start + 1);
    i--;
  }
}

function countRows_(targets) {
  return targets.reduce(function (n, t) { return n + t.rows.length; }, 0);
}

function summarize_(targets) {
  return targets.map(function (t) {
    return '・' + t.sheet.getName() + '：' + t.rows.length + ' 行';
  }).join('\n');
}

function installWeeklyTrigger() {
  removeWeeklyTrigger();
  ScriptApp.newTrigger('archiveCompletedTasks')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .inTimezone('Asia/Tokyo')
    .create();
  SpreadsheetApp.getActiveSpreadsheet().toast('毎週月曜 7時台に完了タスクを自動で移します。');
}

function removeWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'archiveCompletedTasks') ScriptApp.deleteTrigger(t);
  });
}
