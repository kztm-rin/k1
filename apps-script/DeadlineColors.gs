/**
 * 各 TASK表 タブと統合タブの行を、期日の近さでダッシュボードと同じ色に塗る（条件付き書式）。
 *
 * 期限切れ・今日・1〜3日後・4〜7日後の4段階。ステータスか状態が「完了」の行は塗らない。
 * 既存の条件付き書式は残し、このスクリプトが付けたものだけを入れ替える。
 */

// 期日・状態・ステータスの列（TASK_HEADERS / INTEGRATED_HEADERS 上の位置から求める）
function deadlineColumns_(headers) {
  var letter = function (name) {
    return columnLetter_(headers.indexOf(name) + 1);
  };
  return { due: letter('期日'), state: letter('状態'), status: letter('ステータス') };
}

function applyDeadlineColorsFromMenu() {
  var result = applyDeadlineColors_(false);
  var msg = result.done.length + ' タブに期日の色を付けました。';
  if (result.failed.length) msg += '\n\n付けられなかったタブ：\n' + result.failed.join('\n');
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * @param {boolean} onlyMissing まだ色の設定がないタブだけに付ける（開いたとき用）
 */
function applyDeadlineColors_(onlyMissing) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targets = findSourceSheets_(ss).map(function (s) {
    return { sheet: s.sheet, firstRow: s.firstDataRow, headers: TASK_HEADERS };
  });
  var integrated = ss.getSheetByName(INTEGRATED_SHEET_NAME);
  if (integrated) targets.push({ sheet: integrated, firstRow: 2, headers: INTEGRATED_HEADERS });

  var result = { done: [], failed: [] };
  targets.forEach(function (t) {
    try {
      var rules = t.sheet.getConditionalFormatRules();
      var others = rules.filter(function (r) { return !isDeadlineRule_(r); });
      if (onlyMissing && others.length < rules.length) return;

      t.sheet.setConditionalFormatRules(buildDeadlineRules_(t.sheet, t.firstRow, t.headers).concat(others));
      result.done.push(t.sheet.getName());
    } catch (e) {
      console.warn(t.sheet.getName(), e);
      result.failed.push('・' + t.sheet.getName() + '：' + e.message);
    }
  });
  return result;
}

function buildDeadlineRules_(sheet, firstRow, headers) {
  var c = deadlineColumns_(headers);
  var r = firstRow;
  var range = sheet.getRange(r, 1, sheet.getMaxRows() - r + 1, headers.length);
  var days = '$' + c.due + r + '-TODAY()';
  var rule = function (expr, color) {
    var formula = '=AND(ISNUMBER($' + c.due + r + '),' +
      '$' + c.status + r + '<>"完了",$' + c.state + r + '<>"完了",' + days + expr + ')';
    return SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula).setBackground(color).setRanges([range]).build();
  };
  return [
    rule('<0', DASHBOARD_COLORS.overdue),
    rule('=0', DASHBOARD_COLORS.today),
    rule('<=3', DASHBOARD_COLORS.within3),
    rule('<=7', DASHBOARD_COLORS.within7)
  ];
}

function isDeadlineRule_(rule) {
  var cond = rule.getBooleanCondition();
  if (!cond || cond.getCriteriaType() !== SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) return false;
  var f = String(cond.getCriteriaValues()[0]).replace(/\s+/g, '');
  return /^=AND\(ISNUMBER\(\$[A-Z]+\d+\),\$[A-Z]+\d+<>"完了",\$[A-Z]+\d+<>"完了",\$[A-Z]+\d+-TODAY\(\)(<0|=0|<=3|<=7)\)$/.test(f);
}

function columnLetter_(n) {
  var s = '';
  while (n > 0) {
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
