/**
 * ダッシュボードの一覧の数式（B列から K列まで広がる）。定数は Dashboard.gs
 * 各 TASK表 タブを直接つなぐので、タブの追加・削除・名前変更のときは統合タブと一緒に書き直す。
 */

/**
 * 一覧の数式（B列から K列まで広がる）を今のタブ構成に合わせる。
 * タブの追加・削除・名前変更のときは統合タブと一緒に呼ばれる。
 */
function updateDashboardFormula_(ss, sources) {
  var sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) return;
  // 古いレイアウトのまま数式だけ差し替えると表示が壊れるので、作り直し済みのときだけ更新する
  var lastHeader = sheet.getRange(DASHBOARD_LIST_ROW - 1, DASHBOARD_HEADERS.length).getValue();
  if (lastHeader !== DASHBOARD_HEADERS[DASHBOARD_HEADERS.length - 1]) return;
  var anchor = sheet.getRange(DASHBOARD_LIST_ROW, 2);
  var formula = buildDashboardFormula_(sources);
  if (normalizeFormula_(anchor.getFormula()) !== normalizeFormula_(formula)) anchor.setFormula(formula);
}

function buildDashboardFormula_(sources) {
  if (sources.length === 0) return '="該当するタスクはありません"';
  // 元の列：1カテゴリ 2項目 3Task 4状況詳細 5担当 6開始 7期日 8残日数 9状態 10重要度 11URL1 12URL2 13ステータス 14リンク先
  // （hits ではさらに 15 に担当名。空欄は（担当なし））
  // 出力：期日 残日数 カテゴリ 項目 Task 担当 状態 重要度 状況詳細 リンク先
  return '=ARRAYFORMULA(IFERROR(LET(' +
    'tasks,' + buildSourceStack_(sources, true) + ',' +
    'due,CHOOSECOLS(tasks,7),' +
    // 担当が空欄の行は（担当なし）として扱う。H2 が空欄か「全て」なら全員、それ以外はその担当だけ
    'who,IF(CHOOSECOLS(tasks,5)="","' + DASHBOARD_NO_OWNER + '",CHOOSECOLS(tasks,5)),' +
    'hits,FILTER(HSTACK(tasks,who),ISNUMBER(due),due<=TODAY()+$E$2,' +
    'CHOOSECOLS(tasks,13)<>"完了",CHOOSECOLS(tasks,9)<>"完了",' +
    '($H$2="")+($H$2="' + DASHBOARD_ALL + '")+(who=$H$2)),' +
    // 担当ごとにまとめ、各担当の見出し行（■ 担当名（件数））の下に期日順で並べる
    'owner,CHOOSECOLS(hits,15),' +
    'items,HSTACK(CHOOSECOLS(hits,7),CHOOSECOLS(hits,7)-TODAY(),CHOOSECOLS(hits,1,2,3),owner,CHOOSECOLS(hits,9,10,4,14)),' +
    // 担当は名前順、（担当なし）は最後（担当が1人でも動くよう、並べ替え用の列を横に付けてから並べる）
    'owners,UNIQUE(owner),' +
    'people,CHOOSECOLS(SORT(HSTACK(owners="' + DASHBOARD_NO_OWNER + '",owners),1,TRUE,2,TRUE),2),' +
    'grouped,REDUCE(IF(SEQUENCE(1,10),""),people,LAMBDA(acc,p,LET(' +
    'mine,SORT(FILTER(items,CHOOSECOLS(items,6)=p),1,TRUE,4,TRUE),' +
    'VSTACK(acc,HSTACK("■ "&p&"（"&ROWS(mine)&"件）",IF(SEQUENCE(1,9),"")),mine)))),' +
    // 先頭の空行（REDUCE の初期値）を除く。DROP は Google スプレッドシートにないので FILTER で代用する
    'FILTER(grouped,SEQUENCE(ROWS(grouped))>1)' +
    '),"該当するタスクはありません"))';
}
