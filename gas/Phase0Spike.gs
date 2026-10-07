/**
 * Phase 0 検証スパイク
 *
 * 目的：本実装に入る前に「Apps Script だけで今と同等の PDF が作れるか」を実地確認する。
 * ここが通れば残りは作業。通らなければ方式を再考する。
 *
 * 使い方：
 *   1. スプレッドシートを開く →「拡張機能」→「Apps Script」
 *   2. このファイルの中身を丸ごと貼り付けて保存
 *   3. 関数 runPhase0Spike を選んで実行（初回は権限承認のダイアログが出る）
 *   4. 実行ログに出る URL を開いて PDF を目視確認
 *   5. 続けて checkDateHeaders を実行し、実シートの日付ヘッダー構造を確認
 *
 * 検証したいのは次の4点（PDF の1ページ目にも同じことが書いてあります）：
 *   A. 日本語フォントが出るか（豆腐＝□ にならないか）
 *   B. 0.4pt の細い罫線が消えずに出るか
 *   C. @page { size: A4 landscape } が効くか
 *   D. page-break-after で意図通りに改ページされるか（全3ページになるか）
 */

var SHEET_ID = '1nyLCUISsTswtQ_nQhNkpJGe0tdO7trZUrIaO8eps6Gw';

// checkDataVolume で調べたい日付。空なら「ヘッダーにある最後の日付」を使う。
// 例: '9月18日'
var TARGET_DATE = '';

// 1ページ9人レイアウトに収まる 1選手あたりの種目数の上限。
// 実測値：全員11種目なら 198.1mm で収まり、12種目だと 211.7mm で A4横(上限200mm)を超える。
var ROWS_PER_PLAYER_LIMIT = 11;

// ---- 現行 app.py から移植した見た目の定数 ----
var C_RED    = '#DC143C';
var C_ODD    = '#FFFFFF';
var C_EVEN   = '#F2F2F2';
var C_BORDER = '#CCCCCC';
var C_DATE   = '#BBBBBB';
var C_WHITE  = '#FFFFFF';
var C_BODY   = '#1A1A1A';

var FS_TITLE = 16;    // タイトル
var FS_NAME  = 10;    // 選手名・列見出し行
var FS_TABLE = 6.5;   // データ行
var FS_DATE  = 9;     // 右上の日付

// COL_W = [4.57, 1.31, 2.16, 2.85]cm（合計 10.89cm）を比率に変換
var COL_PCT = [41.96, 12.03, 19.83, 26.17];

var CELL_V = '0.85mm';  // CELL_V 0.085cm
var CELL_H = '1.0mm';   // CELL_H 0.10cm

/*
 * 行の高さについて（ここは実測で追い込んだ箇所です）
 *
 * pptx では row.height が「行全体の高さ」ですが、HTML の td の height は
 * 「内容領域の高さ」で、padding と border はその外側に積まれます。
 * さらに日本語フォントの line-height:normal は 1.5em 前後あるため、
 * 素直に height を指定すると 1行 4.6mm のつもりが 6.5mm になり、
 * 9人×3段が A4 横に収まらなくなります（実測 228.9mm > 上限 200mm）。
 *
 * そこで height は指定せず、line-height と padding で目標の行高を作ります。
 *   データ行 : 6.5pt × 1.15 + 余白1.7mm + 罫線0.28mm ≒ 4.6mm（DATA_H 相当）
 *   見出し行 : 10pt  × 1.10 + 余白1.7mm + 罫線0.28mm ≒ 5.9mm（NAME_H 相当）
 * 見出し行を 1.0 まで詰めると "Rep" の p が欠けるため 1.1 で止めています。
 */
var LINE_H_DATA = '1.15';
var LINE_H_NAME = '1.1';

var MAX_PER_PAGE = 9;
var COLS = 3;

// Linux 上の変換エンジンで動くため Noto を後ろに置く
var FONT_STACK = "'Meiryo','Noto Sans CJK JP','Noto Sans JP','Hiragino Kaku Gothic ProN',sans-serif";


// =====================================================================
// メイン：ダミーデータで PDF を1本出す
// =====================================================================
function runPhase0Spike() {
  var players = dummyPlayers_();
  var html = buildHtml_(players);

  // ここが今回の肝。HTML の Blob を PDF に変換できるかどうか。
  var pdf = Utilities
    .newBlob(html, MimeType.HTML, 'phase0_spike.html')
    .getAs(MimeType.PDF)
    .setName('phase0_spike_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss') + '.pdf');

  var file = DriveApp.createFile(pdf);

  // HTML 側も残しておくと、崩れたときに原因を切り分けやすい
  var htmlFile = DriveApp.createFile(
    Utilities.newBlob(html, MimeType.HTML, 'phase0_spike.html')
  );

  var msg = [
    '',
    '===== Phase 0 検証スパイク 完了 =====',
    'PDF  : ' + file.getUrl(),
    'HTML : ' + htmlFile.getUrl() + '  ← 崩れたとき用の元データ',
    '',
    'PDF を開いて、1ページ目の A〜D のチェック項目を目視で確認してください。',
    '全部で 3 ページあれば D（改ページ）は OK です。',
    ''
  ].join('\n');
  Logger.log(msg);
  return file.getUrl();
}


// =====================================================================
// 実シートの日付ヘッダー構造を確認する
// （Phase 2 のデータ抽出が成立するかの事前チェック）
// =====================================================================
function checkDateHeaders() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheets = ss.getSheets();
  var out = [''];
  out.push('===== 日付ヘッダー構造チェック =====');
  out.push('スプレッドシート: ' + ss.getName());
  out.push('シート（選手）数: ' + sheets.length);
  out.push('');

  // --- 1枚目を詳細に見る ---
  var first = sheets[0];
  var detail = scanDateRow_(first);
  out.push('--- 詳細: 「' + first.getName() + '」 ---');
  if (!detail.row) {
    out.push('!! 1〜2行目に m月d日 形式のセルが見つかりませんでした。');
    out.push('   1行目の生の値（先頭20列）:');
    out.push('   ' + JSON.stringify(
      first.getRange(1, 1, 1, Math.min(20, first.getLastColumn())).getValues()[0]
    ));
    Logger.log(out.join('\n'));
    return;
  }

  out.push('日付が入っている行 : ' + detail.row + ' 行目');
  out.push('値の型             : ' + detail.valueType + '（Date 型なら書式変更に強い）');
  out.push('見つかった日付数   : ' + detail.entries.length);
  out.push('最初の日付         : ' + detail.entries[0].key +
           '  → ' + colName_(detail.entries[0].col) + '列（' + detail.entries[0].col + '）');

  // 現行 app.py の前提（L列=12 起点）と一致するか
  if (detail.entries[0].col === 12) {
    out.push('  → app.py の前提（L列起点）と一致しています。');
  } else {
    out.push('  !! app.py は L列(12) 起点を前提にしていますが、実際は ' +
             colName_(detail.entries[0].col) + '列でした。');
  }

  // --- 「4列飛ばし」が本当に守られているか ---
  var gaps = [];
  for (var i = 1; i < detail.entries.length; i++) {
    var d = detail.entries[i].col - detail.entries[i - 1].col;
    if (d !== 4) {
      gaps.push('  ' + detail.entries[i - 1].key + '(' + colName_(detail.entries[i - 1].col) + ') → ' +
                detail.entries[i].key + '(' + colName_(detail.entries[i].col) + ') が ' + d + '列');
    }
  }
  out.push('');
  if (gaps.length === 0) {
    out.push('間隔チェック : すべて 4 列間隔。現行の算術ロジックでもズレは出ていません。');
  } else {
    out.push('間隔チェック : !! 4列でない箇所が ' + gaps.length + ' 件あります。');
    out.push('               現行 app.py の get_date_col_range() は');
    out.push('               ここから先の日付で列がズレている可能性があります。');
    out.push(gaps.join('\n'));
  }

  out.push('');
  out.push('全日付一覧:');
  out.push('  ' + detail.entries.map(function (e) {
    return e.key + '(' + colName_(e.col) + ')';
  }).join(', '));

  // --- 全シートで構造が揃っているか ---
  out.push('');
  out.push('--- 全シートの整合性 ---');
  var baseSig = detail.entries.map(function (e) { return e.key + ':' + e.col; }).join(',');
  var mismatches = [];
  for (var s = 1; s < sheets.length; s++) {
    var d2 = scanDateRow_(sheets[s]);
    var sig = d2.entries.map(function (e) { return e.key + ':' + e.col; }).join(',');
    if (sig !== baseSig) {
      mismatches.push('  「' + sheets[s].getName() + '」: 日付数 ' + d2.entries.length +
                      '（基準は ' + detail.entries.length + '）' +
                      (d2.entries.length ? ' 先頭 ' + d2.entries[0].key +
                        '(' + colName_(d2.entries[0].col) + ')' : ' 日付なし'));
    }
  }
  if (mismatches.length === 0) {
    out.push('全 ' + sheets.length + ' シートで日付の並びが一致しています。');
  } else {
    out.push('!! 基準シートと異なるものが ' + mismatches.length + ' 件:');
    out.push(mismatches.join('\n'));
    out.push('（選手ごとに列がずれているなら、シート単位で列を探す実装が必要です）');
  }

  out.push('');
  Logger.log(out.join('\n'));
}


// =====================================================================
// 実データの「量」を確認する
// （1ページ9人のレイアウトに本当に収まるか、種目名が列幅に入るかの事前チェック）
// =====================================================================
function checkDataVolume() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheets = ss.getSheets();
  var out = [''];
  out.push('===== 実データ量チェック =====');

  var base = scanDateRow_(sheets[0]);
  if (!base.row) {
    Logger.log('日付行が見つかりません。先に checkDateHeaders を実行してください。');
    return;
  }
  var target = TARGET_DATE || base.entries[base.entries.length - 1].key;
  out.push('対象日: ' + target +
           (TARGET_DATE ? '（TARGET_DATE で指定）' : '（ヘッダー上の最後の日付を自動選択）'));
  out.push('');

  var counts = [];      // { name, rows }
  var noData = [];
  var noDate = [];
  var longestItem = '';
  var sample = null;

  for (var s = 0; s < sheets.length; s++) {
    var sh = sheets[s];
    var d = scanDateRow_(sh);
    if (!d.row) { noDate.push(sh.getName()); continue; }

    var col = null;
    for (var i = 0; i < d.entries.length; i++) {
      if (d.entries[i].key === target) { col = d.entries[i].col; break; }
    }
    if (col === null) { noData.push(sh.getName() + '（この日の列なし）'); continue; }

    // app.py と同じ範囲：3行目から最大99行、4列
    var lastRow = sh.getLastRow();
    if (lastRow < 3) { noData.push(sh.getName() + '（行なし）'); continue; }
    var n = Math.min(99, lastRow - 2);
    var vals = sh.getRange(3, col, n, 4).getValues();

    // app.py の dropna(subset=[セット,レップ,重量], how='all') と同じ判定
    var rows = vals.filter(function (r) {
      return String(r[1]).trim() !== '' || String(r[2]).trim() !== '' || String(r[3]).trim() !== '';
    });
    if (rows.length === 0) { noData.push(sh.getName()); continue; }

    counts.push({ name: sh.getName(), rows: rows.length });
    rows.forEach(function (r) {
      var item = String(r[0]).trim();
      if (item.length > longestItem.length) longestItem = item;
    });
    if (!sample) sample = { name: sh.getName(), rows: rows.slice(0, 4) };
  }

  if (counts.length === 0) {
    out.push('!! ' + target + ' にデータのあるシートが1つもありませんでした。');
    out.push('   TARGET_DATE を別の日付に変えて再実行してください。');
    Logger.log(out.join('\n'));
    return;
  }

  var nums = counts.map(function (c) { return c.rows; });
  var max = Math.max.apply(null, nums);
  var min = Math.min.apply(null, nums);
  var avg = nums.reduce(function (a, b) { return a + b; }, 0) / nums.length;

  out.push('データのある選手 : ' + counts.length + ' 名  → ' +
           Math.ceil(counts.length / MAX_PER_PAGE) + ' ページ（9人/ページ）');
  out.push('種目数           : 最大 ' + max + ' / 最小 ' + min + ' / 平均 ' + avg.toFixed(1));
  out.push('');

  // --- レイアウトに収まるか ---
  var overflow = counts.filter(function (c) { return c.rows > ROWS_PER_PLAYER_LIMIT; });
  if (overflow.length === 0) {
    out.push('レイアウト判定 : OK。最大 ' + max + ' 種目は上限 ' +
             ROWS_PER_PLAYER_LIMIT + ' 種目に収まります。');
  } else {
    out.push('レイアウト判定 : !! 上限 ' + ROWS_PER_PLAYER_LIMIT +
             ' 種目を超える選手が ' + overflow.length + ' 名います。');
    out.push('                 このままだと A4 横1ページに 3行 × 3列 が入りきりません。');
    out.push('                 対策：フォントを少し小さくする／行の高さを詰める／' +
             '該当選手だけ1ページに8人にする、などで調整します。');
    out.push(overflow.map(function (c) {
      return '  ' + c.name + ' : ' + c.rows + ' 種目';
    }).join('\n'));
  }

  // --- 種目名が列幅に入るか ---
  // 種目列 = 表幅の 41.96%、A4横の使用幅 287mm / 3列 → 約 94mm、左右余白を引いて約 92mm
  var itemColMM = 92 * 0.4196;
  var estMM = longestItem.length * (FS_TABLE * 25.4 / 72);  // 全角1文字 ≒ 1em として概算
  out.push('');
  out.push('最長の種目名 : 「' + longestItem + '」（' + longestItem.length + ' 文字）');
  out.push('  種目列の幅 : 約 ' + itemColMM.toFixed(1) + 'mm / 必要幅の概算 : 約 ' + estMM.toFixed(1) + 'mm');
  if (estMM > itemColMM) {
    out.push('  !! はみ出す可能性があります。折り返すか、種目列を広げるか、' +
             'フォントを下げるかの判断が要ります。');
    out.push('     （現行の PDF でこの種目名がどう出ているかも確認してください）');
  } else {
    out.push('  → 収まる見込みです。');
  }

  if (noData.length) {
    out.push('');
    out.push('この日データなし (' + noData.length + '): ' + noData.join(', '));
  }
  if (noDate.length) {
    out.push('日付行なし (' + noDate.length + '): ' + noDate.join(', ') +
             '  ← 選手以外のシートならスキップ対象です');
  }

  if (sample) {
    out.push('');
    out.push('抽出サンプル（「' + sample.name + '」の先頭4行）:');
    sample.rows.forEach(function (r) {
      out.push('  ' + JSON.stringify(r));
    });
  }

  out.push('');
  Logger.log(out.join('\n'));
}


/** 1〜2行目を走査して日付セルの位置を拾う */
function scanDateRow_(sheet) {
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  if (lastCol < 1 || lastRow < 1) return { row: null, entries: [], valueType: '-' };
  var rows = sheet.getRange(1, 1, Math.min(2, lastRow), lastCol).getValues();

  for (var r = 0; r < rows.length; r++) {
    var entries = [];
    var valueType = '-';
    for (var c = 0; c < rows[r].length; c++) {
      var key = dateKey_(rows[r][c]);
      if (key) {
        entries.push({ key: key, col: c + 1 });
        if (valueType === '-') {
          valueType = (rows[r][c] instanceof Date) ? 'Date' : 'String';
        }
      }
    }
    // 日付が2つ以上並んでいる行を「日付行」とみなす
    if (entries.length >= 2) {
      return { row: r + 1, entries: entries, valueType: valueType };
    }
  }
  return { row: null, entries: [], valueType: '-' };
}


/**
 * セルの値を「4月22日」形式のキーに正規化する。
 * 結合セルは左上にだけ値が入り、残りは空文字が返るので、
 * 非空セル＝その日の開始列になる。
 * Date 型でも文字列でも同じキーになるようにしておく。
 */
function dateKey_(v) {
  if (v instanceof Date) {
    return (v.getMonth() + 1) + '月' + v.getDate() + '日';
  }
  var m = String(v).match(/(\d+)\s*月\s*(\d+)\s*日/);
  return m ? (Number(m[1]) + '月' + Number(m[2]) + '日') : null;
}


/** 1ベースの列番号 → A, B, ... Z, AA */
function colName_(n) {
  var s = '';
  while (n > 0) {
    var rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}


// =====================================================================
// HTML 組み立て
// =====================================================================
function buildHtml_(players) {
  var today = new Date();
  var titleText = (today.getMonth() + 1) + '月' + today.getDate() + '日 トレーニング実施記録';
  var dateText = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd') + ' Workout Log';

  var totalPages = Math.ceil(players.length / MAX_PER_PAGE) + 1;
  var pages = [];
  pages.push(diagnosticsPage_(totalPages));   // 1ページ目：検証パネル
  for (var i = 0; i < players.length; i += MAX_PER_PAGE) {
    pages.push(reportPage_(
      players.slice(i, i + MAX_PER_PAGE), titleText, dateText,
      (i / MAX_PER_PAGE) + 2, totalPages
    ));
  }

  return [
    '<!DOCTYPE html>',
    '<html><head><meta charset="utf-8">',
    '<style>', css_(), '</style>',
    '</head><body>',
    pages.join('\n'),
    '</body></html>'
  ].join('\n');
}


function css_() {
  return [
    // --- C の検証対象：これが効かないと A4 横にならない ---
    '@page { size: A4 landscape; margin: 5mm; }',
    'html, body { margin:0; padding:0; }',
    'body { font-family:' + FONT_STACK + '; color:' + C_BODY + '; }',

    // --- D の検証対象：改ページ ---
    '.page { page-break-after: always; position: relative; }',
    '.page.last { page-break-after: auto; }',

    '.band { height:1.2mm; background:' + C_RED + '; margin-bottom:3mm; }',
    '.hdr { width:100%; border-collapse:collapse; margin-bottom:4mm; }',
    '.hdr td { border:none; padding:0; vertical-align:bottom; }',
    '.title { font-size:' + FS_TITLE + 'pt; font-weight:bold; color:' + C_RED + '; }',
    '.dt { font-size:' + FS_DATE + 'pt; color:' + C_DATE + '; text-align:right; }',

    // 外側グリッド：CSS Grid / Flexbox は変換エンジンで怪しいので table で組む
    '.grid { width:100%; border-collapse:collapse; table-layout:fixed; }',
    // 内側の選手テーブルの td にも効いてしまわないよう、クラスで限定する
    '.grid td.slot, .grid td.gap { vertical-align:top; padding:0 0 5mm 0; border:none; }',
    '.gap { width:1.5%; }',
    '.slot { width:32.33%; }',

    // 選手テーブル
    '.ptbl { width:100%; border-collapse:collapse; table-layout:fixed; }',
    // height は指定しない（上のコメント参照）。行高は line-height と padding で作る。
    '.ptbl td { padding:' + CELL_V + ' ' + CELL_H + '; overflow:hidden; }',
    // --- B の検証対象：0.4pt の罫線 ---
    '.ptbl td.d { border:0.4pt solid ' + C_BORDER + '; font-size:' + FS_TABLE + 'pt;' +
      ' line-height:' + LINE_H_DATA + '; }',
    '.ptbl td.h { background:' + C_RED + '; color:' + C_WHITE + '; font-weight:bold;' +
      ' font-size:' + FS_NAME + 'pt; line-height:' + LINE_H_NAME + ';' +
      ' border:0.4pt solid ' + C_RED + '; }',
    '.odd  { background:' + C_ODD + '; }',
    '.even { background:' + C_EVEN + '; }',
    '.l { text-align:left; }',
    '.c { text-align:center; }',

    // 検証パネル用
    '.diag { font-size:10pt; line-height:1.7; }',
    '.diag h1 { font-size:14pt; color:' + C_RED + '; margin:0 0 3mm 0; }',
    '.diag h2 { font-size:11pt; margin:6mm 0 2mm 0; border-left:3pt solid ' + C_RED + '; padding-left:2mm; }',
    '.note { font-size:9pt; color:#666666; }',
    '.ruler { width:100%; border-collapse:collapse; table-layout:fixed; }',
    '.ruler td { height:6mm; border:0.5pt solid #888888; font-size:7pt; text-align:center; }',
    '.ladder { border-collapse:collapse; }',
    '.ladder td { padding:1.5mm 4mm; font-size:9pt; }',
    '.pagemark { position:absolute; bottom:0; right:0; font-size:8pt; color:' + C_DATE + '; }'
  ].join('\n');
}


/** 1ページ目：4つの検証項目をその場で確認できるパネル */
function diagnosticsPage_(totalPages) {
  // 幅の校正目盛り：A4横(297mm) - 左右マージン(5mm×2) = 287mm がページ幅いっぱい
  var ruler = [];
  for (var i = 1; i <= 10; i++) {
    ruler.push('<td class="' + (i % 2 ? 'odd' : 'even') + '">' + (i * 10) + '</td>');
  }

  // 罫線の太さラダー：どの太さから消えずに出るかを見る
  var widths = [0.25, 0.4, 0.5, 0.75, 1.0];
  var ladder = widths.map(function (w) {
    return '<tr>' +
      '<td style="width:30mm;">' + w + 'pt</td>' +
      '<td style="width:120mm; border-top:' + w + 'pt solid ' + C_BORDER + ';"></td>' +
      '<td style="border:' + w + 'pt solid ' + C_BORDER + '; width:40mm;">囲み罫</td>' +
      '</tr>';
  }).join('\n');

  return [
    '<div class="page">',
    '  <div class="band"></div>',
    '  <div class="diag">',
    '  <h1>Phase 0 検証パネル — 以下の A〜D を目視で確認してください</h1>',

    '  <h2>A. 日本語フォント（□ や ? になっていないか）</h2>',
    '  <div style="font-size:' + FS_TABLE + 'pt;">6.5pt（データ行）： ' +
       'ベンチプレス／スクワット／デッドリフト／懸垂 ＃1宗山 ##60 Waters 75kg×10回 膝関節屈曲</div>',
    '  <div style="font-size:' + FS_NAME + 'pt;">10pt（選手名行）： ' +
       'ベンチプレス／スクワット ＃13藤原 濁点「ボ」半濁点「ポ」長音「ー」</div>',
    '  <div style="font-size:' + FS_TITLE + 'pt;">16pt（タイトル）： 9月20日 トレーニング実施記録</div>',
    '  <div class="note">※ 一番上の 6.5pt が潰れずに読めるかが重要です。' +
       '読めない場合はデータ行のフォントサイズを上げる必要があります。</div>',

    '  <h2>B. 罫線の太さ（現行は 0.4pt。消えている段があるか）</h2>',
    '  <table class="ladder">' + ladder + '</table>',
    '  <div class="note">※ 0.4pt の線が見えていれば現行と同じ見た目を再現できます。' +
       '消えていたら、見えている一番細い値を採用します。</div>',

    '  <h2>C. ページサイズ（A4 横になっているか）</h2>',
    '  <div style="height:3mm; background:' + C_RED + '; width:100%;"></div>',
    '  <div class="note" style="margin-top:1mm;">' +
       '↑ この赤いバーがページの左端から右端いっぱいに伸びていて、かつ用紙が<b>横長</b>なら OK です。' +
       '縦長になっていたら @page が効いていません（その場合は別の指定方法を試します）。</div>',
    '  <table class="ruler" style="margin-top:3mm;"><tr>' + ruler.join('') + '</tr></table>',
    '  <div class="note" style="margin-top:1mm;">' +
       '↑ この目盛りは合計 287mm 相当です。印刷して右端の「100」の位置が' +
       '左端から約 287mm なら、原寸で出力されています。</div>',

    '  <h2>D. 改ページ</h2>',
    '  <div>この PDF は<b>全 ' + totalPages + ' ページ</b>になるはずです' +
       '（検証パネル＋選手9人＋選手2人）。' +
       'ページ数が違う、または途中で表が分断されている場合は改ページ制御の調整が必要です。</div>',

    '  <h2>次のページからは、現行 PDF と同じ見た目を再現したモックです</h2>',
    '  <div>手元の現行 PDF と並べて、色・文字サイズ・列幅・行間を見比べてください。' +
       '気になる差分を教えていただければ Phase 3 で詰めます。</div>',
    '  </div>',
    '  <div class="pagemark">PAGE 1 / ' + totalPages + '</div>',
    '</div>'
  ].join('\n');
}


/** 2ページ目以降：現行 PDF と同じ見た目のモック */
function reportPage_(pagePlayers, titleText, dateText, pageNo, pageTotal) {
  var cells = [];
  for (var i = 0; i < pagePlayers.length; i++) {
    cells.push(playerTable_(pagePlayers[i][0], pagePlayers[i][1]));
  }

  var rows = [];
  for (var r = 0; r < Math.ceil(cells.length / COLS); r++) {
    var tds = [];
    for (var c = 0; c < COLS; c++) {
      var idx = r * COLS + c;
      if (c > 0) tds.push('<td class="gap"></td>');
      tds.push('<td class="slot">' + (idx < cells.length ? cells[idx] : '') + '</td>');
    }
    rows.push('<tr>' + tds.join('') + '</tr>');
  }

  var isLast = (pageNo === pageTotal);
  return [
    '<div class="page' + (isLast ? ' last' : '') + '">',
    '  <div class="band"></div>',
    '  <table class="hdr"><tr>',
    '    <td class="title">' + esc_(titleText) + '</td>',
    '    <td class="dt">' + esc_(dateText) + '</td>',
    '  </tr></table>',
    '  <table class="grid">' + rows.join('\n') + '</table>',
    '  <div class="pagemark">PAGE ' + pageNo + ' / ' + pageTotal + '</div>',
    '</div>'
  ].join('\n');
}


/** 選手1人分のテーブル（1行目に選手名＋列見出し、以下データ行） */
function playerTable_(name, exercises) {
  var cols = COL_PCT.map(function (p) {
    return '<col style="width:' + p + '%">';
  }).join('');

  var head = '<tr>' +
    '<td class="h l">' + esc_(name) + '</td>' +
    '<td class="h c">Set</td>' +
    '<td class="h c">Rep</td>' +
    '<td class="h l">重量</td>' +
    '</tr>';

  var body = exercises.map(function (ex, i) {
    var bg = (i % 2 === 0) ? 'odd' : 'even';
    return '<tr>' +
      '<td class="d l ' + bg + '">' + esc_(ex[0]) + '</td>' +
      '<td class="d c ' + bg + '">' + esc_(fmtNum_(ex[1])) + '</td>' +
      '<td class="d c ' + bg + '">' + esc_(fmtNum_(ex[2])) + '</td>' +
      '<td class="d l ' + bg + '">' + esc_(fmtNum_(ex[3])) + '</td>' +
      '</tr>';
  }).join('\n');

  return '<table class="ptbl"><colgroup>' + cols + '</colgroup>' + head + body + '</table>';
}


/** 3.0 -> "3"、それ以外はそのまま（app.py の fmt_num 相当） */
function fmtNum_(v) {
  var s = String(v).trim();
  if (/^-?\d+(\.0+)?$/.test(s)) {
    return String(parseInt(s, 10));
  }
  return String(v);
}


function esc_(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}


// =====================================================================
// ダミーデータ（11人＝9人+2人 で 2 ページに割れるようにしてある）
// =====================================================================
function dummyPlayers_() {
  var menu = [
    ['ベンチプレス', '3', '10', '60'],
    ['スクワット', '4', '8', '100'],
    ['デッドリフト', '3', '5', '120'],
    ['ラットプルダウン', '3', '12', '45'],
    ['ショルダープレス', '3', '10', '30'],
    ['レッグカール', '3', '15', '25'],
    ['アームカール', '3.0', '12.0', '15.0'],
    ['プランク', '3', '60秒', '自重'],
    ['メディシンボールスロー', '2', '10', '4kg']
  ];
  var names = [
    '#1宗山', '#13藤原', '##60 Waters', '#7田中', '#22佐々木',
    '#5ロドリゲス', '#31髙橋', '#9渡邊', '#44中村', '#18小林', '#2山口'
  ];
  return names.map(function (n, i) {
    // 行数を変えて、高さが違うときのレイアウト崩れも見る
    var n_rows = 4 + (i % 6);
    return [n, menu.slice(0, n_rows)];
  });
}
