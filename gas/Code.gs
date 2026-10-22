/**
 * トレーニング実施記録 PDF 生成
 *
 * スプレッドシートのメニューから、日付と掲載選手を選んで PDF を作ります。
 * サーバー（Streamlit / Cloud Run / LibreOffice）は不要で、すべてこのシート内で完結します。
 *
 * ファイル構成:
 *   このファイル … メニュー・データ取得・HTML生成・PDF化
 *   Dialog.html  … 日付と選手を選ぶダイアログ
 *
 * ■ 注意：onOpen() はプロジェクト全体で1つだけ
 *   Apps Script はプロジェクト内の全 .gs が1つのグローバル空間を共有するため、
 *   onOpen() を2つ書くと片方が黙って上書きされ、メニューが出なくなります。
 *   他の .gs を追加して、そこにも onOpen() がある場合は、
 *   下の rptBuildMenu_() の呼び出しをそちらの onOpen() に移してください:
 *
 *     function onOpen() {
 *       ...既存の処理...
 *       rptBuildMenu_();     // ← この1行を足す
 *     }
 *
 *   同じ理由で、このファイルの名前はすべて rpt で始めています。
 *   既存スクリプトと名前がぶつからないようにするためです。
 */


// =====================================================================
// 設定 — 見た目や既定値を変えたいときはここだけ触れば済みます
// =====================================================================
var RPT_CFG = {
  // PDF の保存先フォルダ名（My Drive 直下。無ければ自動で作ります）
  DRIVE_FOLDER: 'トレーニング実施記録PDF',

  // 既定で「掲載しない」選手。ダイアログから上書き保存できます
  DEFAULT_EXCLUDE: ['#1宗山', '#13藤原', '##60 Waters'],

  // 選手以外のシート名（日付行が無いシートは指定しなくても自動で除外されます）
  SKIP_SHEETS: [],

  // データの読み取り位置（1〜2行目がヘッダー、3行目からデータ）
  DATA_START_ROW: 3,
  MAX_DATA_ROWS: 99,

  COLS: 3,       // 1ページあたりの横方向の列数
  MAX_BANDS: 3   // 1ページあたりの縦方向の段数（3×3＝9人。現行 PDF と同じ密度）
};


// 色・フォント（app.py から移植）
var RPT_STYLE = {
  RED:    '#DC143C',
  ODD:    '#FFFFFF',
  EVEN:   '#F2F2F2',
  BORDER: '#CCCCCC',
  DATE:   '#BBBBBB',
  WHITE:  '#FFFFFF',
  BODY:   '#1A1A1A',

  FS_TITLE: 16,
  FS_NAME:  10,
  FS_TABLE: 6.5,
  FS_DATE:  9,

  FONT: "'Meiryo','Noto Sans CJK JP','Noto Sans JP','Hiragino Kaku Gothic ProN',sans-serif",

  // 種目 / Set / Rep / 重量 の列幅比（COL_W [4.57,1.31,2.16,2.85]cm 由来）
  COL_PCT: [41.96, 12.03, 19.83, 26.17],

  CELL_V: '0.85mm',
  CELL_H: '1.0mm',

  /*
   * 行の高さについて（Phase 0 の実測で確定した箇所です。触るときは注意）
   *
   * pptx の row.height は「行全体」の高さですが、HTML の td の height は
   * 「内容領域」の高さで、padding と border はその外側に積まれます。
   * さらに日本語フォントの line-height:normal は 1.5em 前後あるため、
   * 素直に height を指定すると 1行 4.6mm のつもりが 6.5mm になり、
   * 9人×3段が A4 横に収まりません（実測 228.9mm > 上限 200mm）。
   *
   * そこで height は指定せず、line-height と padding で行高を作ります。
   *   データ行 : 6.5pt × 1.15 + 余白1.7mm + 罫線0.28mm ≒ 4.5mm
   *   見出し行 : 10pt  × 1.10 + 余白1.7mm + 罫線0.28mm ≒ 5.75mm
   * 見出し行を 1.0 まで詰めると "Rep" の p の下が欠けるため 1.1 で止めています。
   */
  LINE_H_DATA: '1.15',
  LINE_H_NAME: '1.1'
};


// ページ送りの計算に使う実測値 (mm)。上の line-height を変えたらここも合わせること。
var RPT_M = {
  DATA_H:     4.5,    // データ1行の高さ
  NAME_H:     5.75,   // 選手名＋列見出し行の高さ
  BAND_GAP:   5.0,    // 段と段の間隔
  HEADER_H:   17.4,   // 上部の赤帯＋タイトル行が占める高さ
  GRID_AVAIL: 182.0,  // A4横の使用高さ 200mm から HEADER_H を引いた値（安全側に丸め）

  /*
   * 段の固定高さ。
   *
   * これを指定しないと、各段の高さが「その段で一番種目数の多い選手」に
   * 合わせて決まるため、1段目が短いと2段目・3段目がせり上がり、
   * ページごとに選手の縦位置がバラバラになります（pptx 版は ROW_Y で
   * 固定位置に置いていたので揃っていました）。
   *
   * 52mm は (52 - NAME_H) / DATA_H ≒ 10種目 まで吸収できる高さです。
   * これを超える段だけ下へ伸びます。
   * 通常ページは 3段 × (52 + 5) + HEADER_H = 188.4mm となり、
   * A4横の上限 200mm に対して 11.6mm の余裕が残ります。
   * この余裕は、実測値(DATA_H/NAME_H)が変換エンジン側で数%ずれても
   * ページが溢れないための保険です。小さくしすぎないこと。
   */
  BAND_H:     52.0
};


// =====================================================================
// メニュー
// =====================================================================

/**
 * スプレッドシートを開いたときに自動で走り、メニューを作ります。
 * プロジェクト内に onOpen() は1つしか置けません（ファイル冒頭の注意を参照）。
 */
function onOpen() {
  rptBuildMenu_();
  trqBuildMenu_();   // TraqImport.gs の「TRAQ取り込み」メニュー
  try { trqShowPendingAlert_(); } catch (e) { /* 自動取り込みのエラー表示。出せなくてもメニューは出す */ }
}


/**
 * メニューの組み立て本体。
 * 他のファイルの onOpen() から呼びたい場合は、この関数を呼んでください。
 * 末尾は addToUi()。Menu に create() というメソッドはありません。
 */
function rptBuildMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('レポート作成')
    .addItem('PDFを作成…', 'rptShowDialog')
    .addSeparator()
    .addItem('日付ヘッダーを確認（診断）', 'rptShowDiagnostics')
    .addToUi();
}


/** メニュー「PDFを作成…」 */
function rptShowDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Dialog')
    .setWidth(540)
    .setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'トレーニング実施記録 PDF');
}


// =====================================================================
// ダイアログ（Dialog.html）から google.script.run で呼ばれる関数
// =====================================================================

/**
 * ダイアログを開いた直後に呼ばれる。
 * 日付一覧・既定日・その日の選手を「1往復で」まとめて返します。
 * 以前は日付と選手で2回に分けており、その分だけ待ち時間が増えていました。
 */
function rptGetInitialData(forceRefresh) {
  var t0 = new Date().getTime();
  var idx = rpt_loadIndex_(forceRefresh);
  if (!idx.dates.length) {
    throw new Error('1〜2行目に「4月22日」形式の日付が見つかりませんでした。' +
                    'メニューの「日付ヘッダーを確認（診断）」で状況を確認してください。');
  }
  var def = rpt_pickDefaultDate_(idx.dates);
  var res = rptGetPlayers(def);
  res.dates = idx.dates;
  res.defaultDate = def;
  res.fast = rpt_hasSheetsApi_();
  res.elapsedMs = new Date().getTime() - t0;
  return res;
}


/** 日付が選ばれるたびに呼ばれる。その日にデータがある選手と種目数を返す */
function rptGetPlayers(dateKey) {
  var t0 = new Date().getTime();
  var players = rpt_readPlayers_(SpreadsheetApp.getActiveSpreadsheet(), dateKey);
  var excluded = rpt_loadDefaultExclude_();
  return {
    players: players.map(function (p) {
      return {
        name: p.name,
        rows: p.rows.length,
        excluded: excluded.indexOf(p.name) >= 0
      };
    }),
    estimatedPages: rpt_paginate_(players).length,
    elapsedMs: new Date().getTime() - t0
  };
}


/** 「PDFを作成」で呼ばれる本体 */
function rptGenerate(dateKey, excludeNames, saveAsDefault) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var all = rpt_readPlayers_(ss, dateKey);

  var exclude = excludeNames || [];
  var players = all.filter(function (p) { return exclude.indexOf(p.name) < 0; });
  if (!players.length) {
    throw new Error('掲載する選手が1人も選ばれていません。');
  }

  if (saveAsDefault) rpt_saveDefaultExclude_(exclude);

  var pages = rpt_paginate_(players);
  var html = rpt_buildHtml_(pages, dateKey);

  var md = rpt_parseDateKey_(dateKey);
  var fileName = md.month + '月' + md.day + '日トレーニング実施記録.pdf';

  var pdf = Utilities
    .newBlob(html, MimeType.HTML, fileName.replace(/\.pdf$/, '.html'))
    .getAs(MimeType.PDF)
    .setName(fileName);

  var folder = rpt_getOrCreateFolder_(RPT_CFG.DRIVE_FOLDER);

  // 同じ日の PDF を作り直した場合は古いものを置き換える（ゴミが溜まらないように）
  var existing = folder.getFilesByName(fileName);
  var replaced = false;
  while (existing.hasNext()) {
    existing.next().setTrashed(true);
    replaced = true;
  }

  var file = folder.createFile(pdf);

  // 種目数が多すぎて1ページに収まらない選手がいれば知らせる
  var tooTall = players.filter(function (p) {
    return RPT_M.NAME_H + p.rows.length * RPT_M.DATA_H > RPT_M.GRID_AVAIL;
  }).map(function (p) { return p.name + '（' + p.rows.length + '種目）'; });

  return {
    url: file.getUrl(),
    name: fileName,
    playerCount: players.length,
    excludedCount: exclude.length,
    pageCount: pages.length,
    replaced: replaced,
    warnings: tooTall.length
      ? ['1ページに収まらない選手がいます: ' + tooTall.join('、') +
         '。表が切れていないか PDF を確認してください。']
      : []
  };
}


// =====================================================================
// シートの読み取り
// =====================================================================

/*
 * ■ 速度について
 *
 * 以前はシート1枚ごとに getLastColumn / getLastRow / getRange を呼んでいたため、
 * 日付一覧で 4回/シート、選手読込で 6回/シート の往復が発生していました。
 * 40シートなら 160回 と 240回 で、これが「1分待つ」原因でした。
 *
 * 対策は2つです。
 *   1) Sheets API（拡張サービス）の batchGet を使い、全シート分を1往復で取る。
 *      「サービス +」から Google Sheets API を追加すると有効になります。
 *   2) 追加していない場合は従来どおりシートごとに読みますが、
 *      ヘッダーの走査結果をキャッシュして2回目以降を省きます。
 *
 * どちらでも動きます。1) を有効にすると体感が大きく変わります。
 */

/** Sheets API（拡張サービス）が有効かどうか */
function rpt_hasSheetsApi_() {
  try {
    return (typeof Sheets !== 'undefined') && !!(Sheets && Sheets.Spreadsheets);
  } catch (e) {
    return false;
  }
}


/** シート名を A1 記法で安全に使えるよう引用する */
function rpt_a1Sheet_(name) {
  return "'" + String(name).replace(/'/g, "''") + "'";
}


/** batchGet を URL 長の上限に触れないよう分割して実行し、値の配列を返す */
function rpt_batchGet_(ssId, ranges) {
  var out = [];
  var CHUNK = 50;
  for (var i = 0; i < ranges.length; i += CHUNK) {
    var part = ranges.slice(i, i + CHUNK);
    var res = Sheets.Spreadsheets.Values.batchGet(ssId, {
      ranges: part,
      valueRenderOption: 'FORMATTED_VALUE',
      majorDimension: 'ROWS'
    });
    var vrs = res.valueRanges || [];
    for (var j = 0; j < part.length; j++) {
      out.push((vrs[j] && vrs[j].values) ? vrs[j].values : []);
    }
  }
  return out;
}


/**
 * 各シートの1〜2行目を取得する。
 * 戻り値: [{ name, rows: [[...], [...]] }, ...]
 */
function rpt_fetchHeaders_(ss, sheets) {
  var names = sheets.map(function (s) { return s.getName(); });

  if (rpt_hasSheetsApi_()) {
    // 全シート分をまとめて1往復
    var ranges = names.map(function (n) { return rpt_a1Sheet_(n) + '!1:2'; });
    var values = rpt_batchGet_(ss.getId(), ranges);
    return names.map(function (n, i) { return { name: n, rows: values[i] }; });
  }

  // フォールバック：シートごとに1往復（表示値で取る）
  return sheets.map(function (s) {
    return { name: s.getName(), rows: s.getRange(1, 1, 2, s.getMaxColumns()).getDisplayValues() };
  });
}


/**
 * 日付ヘッダーの索引を作る。
 * 戻り値: { dates:[キー], dateRow, sheets:[{ name, cols:{キー:列番号} }] }
 */
function rpt_buildIndex_(ss) {
  var sheets = ss.getSheets().filter(function (s) {
    return RPT_CFG.SKIP_SHEETS.indexOf(s.getName()) < 0;
  });
  var headers = rpt_fetchHeaders_(ss, sheets);

  var seen = {};
  var dates = [];
  var dateRow = null;
  var out = [];

  for (var i = 0; i < headers.length; i++) {
    var rows = headers[i].rows || [];
    var found = null;

    for (var r = 0; r < Math.min(2, rows.length); r++) {
      var entries = [];
      var row = rows[r] || [];
      for (var c = 0; c < row.length; c++) {
        var key = rpt_dateKey_(row[c]);
        if (key) entries.push({ key: key, col: c + 1 });
      }
      // 日付が2つ以上並んでいる行を「日付行」とみなす
      if (entries.length >= 2) { found = { row: r + 1, entries: entries }; break; }
    }
    if (!found) continue;              // 日付行が無いシート＝選手シートではない

    if (dateRow === null) dateRow = found.row;
    var cols = {};
    for (var j = 0; j < found.entries.length; j++) {
      var e = found.entries[j];
      cols[e.key] = e.col;
      if (!seen[e.key]) { seen[e.key] = true; dates.push(e.key); }
    }
    out.push({ name: headers[i].name, cols: cols });
  }

  return { dates: dates, dateRow: dateRow, sheets: out };
}


/**
 * 索引をキャッシュ経由で取得する。
 * 日付ヘッダーは滅多に変わらないので、6時間保持して2回目以降の待ちを無くします。
 * 日付列を足した直後などはダイアログの「更新」から forceRefresh で取り直せます。
 */
function rpt_loadIndex_(forceRefresh) {
  var cache = CacheService.getDocumentCache();
  var KEY = 'rptIndex_v1';

  if (!forceRefresh && cache) {
    try {
      var hit = cache.get(KEY);
      if (hit) return JSON.parse(hit);
    } catch (e) { /* 壊れていたら作り直す */ }
  }

  var idx = rpt_buildIndex_(SpreadsheetApp.getActiveSpreadsheet());

  if (cache) {
    try {
      cache.put(KEY, JSON.stringify(idx), 21600);   // 6時間
    } catch (e) {
      // 100KB を超えると入らない。キャッシュ無しでも動作には影響しない
    }
  }
  return idx;
}


/** 今日以前で最も新しい日付を既定にする。無ければ最後の日付 */
function rpt_pickDefaultDate_(keys) {
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var best = null;
  for (var i = 0; i < keys.length; i++) {
    var md = rpt_parseDateKey_(keys[i]);
    var d = new Date(today.getFullYear(), md.month - 1, md.day);
    if (d <= today && (best === null || d > best.d)) best = { d: d, key: keys[i] };
  }
  return best ? best.key : keys[keys.length - 1];
}


/**
 * 指定日の全選手データを読む。
 * 戻り値: [{ name, rows: [[種目, Set, Rep, 重量], ...] }, ...]（シートの並び順）
 */
function rpt_readPlayers_(ss, dateKey) {
  var idx = rpt_loadIndex_(false);

  // その日の列を持つシートだけに絞る
  var targets = [];
  for (var i = 0; i < idx.sheets.length; i++) {
    var col = idx.sheets[i].cols[dateKey];
    if (col) targets.push({ name: idx.sheets[i].name, col: col });
  }
  if (!targets.length) return [];

  var startRow = RPT_CFG.DATA_START_ROW;
  var endRow = startRow + RPT_CFG.MAX_DATA_ROWS - 1;
  var grids;

  if (rpt_hasSheetsApi_()) {
    // 全選手分をまとめて1往復
    var ranges = targets.map(function (t) {
      return rpt_a1Sheet_(t.name) + '!' +
             rpt_colName_(t.col) + startRow + ':' + rpt_colName_(t.col + 3) + endRow;
    });
    grids = rpt_batchGet_(ss.getId(), ranges);
  } else {
    // フォールバック：シートごとに1往復
    grids = targets.map(function (t) {
      var sh = ss.getSheetByName(t.name);
      if (!sh) return [];
      var n = Math.min(RPT_CFG.MAX_DATA_ROWS, sh.getMaxRows() - startRow + 1);
      if (n <= 0) return [];
      return sh.getRange(startRow, t.col, n, 4).getDisplayValues();
    });
  }

  var result = [];
  for (var k = 0; k < targets.length; k++) {
    var vals = grids[k] || [];
    var rows = [];
    for (var r = 0; r < vals.length; r++) {
      var v = vals[r] || [];
      // app.py の dropna(subset=[セット,レップ,重量], how='all') と同じ判定
      if (rpt_cellText_(v[1]) === '' && rpt_cellText_(v[2]) === '' && rpt_cellText_(v[3]) === '') {
        continue;
      }
      rows.push([rpt_cellText_(v[0]), rpt_cellText_(v[1]),
                 rpt_cellText_(v[2]), rpt_cellText_(v[3])]);
    }
    if (rows.length) result.push({ name: String(targets[k].name).trim(), rows: rows });
  }
  return result;
}


/**
 * セルの値を「4月22日」形式のキーに正規化する。
 * 結合セルは左上にだけ値が入り、残りは空文字が返るので、
 * 非空セル＝その日の開始列になる。
 *
 * 高速化のため表示値（文字列）で読むようにしたので、セルの表示書式が
 * 「4月22日」以外に変えられても拾えるよう、次の形も受け付ける:
 *   4月22日 / 4/22 / 2024/4/22 / 2024-04-22 / Date型
 */
function rpt_dateKey_(v) {
  if (v instanceof Date) return (v.getMonth() + 1) + '月' + v.getDate() + '日';

  var s = String(v == null ? '' : v).trim();
  if (!s) return null;

  var m = s.match(/^(\d{1,2})\s*月\s*(\d{1,2})\s*日$/);
  if (m) return rpt_mdKey_(m[1], m[2]);

  // 2024/4/22、2024-04-22（年つき）
  m = s.match(/^\d{4}[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (m) return rpt_mdKey_(m[1], m[2]);

  // 4/22、04-22（年なし）
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})$/);
  if (m) return rpt_mdKey_(m[1], m[2]);

  return null;
}


/** 月日が実在する範囲なら「M月D日」を返す。そうでなければ null */
function rpt_mdKey_(mo, d) {
  var month = Number(mo), day = Number(d);
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  return month + '月' + day + '日';
}


function rpt_parseDateKey_(key) {
  var m = String(key).match(/(\d+)\s*月\s*(\d+)\s*日/);
  if (!m) throw new Error('日付の形式が不正です: ' + key);
  return { month: Number(m[1]), day: Number(m[2]) };
}


/** セルの値を表示用の文字列にする（3.0 -> "3"、Date -> m/d） */
function rpt_cellText_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return (v.getMonth() + 1) + '/' + v.getDate();
  if (typeof v === 'number') {
    return (v === Math.floor(v)) ? String(Math.floor(v)) : String(v);
  }
  var s = String(v).trim();
  if (/^-?\d+\.0+$/.test(s)) return String(parseInt(s, 10));
  return s;
}


// =====================================================================
// ページ分割
// =====================================================================

/**
 * 3人ずつの「段」に分けてページへ詰める。
 *
 * 通常は 3段＝9人で改ページし、現行 PDF と同じ 3×3 の密度を保ちます。
 * ただし種目数の多い選手がいると 3段では A4 横に収まらないため、
 * 高さも積算して、溢れる場合は段数を減らして次ページへ送ります。
 */
function rpt_paginate_(players) {
  var pages = [];
  var page = [];
  var used = 0;
  var bands = 0;

  for (var i = 0; i < players.length; i += RPT_CFG.COLS) {
    var band = players.slice(i, i + RPT_CFG.COLS);
    var maxRows = 0;
    for (var j = 0; j < band.length; j++) {
      if (band[j].rows.length > maxRows) maxRows = band[j].rows.length;
    }
    // 段は BAND_H で固定。それを超える種目数の段だけ実際の高さで数える
    var bandH = Math.max(RPT_M.BAND_H, RPT_M.NAME_H + maxRows * RPT_M.DATA_H) + RPT_M.BAND_GAP;

    if (page.length && (bands >= RPT_CFG.MAX_BANDS || used + bandH > RPT_M.GRID_AVAIL)) {
      pages.push(page);
      page = [];
      used = 0;
      bands = 0;
    }
    page = page.concat(band);
    used += bandH;
    bands++;
  }
  if (page.length) pages.push(page);
  return pages;
}


// =====================================================================
// HTML 生成
// =====================================================================
function rpt_buildHtml_(pages, dateKey) {
  var md = rpt_parseDateKey_(dateKey);
  var titleText = md.month + '月' + md.day + '日 トレーニング実施記録';
  var dateText = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd') + ' Workout Log';

  var body = [];
  for (var i = 0; i < pages.length; i++) {
    body.push(rpt_renderPage_(pages[i], titleText, dateText, i === pages.length - 1));
  }

  return [
    '<!DOCTYPE html>',
    '<html><head><meta charset="utf-8"><title>' + rpt_esc_(titleText) + '</title>',
    '<style>', rpt_css_(), '</style>',
    '</head><body>',
    body.join('\n'),
    '</body></html>'
  ].join('\n');
}


function rpt_css_() {
  var S = RPT_STYLE;
  return [
    '@page { size: A4 landscape; margin: 5mm; }',

    /*
     * 背景色を PDF に残すための指定。
     * 印刷系のレンダラは既定で「背景グラフィック」を省くため、これが無いと
     * 罫線と文字は出るのに背景色だけが白く抜けます（ヘッダー行や縞模様が消える）。
     * 併せて td には bgcolor 属性も付けてあります（rpt_playerTable_ 参照）。
     */
    '* { -webkit-print-color-adjust: exact; print-color-adjust: exact;' +
      ' color-adjust: exact; }',

    'html, body { margin:0; padding:0; }',
    'body { font-family:' + S.FONT + '; color:' + S.BODY + '; }',

    '.page { page-break-after: always; }',
    '.page.last { page-break-after: auto; }',

    // 上端の赤帯。div + background だと落ちることがあるため table + bgcolor で描く
    '.band { width:100%; border-collapse:collapse; margin-bottom:3mm; }',
    '.band td { height:1.2mm; padding:0; border:none;' +
      ' background-color:' + S.RED + '; font-size:0; line-height:0; }',

    '.hdr { width:100%; border-collapse:collapse; margin-bottom:4mm; }',
    '.hdr td { border:none; padding:0; vertical-align:bottom; }',
    '.title { font-size:' + S.FS_TITLE + 'pt; font-weight:bold; color:' + S.RED + '; }',
    '.dt { font-size:' + S.FS_DATE + 'pt; color:' + S.DATE + '; text-align:right; }',

    // 外側グリッドは table で組む（CSS Grid / Flexbox は変換エンジンで不確実なため）。
    // 内側の選手テーブルの td に効かないよう、クラスで限定すること。
    '.grid { width:100%; border-collapse:collapse; table-layout:fixed; }',
    // height は「最低の高さ」として効く。これで各段の開始位置が毎ページ揃う
    '.grid td.slot, .grid td.gap { vertical-align:top; height:' + RPT_M.BAND_H + 'mm;' +
      ' padding:0 0 ' + RPT_M.BAND_GAP + 'mm 0; border:none; }',
    '.gap { width:1.5%; }',
    '.slot { width:32.33%; }',

    '.ptbl { width:100%; border-collapse:collapse; table-layout:fixed; }',
    // height は指定しない。行高は line-height と padding で作る（上部コメント参照）
    '.ptbl td { padding:' + S.CELL_V + ' ' + S.CELL_H + '; overflow:hidden; }',
    '.ptbl td.d { border:0.4pt solid ' + S.BORDER + '; font-size:' + S.FS_TABLE + 'pt;' +
      ' line-height:' + S.LINE_H_DATA + '; }',
    '.ptbl td.h { background-color:' + S.RED + '; color:' + S.WHITE + '; font-weight:bold;' +
      ' font-size:' + S.FS_NAME + 'pt; line-height:' + S.LINE_H_NAME + ';' +
      ' border:0.4pt solid ' + S.RED + '; }',
    '.ptbl td.odd  { background-color:' + S.ODD + '; }',
    '.ptbl td.even { background-color:' + S.EVEN + '; }',
    '.l { text-align:left; }',
    '.c { text-align:center; }'
  ].join('\n');
}


function rpt_renderPage_(pagePlayers, titleText, dateText, isLast) {
  var rows = [];
  for (var i = 0; i < pagePlayers.length; i += RPT_CFG.COLS) {
    var tds = [];
    for (var c = 0; c < RPT_CFG.COLS; c++) {
      var p = pagePlayers[i + c];
      if (c > 0) tds.push('<td class="gap"></td>');
      tds.push('<td class="slot">' + (p ? rpt_playerTable_(p.name, p.rows) : '') + '</td>');
    }
    rows.push('<tr>' + tds.join('') + '</tr>');
  }

  return [
    '<div class="page' + (isLast ? ' last' : '') + '">',
    '  <table class="band"><tr><td bgcolor="' + RPT_STYLE.RED + '"></td></tr></table>',
    '  <table class="hdr"><tr>',
    '    <td class="title">' + rpt_esc_(titleText) + '</td>',
    '    <td class="dt">' + rpt_esc_(dateText) + '</td>',
    '  </tr></table>',
    '  <table class="grid">' + rows.join('\n') + '</table>',
    '</div>'
  ].join('\n');
}


/** 選手1人分のテーブル（1行目に選手名＋列見出し、以下データ行） */
function rpt_playerTable_(name, exercises) {
  var cols = '';
  for (var i = 0; i < RPT_STYLE.COL_PCT.length; i++) {
    cols += '<col style="width:' + RPT_STYLE.COL_PCT[i] + '%">';
  }

  // bgcolor 属性は CSS の background-color が落ちたときの保険。
  // 印刷系レンダラは背景色を省くことがあるが、この属性は尊重されやすい。
  var hb = ' bgcolor="' + RPT_STYLE.RED + '"';
  var head = '<tr>' +
    '<td class="h l"' + hb + '>' + rpt_esc_(name) + '</td>' +
    '<td class="h c"' + hb + '>Set</td>' +
    '<td class="h c"' + hb + '>Rep</td>' +
    '<td class="h l"' + hb + '>重量</td>' +
    '</tr>';

  var body = '';
  for (var r = 0; r < exercises.length; r++) {
    var ex = exercises[r];
    var odd = (r % 2 === 0);
    var bg = odd ? 'odd' : 'even';
    var db = ' bgcolor="' + (odd ? RPT_STYLE.ODD : RPT_STYLE.EVEN) + '"';
    body += '<tr>' +
      '<td class="d l ' + bg + '"' + db + '>' + rpt_esc_(ex[0]) + '</td>' +
      '<td class="d c ' + bg + '"' + db + '>' + rpt_esc_(ex[1]) + '</td>' +
      '<td class="d c ' + bg + '"' + db + '>' + rpt_esc_(ex[2]) + '</td>' +
      '<td class="d l ' + bg + '"' + db + '>' + rpt_esc_(ex[3]) + '</td>' +
      '</tr>';
  }

  return '<table class="ptbl"><colgroup>' + cols + '</colgroup>' + head + body + '</table>';
}


function rpt_esc_(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}


// =====================================================================
// 既定の除外選手（ドキュメントプロパティに保存。コードを編集せずに変えられる）
// =====================================================================
function rpt_loadDefaultExclude_() {
  var raw = PropertiesService.getDocumentProperties().getProperty('RPT_DEFAULT_EXCLUDE');
  if (raw === null) return RPT_CFG.DEFAULT_EXCLUDE.slice();
  if (raw === '') return [];
  return raw.split('\n');
}


function rpt_saveDefaultExclude_(names) {
  PropertiesService.getDocumentProperties()
    .setProperty('RPT_DEFAULT_EXCLUDE', (names || []).join('\n'));
}


// =====================================================================
// Drive
// =====================================================================
function rpt_getOrCreateFolder_(name) {
  var it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}


// =====================================================================
// 診断：日付ヘッダーの状態をダイアログで表示する
// =====================================================================
function rptShowDiagnostics() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var t0 = new Date().getTime();
  var idx = rpt_loadIndex_(true);          // 診断なので必ず取り直す
  var elapsed = new Date().getTime() - t0;
  var lines = [];

  lines.push('読み取り方式 : ' + (rpt_hasSheetsApi_()
    ? 'Sheets API（高速・全シートを1往復）'
    : 'シートごとの読み取り（低速）\n' +
      '  → エディタ左の「サービス +」から Google Sheets API を追加すると高速化します'));
  lines.push('読み取り時間 : ' + (elapsed / 1000).toFixed(1) + ' 秒');
  lines.push('');

  var withDate = idx.sheets.length;
  var allSheets = ss.getSheets().length;
  var noDate = allSheets - withDate;

  if (!withDate) {
    lines.push('1〜2行目に「4月22日」形式の日付が見つかりませんでした。');
    lines.push('日付の入っている行と書式を確認してください。');
  } else {
    var base = idx.sheets[0];
    var keys = [];
    for (var k in base.cols) { if (base.cols.hasOwnProperty(k)) keys.push(k); }
    keys.sort(function (a, b) { return base.cols[a] - base.cols[b]; });

    lines.push('選手シート : ' + withDate + ' 枚（日付行なし ' + noDate + ' 枚は自動で除外）');
    lines.push('基準シート : ' + base.name);
    lines.push('日付の行   : ' + idx.dateRow + ' 行目');
    lines.push('日付の数   : ' + keys.length);
    lines.push('最初の日付 : ' + keys[0] + '（' + rpt_colName_(base.cols[keys[0]]) + '列）');

    var gaps = [];
    for (var j = 1; j < keys.length; j++) {
      var diff = base.cols[keys[j]] - base.cols[keys[j - 1]];
      if (diff !== 4) {
        gaps.push(keys[j - 1] + ' → ' + keys[j] + ' が ' + diff + '列');
      }
    }
    lines.push('');
    lines.push(gaps.length === 0
      ? '列間隔 : すべて4列間隔です。'
      : '列間隔 : 4列でない箇所が ' + gaps.length + ' 件あります（このツールは自動で追従します）:\n  ' +
        gaps.join('\n  '));

    // 基準シートと日付の並びが違うシートを洗い出す
    var mismatch = [];
    for (var s = 1; s < idx.sheets.length; s++) {
      var n = 0;
      for (var kk in idx.sheets[s].cols) {
        if (idx.sheets[s].cols.hasOwnProperty(kk)) n++;
      }
      if (n !== keys.length) mismatch.push(idx.sheets[s].name + '（' + n + '日分）');
    }
    lines.push('');
    lines.push(mismatch.length === 0
      ? '整合性 : 全シートで日付の数が一致しています。'
      : '整合性 : 基準と日付数が違うシートが ' + mismatch.length + ' 枚あります:\n  ' +
        mismatch.join(', ') + '\n  （シートごとに列を探しているので、出力は正しく行われます）');
  }

  SpreadsheetApp.getUi().alert('日付ヘッダーの診断', lines.join('\n'),
                               SpreadsheetApp.getUi().ButtonSet.OK);
}


/** 1ベースの列番号 → A, B, ... Z, AA */
function rpt_colName_(n) {
  var s = '';
  while (n > 0) {
    var rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
