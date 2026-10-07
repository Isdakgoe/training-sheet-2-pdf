/**
 * TRAQ → スプレッドシート 取り込み
 *
 * TRAQ（traq.drivelinebaseball.com）の「Input data」ポップアップに入力された
 * トレーニング結果を読み、選手ごとのシートの日付列（種目/セット/レップ/重量）へ転記します。
 *
 * ■ 仕組み（ブラウザ操作はせず、TRAQ が内部で使っている通信をそのまま呼びます）
 *   1. POST /login                              … ログイン（CSRFトークン付きフォーム）
 *   2. GET  /athletes/view/{選手ID}              … 一覧の <tr onclick="strengthInput(ID)"><td>10/07/2026</td>
 *                                                   から「日付 → workout_id」を拾う
 *   3. GET  /calendar/strength/modal?workout_id= … ポップアップの HTML（JSON の data に入っている）
 *
 * ■ 初回の準備
 *   1. メニュー「TRAQ取り込み」→「ログイン情報を設定」で TRAQ のメールとパスワードを保存
 *      （ユーザープロパティに保存されるので、他の編集者からは見えません）
 *   2. メニュー「TRAQ取り込み」→「選手IDの対応表を作成/更新」で「TRAQ設定」シートを作り、
 *      B列に各選手の TRAQ の ID（または athletes/view/124653 の URL）を貼る
 *
 * ■ 転記ルール
 *   種目   … Exercise をそのまま
 *   セット … 重量か回数のどちらかが入力されたセットの数
 *   レップ … Reps 列をそのまま（3e / 10e など）
 *   重量   … 入力されたセットの重量。数字なら kg を付け、空なら BW。
 *            セットごとに違えば「50kg, 55kg, 60kg」のように並べる
 *   1セットも入力されていない種目は「未実施」とみなして書きません。
 *
 * 名前はすべて trq で始めています（Code.gs の rpt と同じ理由で、名前の衝突を避けるため）。
 */


// =====================================================================
// 設定
// =====================================================================
var TRQ_CFG = {
  BASE_URL: 'https://traq.drivelinebaseball.com',

  // 選手シート名 ↔ TRAQ の選手ID を並べるシート
  MAP_SHEET: 'TRAQ設定',

  // 毎日の自動取り込みを走らせる時刻（0〜23時。分は Google 側で前後します）
  DAILY_HOUR: 22
};


// =====================================================================
// メニュー（Code.gs の onOpen() から trqBuildMenu_() を呼んでいます）
// =====================================================================
function trqBuildMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('TRAQ取り込み')
    .addItem('TRAQから取り込む…', 'trqImportPrompt')
    .addSeparator()
    .addItem('選手IDの対応表を作成/更新', 'trqSetupMapSheet')
    .addItem('ログイン情報を設定', 'trqSetCredentials')
    .addItem('毎日の自動取り込みを ON/OFF', 'trqToggleDailyTrigger')
    .addToUi();
}


/** メニュー「TRAQから取り込む…」 */
function trqImportPrompt() {
  var ui = SpreadsheetApp.getUi();
  var today = new Date();
  var todayLabel = (today.getMonth() + 1) + '月' + today.getDate() + '日';

  var res = ui.prompt('TRAQから取り込む',
    '取り込む日付を入力してください（例: 10/7、10月7日）。\n空欄のままOKで今日（' + todayLabel + '）を取り込みます。',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;

  var date = trq_parseInputDate_(res.getResponseText(), today);
  if (!date) {
    ui.alert('日付が読み取れませんでした: ' + res.getResponseText());
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('TRAQ からデータを取得しています…', 'TRAQ取り込み', -1);
  var fetched = trq_fetchAll_(ss, date);

  // 既に何か書かれている選手は、上書きしてよいか確認する
  var dateKey = trq_dateKey_(date);
  var filled = fetched.results.filter(function (r) {
    return r.rows && r.rows.length && trq_hasExistingData_(ss, r.sheet, dateKey);
  }).map(function (r) { return r.sheet; });

  var overwrite = true;
  if (filled.length) {
    var ans = ui.alert('上書きの確認',
      dateKey + ' に既にデータが入っている選手がいます:\n  ' + filled.join('、') +
      '\n\nTRAQ の内容で上書きしますか？\n（「いいえ」ならその選手は飛ばします）',
      ui.ButtonSet.YES_NO_CANCEL);
    if (ans === ui.Button.CANCEL || ans === ui.Button.CLOSE) { ss.toast('中止しました', 'TRAQ取り込み'); return; }
    overwrite = (ans === ui.Button.YES);
  }

  var report = trq_writeAll_(ss, dateKey, fetched, overwrite);
  ss.toast('完了しました', 'TRAQ取り込み', 3);
  ui.alert('TRAQ取り込み（' + dateKey + '）', report, ui.ButtonSet.OK);
}


/** 時間主導トリガーから呼ばれる。今日の分を取り込み、既にデータがある選手は上書きしない */
function trqImportScheduled() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var date = new Date();
  var fetched = trq_fetchAll_(ss, date);
  var report = trq_writeAll_(ss, trq_dateKey_(date), fetched, false);
  console.log(report);
}


/** メニュー「ログイン情報を設定」 */
function trqSetCredentials() {
  var ui = SpreadsheetApp.getUi();
  var id = ui.prompt('TRAQ ログイン情報 (1/2)', 'TRAQ のメールアドレス（TRAQ_ID）', ui.ButtonSet.OK_CANCEL);
  if (id.getSelectedButton() !== ui.Button.OK || !id.getResponseText().trim()) return;
  var pw = ui.prompt('TRAQ ログイン情報 (2/2)', 'TRAQ のパスワード（TRAQ_PASS）', ui.ButtonSet.OK_CANCEL);
  if (pw.getSelectedButton() !== ui.Button.OK || !pw.getResponseText()) return;

  PropertiesService.getUserProperties().setProperties({
    TRAQ_ID: id.getResponseText().trim(),
    TRAQ_PASS: pw.getResponseText()
  });

  // 保存したらその場でログインを試して、間違いに早く気づけるようにする
  try {
    trq_login_();
    ui.alert('保存しました。TRAQ へのログインも確認できました。');
  } catch (e) {
    ui.alert('保存しましたが、ログインに失敗しました:\n' + e.message);
  }
}


/** メニュー「選手IDの対応表を作成/更新」 */
function trqSetupMapSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TRQ_CFG.MAP_SHEET);
  if (!sh) {
    sh = ss.insertSheet(TRQ_CFG.MAP_SHEET);
    sh.getRange(1, 1, 1, 3).setValues([['選手シート名', 'TRAQ 選手ID（または URL）', 'メモ']])
      .setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 180);
    sh.setColumnWidth(2, 320);
  }

  // 日付行がある＝選手シート（Code.gs の索引をそのまま使う）
  var idx = rpt_loadIndex_(true);
  var have = trq_readMap_(ss, true).map(function (m) { return m.sheet; });
  var add = idx.sheets
    .map(function (s) { return s.name; })
    .filter(function (n) { return have.indexOf(n) < 0; })
    .map(function (n) { return [n, '', '']; });

  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 3).setValues(add);
  ss.setActiveSheet(sh);
  SpreadsheetApp.getUi().alert(
    (add.length ? add.length + ' 人を追加しました。' : '追加する選手はいませんでした。') +
    '\nB列に TRAQ の選手ID（例: 124653）か、選手ページの URL を貼ってください。\n空欄の選手は取り込みの対象外になります。');
}


/** メニュー「毎日の自動取り込みを ON/OFF」 */
function trqToggleDailyTrigger() {
  var ui = SpreadsheetApp.getUi();
  var mine = ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === 'trqImportScheduled';
  });

  if (mine.length) {
    mine.forEach(function (t) { ScriptApp.deleteTrigger(t); });
    ui.alert('毎日の自動取り込みを OFF にしました。');
    return;
  }

  ScriptApp.newTrigger('trqImportScheduled')
    .timeBased().everyDays(1).atHour(TRQ_CFG.DAILY_HOUR).create();
  ui.alert('毎日の自動取り込みを ON にしました。\n' +
    '毎日 ' + TRQ_CFG.DAILY_HOUR + '時台に、その日の分を取り込みます' +
    '（既にデータがある選手は上書きしません）。');
}


// =====================================================================
// 取得
// =====================================================================

/**
 * 対応表の全選手について、指定日のトレーニングを TRAQ から取得する。
 * 戻り値: { results: [{ sheet, rows: [[種目, セット, レップ, 重量], ...] | null, note }] }
 */
function trq_fetchAll_(ss, date) {
  var map = trq_readMap_(ss, false);
  if (!map.length) {
    throw new Error('「' + TRQ_CFG.MAP_SHEET + '」シートに TRAQ の選手IDが1件もありません。' +
                    'メニューの「選手IDの対応表を作成/更新」から設定してください。');
  }

  var cookies = trq_login_();
  var traqDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yyyy');

  // 選手ページをまとめて取得（1人 1MB 強あるので並列で）
  var pages = UrlFetchApp.fetchAll(map.map(function (m) {
    return trq_req_(TRQ_CFG.BASE_URL + '/athletes/view/' + m.athleteId, cookies);
  }));

  var results = [];
  var modalReqs = [];
  for (var i = 0; i < map.length; i++) {
    var r = { sheet: map[i].sheet, rows: null, note: '' };
    results.push(r);

    if (pages[i].getResponseCode() !== 200) {
      r.note = '選手ページを開けませんでした（HTTP ' + pages[i].getResponseCode() + '、ID ' + map[i].athleteId + '）';
      continue;
    }
    var ids = trq_findWorkoutIds_(pages[i].getContentText(), traqDate);
    if (!ids.length) { r.note = 'この日のトレーニングがありません'; continue; }

    r.rows = [];
    ids.forEach(function (wid) {
      modalReqs.push({ result: r, req: trq_req_(
        TRQ_CFG.BASE_URL + '/calendar/strength/modal?workout_id=' + wid, cookies,
        { 'X-Requested-With': 'XMLHttpRequest', 'Accept': 'application/json' }) });
    });
  }

  // ポップアップの中身をまとめて取得。1日に複数のワークアウトがあれば続けて並べる
  var modals = modalReqs.length ? UrlFetchApp.fetchAll(modalReqs.map(function (m) { return m.req; })) : [];
  for (var k = 0; k < modals.length; k++) {
    var target = modalReqs[k].result;
    var json;
    try { json = JSON.parse(modals[k].getContentText()); } catch (e) { json = null; }
    if (!json || !json.success || !json.data) {
      target.note = 'ポップアップの取得に失敗しました（HTTP ' + modals[k].getResponseCode() + '）';
      continue;
    }
    target.rows = target.rows.concat(trq_toSheetRows_(trq_parseModal_(json.data)));
  }

  results.forEach(function (r) {
    if (r.rows && !r.rows.length && !r.note) r.note = 'TRAQ にまだ入力されていません';
  });
  return { results: results };
}


/** ログインして Cookie ヘッダー文字列を返す */
function trq_login_() {
  var props = PropertiesService.getUserProperties();
  var email = props.getProperty('TRAQ_ID') || PropertiesService.getScriptProperties().getProperty('TRAQ_ID');
  var pass = props.getProperty('TRAQ_PASS') || PropertiesService.getScriptProperties().getProperty('TRAQ_PASS');
  if (!email || !pass) {
    throw new Error('TRAQ のログイン情報が未設定です。メニューの「ログイン情報を設定」から入力してください。');
  }

  var jar = {};
  var page = UrlFetchApp.fetch(TRQ_CFG.BASE_URL + '/login', { followRedirects: false, muteHttpExceptions: true });
  trq_storeCookies_(jar, page);
  var m = page.getContentText().match(/name="_token"\s+value="([^"]+)"/);
  if (!m) throw new Error('TRAQ のログイン画面の形式が変わったようです（_token が見つかりません）。');

  var res = UrlFetchApp.fetch(TRQ_CFG.BASE_URL + '/login', {
    method: 'post',
    payload: { _token: m[1], email: email, password: pass, remember: '1' },
    headers: { Cookie: trq_cookieHeader_(jar) },
    followRedirects: false,
    muteHttpExceptions: true
  });
  trq_storeCookies_(jar, res);

  // 成功時は /profile などへ、失敗時は /login へリダイレクトされる
  var loc = String(trq_header_(res, 'Location') || '');
  if (res.getResponseCode() !== 302 || /\/login\b/.test(loc)) {
    throw new Error('TRAQ にログインできませんでした。メールアドレスとパスワードを確認してください。');
  }
  return trq_cookieHeader_(jar);
}


function trq_req_(url, cookies, extraHeaders) {
  var headers = { Cookie: cookies };
  for (var k in (extraHeaders || {})) headers[k] = extraHeaders[k];
  return { url: url, headers: headers, followRedirects: false, muteHttpExceptions: true };
}


function trq_storeCookies_(jar, res) {
  var sc = trq_header_(res, 'Set-Cookie');
  if (!sc) return;
  (Array.isArray(sc) ? sc : [sc]).forEach(function (c) {
    var kv = String(c).split(';')[0];
    var eq = kv.indexOf('=');
    if (eq > 0) jar[kv.slice(0, eq).trim()] = kv.slice(eq + 1).trim();
  });
}


/** レスポンスヘッダーを大文字小文字を無視して取る（HTTP/2 だと小文字で返ることがある） */
function trq_header_(res, name) {
  var all = res.getAllHeaders();
  for (var k in all) if (k.toLowerCase() === name.toLowerCase()) return all[k];
  return null;
}


function trq_cookieHeader_(jar) {
  return Object.keys(jar).map(function (k) { return k + '=' + jar[k]; }).join('; ');
}


// =====================================================================
// HTML の解析（TRAQ 側の HTML が変わったら直すのはここ）
// =====================================================================

/** 選手ページから、指定日（MM/dd/yyyy）の workout_id を全部拾う */
function trq_findWorkoutIds_(html, traqDate) {
  var re = /strengthInput\((\d+)\)"[^>]*>\s*<td[^>]*>\s*(\d{2}\/\d{2}\/\d{4})\s*<\/td>/g;
  var ids = [];
  var m;
  while ((m = re.exec(html)) !== null) {
    if (m[2] === traqDate && ids.indexOf(m[1]) < 0) ids.push(m[1]);
  }
  return ids;
}


/**
 * ポップアップの HTML を種目ごとに分解する。
 * 戻り値: [{ exercise, reps, sets: [{ weight, reps }, ...] }, ...]
 */
function trq_parseModal_(html) {
  var out = [];
  var chunks = String(html).split(/<tr[\s>]/i).slice(1);

  for (var i = 0; i < chunks.length; i++) {
    var tr = chunks[i].split(/<\/tr>/i)[0];
    if (tr.indexOf('values[') < 0) continue;              // 見出し行などは飛ばす

    var tds = tr.match(/<td[^>]*>[\s\S]*?<\/td>/gi) || [];
    if (tds.length < 2) continue;

    var sets = {};
    var inputs = tr.match(/<input\b[^>]*>/gi) || [];
    inputs.forEach(function (tag) {
      var name = trq_attr_(tag, 'name');
      var nm = name && name.match(/\[(set|rep)_(\d+)\]$/);
      if (!nm) return;                                    // Notes など
      var n = Number(nm[2]);
      sets[n] = sets[n] || { weight: '', reps: '' };
      sets[n][nm[1] === 'set' ? 'weight' : 'reps'] = trq_decode_(trq_attr_(tag, 'value') || '').trim();
    });

    out.push({
      exercise: trq_text_(tds[0]),
      reps: trq_text_(tds[1]),
      sets: Object.keys(sets).sort(function (a, b) { return a - b; }).map(function (k) { return sets[k]; })
    });
  }
  return out;
}


/** 種目の配列 → シートに書く [種目, セット, レップ, 重量] の配列 */
function trq_toSheetRows_(exercises) {
  var rows = [];
  exercises.forEach(function (ex) {
    var done = ex.sets.filter(function (s) { return s.weight !== '' || s.reps !== ''; });
    if (!done.length) return;                             // 未実施の種目は書かない

    var weights = [];
    done.forEach(function (s) {
      var w = s.weight === '' ? 'BW' : (/^\d+(\.\d+)?$/.test(s.weight) ? s.weight + 'kg' : s.weight);
      if (weights.indexOf(w) < 0) weights.push(w);
    });

    rows.push([ex.exercise, done.length, trq_cellValue_(ex.reps), weights.join(', ')]);
  });
  return rows;
}


/** 数字だけなら数値で、それ以外は文字列のまま書く（"10-12" などが日付に化けないよう ' を付ける） */
function trq_cellValue_(s) {
  s = String(s).trim();
  if (/^\d+(\.\d+)?$/.test(s)) return Number(s);
  return s === '' ? '' : "'" + s;
}


function trq_attr_(tag, name) {
  var m = tag.match(new RegExp('\\s' + name + '\\s*=\\s*"([^"]*)"', 'i'));
  return m ? m[1] : null;
}


function trq_text_(html) {
  return trq_decode_(String(html).replace(/<[^>]*>/g, ' ')).replace(/\s+/g, ' ').trim();
}


function trq_decode_(s) {
  return String(s)
    .replace(/&nbsp;/g, ' ')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&#0*39;/g, "'").replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, function (_, d) { return String.fromCharCode(Number(d)); })
    .replace(/&#x([0-9a-f]+);/gi, function (_, h) { return String.fromCharCode(parseInt(h, 16)); })
    .replace(/&amp;/g, '&');
}


// =====================================================================
// 書き込み
// =====================================================================

/** 取得結果をシートへ書き、結果報告の文字列を返す */
function trq_writeAll_(ss, dateKey, fetched, overwrite) {
  var ok = [], skipped = [], failed = [];

  fetched.results.forEach(function (r) {
    if (!r.rows || !r.rows.length) { skipped.push(r.sheet + '：' + r.note); return; }
    var sh = ss.getSheetByName(r.sheet);
    if (!sh) { failed.push(r.sheet + '：シートが見つかりません'); return; }

    try {
      var col = trq_findDateCol_(sh, dateKey) || trq_addDateBlock_(sh, dateKey);
      var start = RPT_CFG.DATA_START_ROW;
      var area = sh.getRange(start, col, RPT_CFG.MAX_DATA_ROWS, 4);

      if (!overwrite && trq_blockHasData_(area)) { skipped.push(r.sheet + '：既にデータあり（上書きせず）'); return; }

      area.clearContent();
      sh.getRange(start, col, r.rows.length, 4).setValues(r.rows);
      ok.push(r.sheet + '（' + r.rows.length + '種目）');
    } catch (e) {
      failed.push(r.sheet + '：' + e.message);
    }
  });

  // PDF 作成側（Code.gs）の日付索引キャッシュを捨てて、追加した日付列がすぐ選べるようにする
  try { CacheService.getDocumentCache().remove('rptIndex_v1'); } catch (e) { /* 無くても動く */ }

  var lines = ['転記 ' + ok.length + ' 人'];
  if (ok.length) lines.push('  ' + ok.join('\n  '));
  if (skipped.length) lines.push('', '対象外 ' + skipped.length + ' 人', '  ' + skipped.join('\n  '));
  if (failed.length) lines.push('', 'エラー ' + failed.length + ' 人', '  ' + failed.join('\n  '));
  return lines.join('\n');
}


/** 1〜2行目から、その日の開始列（結合セルの左端）を探す。無ければ null */
function trq_findDateCol_(sh, dateKey) {
  var head = sh.getRange(1, 1, 2, sh.getLastColumn()).getValues();
  for (var r = 0; r < head.length; r++) {
    for (var c = 0; c < head[r].length; c++) {
      if (rpt_dateKey_(head[r][c]) === dateKey) return c + 1;
    }
  }
  return null;
}


/**
 * 一番右の日付ブロックの隣に、新しい日付の4列を作る。
 * 書式・列幅・2行目の見出し（セット/レップ/重量）は直前のブロックからコピーします。
 */
function trq_addDateBlock_(sh, dateKey) {
  var head = sh.getRange(1, 1, 2, sh.getLastColumn()).getValues();
  var dateRow = 0, lastCol = 0;
  for (var r = 0; r < head.length; r++) {
    for (var c = 0; c < head[r].length; c++) {
      if (rpt_dateKey_(head[r][c])) { dateRow = r + 1; lastCol = Math.max(lastCol, c + 1); }
    }
    if (dateRow) break;
  }
  if (!dateRow) throw new Error('日付の行が見つからないため、日付列を追加できません');

  var newCol = lastCol + 4;
  var need = newCol + 3 - sh.getMaxColumns();
  if (need > 0) sh.insertColumnsAfter(sh.getMaxColumns(), need);

  var rows = sh.getMaxRows();
  sh.getRange(1, lastCol, rows, 4).copyTo(sh.getRange(1, newCol, rows, 4), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  for (var i = 0; i < 4; i++) sh.setColumnWidth(newCol + i, sh.getColumnWidth(lastCol + i));

  // 見出し行（セット/レップ/重量）をコピー
  var other = dateRow === 1 ? 2 : 1;
  sh.getRange(other, lastCol, 1, 4).copyTo(sh.getRange(other, newCol, 1, 4));

  // 日付セル。直前のセルが日付型なら日付型で、文字列なら「10月7日」の文字列で書く
  var prev = sh.getRange(dateRow, lastCol);
  var cell = sh.getRange(dateRow, newCol, 1, 4);
  if (prev.isPartOfMerge()) cell.merge();
  var md = rpt_parseDateKey_(dateKey);
  if (prev.getValue() instanceof Date) {
    var y = prev.getValue().getFullYear();
    if (md.month < prev.getValue().getMonth() + 1) y++;   // 12月 → 1月 の年またぎ
    cell.getCell(1, 1).setValue(new Date(y, md.month - 1, md.day));
  } else {
    cell.getCell(1, 1).setValue("'" + dateKey);
  }
  return newCol;
}


function trq_blockHasData_(range) {
  return range.getValues().some(function (row) {
    return row.some(function (v) { return v !== '' && v !== null; });
  });
}


function trq_hasExistingData_(ss, sheetName, dateKey) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return false;
  var col = trq_findDateCol_(sh, dateKey);
  return !!col && trq_blockHasData_(sh.getRange(RPT_CFG.DATA_START_ROW, col, RPT_CFG.MAX_DATA_ROWS, 4));
}


// =====================================================================
// 対応表・日付
// =====================================================================

/** 「TRAQ設定」シートを読む。includeBlank=false なら ID 未入力の行は除く */
function trq_readMap_(ss, includeBlank) {
  var sh = ss.getSheetByName(TRQ_CFG.MAP_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues();
  var out = [];
  vals.forEach(function (v) {
    var name = String(v[0]).trim();
    if (!name) return;
    var m = String(v[1]).match(/(?:athletes\/view\/)?(\d+)\s*$/);
    if (!m && !includeBlank) return;
    out.push({ sheet: name, athleteId: m ? m[1] : '' });
  });
  return out;
}


/** 「10/7」「10月7日」「2026/10/7」などを Date に。空欄なら today。年が無ければ未来にならない年を選ぶ */
function trq_parseInputDate_(text, today) {
  var s = String(text || '').trim();
  if (!s) return new Date(today.getFullYear(), today.getMonth(), today.getDate());

  var m = s.match(/^(\d{4})\s*[\/\-年]\s*(\d{1,2})\s*[\/\-月]\s*(\d{1,2})\s*日?$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  m = s.match(/^(\d{1,2})\s*[\/\-月]\s*(\d{1,2})\s*日?$/);
  if (!m) return null;
  var d = new Date(today.getFullYear(), Number(m[1]) - 1, Number(m[2]));
  if (d > today) d.setFullYear(d.getFullYear() - 1);
  return d;
}


function trq_dateKey_(date) {
  return (date.getMonth() + 1) + '月' + date.getDate() + '日';
}
