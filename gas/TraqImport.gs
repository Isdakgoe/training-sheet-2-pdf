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
 *   2. メニュー「TRAQ取り込み」→「TRAQの選手一覧を取得」で「TRAQ設定」シートを作る。
 *      A列 TRAQ の選手名 / B列 TRAQ の選手ID は TRAQ の Athletes 一覧から自動で入ります。
 *   3. C列（転載先タブ）をプルダウンから選ぶ。C列が空の選手は取り込みの対象外です。
 *
 * ■ 取り込み
 *   メニュー「TRAQから取り込む…」→ ダイアログで日付と選手（複数選択）を選んで実行。
 *   選んだ選手ごとに B列の ID のページを開き、取得したデータを C列のタブへ書きます。
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
    .addItem('TRAQから取り込む…', 'trqShowDialog')
    .addSeparator()
    .addItem('TRAQの選手一覧を取得', 'trqSyncAthletes')
    .addItem('ログイン情報を設定', 'trqSetCredentials')
    .addItem('毎日の自動取り込みを ON/OFF', 'trqToggleDailyTrigger')
    .addToUi();
}


/** メニュー「TRAQから取り込む…」 */
function trqShowDialog() {
  var html = HtmlService.createHtmlOutputFromFile('TraqDialog')
    .setWidth(480)
    .setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'TRAQ から取り込む');
}


/** 時間主導トリガーから呼ばれる。C列のある全選手の今日の分を取り込み、既にデータがある選手は上書きしない */
function trqImportScheduled() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var date = new Date();
  var map = trq_readMap_(ss).filter(function (m) { return m.sheet; });
  var fetched = trq_fetchAll_(ss, date, map);
  var report = trq_writeAll_(ss, trq_dateKey_(date), fetched, false);
  console.log(report);
}


// =====================================================================
// ダイアログ（TraqDialog.html）から google.script.run で呼ばれる関数
// =====================================================================

/** ダイアログを開いた直後に呼ばれる。今日の日付と、選べる選手（C列が入っている行）を返す */
function trqGetDialogData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var map = trq_readMap_(ss);
  if (!map.length) {
    throw new Error('「' + TRQ_CFG.MAP_SHEET + '」シートに選手がいません。' +
                    'メニューの「TRAQの選手一覧を取得」を先に実行してください。');
  }
  // 前回選んだ選手を覚えておき、次回の既定にする
  var last = PropertiesService.getUserProperties().getProperty('TRAQ_LAST_SELECTED');
  var lastIds = last ? last.split(',') : null;

  var mapped = map.filter(function (m) { return m.sheet; });
  return {
    today: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    athletes: mapped.map(function (m) {
      return {
        id: m.athleteId,
        traqName: m.traqName,
        sheet: m.sheet,
        checked: lastIds ? lastIds.indexOf(m.athleteId) >= 0 : true
      };
    }),
    unmapped: map.length - mapped.length
  };
}


/** 選ばれた選手のうち、その日に既にデータが入っている転載先タブを返す（上書き確認用） */
function trqCheckExisting(dateStr, ids) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dateKey = trq_dateKey_(trq_requireDate_(dateStr));
  var seen = {};
  return trq_selectMap_(ss, ids).filter(function (m) {
    if (seen[m.sheet]) return false;
    seen[m.sheet] = true;
    return trq_hasExistingData_(ss, m.sheet, dateKey);
  }).map(function (m) { return m.sheet; });
}


/** 「取り込む」ボタンの本体。結果報告の文字列を返す */
function trqRunImport(dateStr, ids, overwrite) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var date = trq_requireDate_(dateStr);
  var map = trq_selectMap_(ss, ids);
  if (!map.length) throw new Error('取り込む選手が選ばれていません。');

  PropertiesService.getUserProperties().setProperty('TRAQ_LAST_SELECTED', ids.join(','));

  var fetched = trq_fetchAll_(ss, date, map);
  return trq_writeAll_(ss, trq_dateKey_(date), fetched, overwrite);
}


function trq_requireDate_(dateStr) {
  var date = trq_parseInputDate_(dateStr, new Date());
  if (!date) throw new Error('日付が読み取れませんでした: ' + dateStr);
  return date;
}


// =====================================================================
// メニューの各機能
// =====================================================================


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


/**
 * メニュー「TRAQの選手一覧を取得」
 * TRAQ の Athletes 一覧（全ページ）から選手名と ID を拾い、「TRAQ設定」シートの A・B 列に入れる。
 * 既にいる選手は C列（転載先タブ）を残したまま名前だけ更新し、新しい選手は末尾に足す。
 */
function trqSyncAthletes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('TRAQ の選手一覧を取得しています…', 'TRAQ取り込み', -1);

  var athletes = trq_fetchAthleteList_(trq_login_());
  if (!athletes.length) throw new Error('TRAQ の選手一覧から選手を1人も読み取れませんでした。');

  var sh = ss.getSheetByName(TRQ_CFG.MAP_SHEET);
  if (!sh) {
    sh = ss.insertSheet(TRQ_CFG.MAP_SHEET);
    sh.getRange(1, 1, 1, 3).setValues([['TRAQ 選手名', 'TRAQ 選手ID', '転載先タブ']])
      .setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 200);
    sh.setColumnWidth(2, 110);
    sh.setColumnWidth(3, 180);
  }

  var rowOf = {};
  trq_readMap_(ss).forEach(function (m) { rowOf[m.athleteId] = m.row; });

  var append = [];
  athletes.forEach(function (a) {
    if (rowOf[a.id]) sh.getRange(rowOf[a.id], 1).setValue(a.name);
    else append.push([a.name, a.id]);
  });
  if (append.length) sh.getRange(sh.getLastRow() + 1, 1, append.length, 2).setValues(append);

  // C列は選手タブ名のプルダウンにする（日付行がある＝選手タブ。Code.gs の索引を流用）
  var tabs = rpt_loadIndex_(true).sheets.map(function (s) { return s.name; });
  if (tabs.length && sh.getLastRow() >= 2) {
    sh.getRange(2, 3, sh.getLastRow() - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(tabs, true).setAllowInvalid(false).build());
  }

  ss.setActiveSheet(sh);
  ss.toast('完了しました', 'TRAQ取り込み', 3);
  SpreadsheetApp.getUi().alert(
    'TRAQ の選手 ' + athletes.length + ' 人を確認しました（新規 ' + append.length + ' 人）。\n' +
    'C列（転載先タブ）をプルダウンから選んでください。C列が空の選手は取り込みの対象外です。');
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
    '毎日 ' + TRQ_CFG.DAILY_HOUR + '時台に、C列が入っている全選手のその日の分を取り込みます' +
    '（既にデータがある選手は上書きしません）。');
}


// =====================================================================
// 取得
// =====================================================================

/**
 * 指定した選手について、指定日のトレーニングを TRAQ から取得する。
 * map: [{ athleteId, traqName, sheet }]（sheet ＝ 転載先タブ）
 * 戻り値: { results: [{ sheet, rows: [[種目, セット, レップ, 重量], ...] | null, note }] }
 */
function trq_fetchAll_(ss, date, map) {
  if (!map.length) {
    throw new Error('取り込める選手がいません。「' + TRQ_CFG.MAP_SHEET + '」シートの C列（転載先タブ）を設定してください。');
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

/**
 * 「TRAQ設定」シートを読む。B列に ID がある行だけ返す（C列は空のこともある）。
 * 戻り値: [{ traqName, athleteId, sheet, row }]
 */
function trq_readMap_(ss) {
  var sh = ss.getSheetByName(TRQ_CFG.MAP_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getDisplayValues();
  var out = [];
  vals.forEach(function (v, i) {
    var m = String(v[1]).match(/(?:athletes\/view\/)?(\d+)\s*$/);
    if (!m) return;
    out.push({ traqName: String(v[0]).trim(), athleteId: m[1], sheet: String(v[2]).trim(), row: i + 2 });
  });
  return out;
}


/** ダイアログで選ばれた ID のうち C列が入っている行を、シートの並び順で返す */
function trq_selectMap_(ss, ids) {
  ids = (ids || []).map(String);
  return trq_readMap_(ss).filter(function (m) {
    return m.sheet && ids.indexOf(m.athleteId) >= 0;
  });
}


/** TRAQ の Athletes 一覧を全ページ読み、[{ id, name }] を返す */
function trq_fetchAthleteList_(cookies) {
  var out = [];
  var seen = {};
  for (var page = 1; page <= 30; page++) {
    var res = UrlFetchApp.fetch(
      TRQ_CFG.BASE_URL + '/athletes?view=table&location=athletes.index&page=' + page,
      { headers: { Cookie: cookies }, followRedirects: false, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) break;

    var html = res.getContentText();
    var re = /<tr class="userRow-(\d+)">([\s\S]*?)<\/tr>/g;
    var m, added = 0;
    while ((m = re.exec(html)) !== null) {
      if (seen[m[1]]) continue;
      seen[m[1]] = true;
      // 名と姓が別々のリンクになっている
      var names = [];
      var nr = /data-athlete-profile-link[^>]*>([\s\S]*?)<\/a>/g, n;
      while ((n = nr.exec(m[2])) !== null) names.push(trq_text_(n[1]));
      out.push({ id: m[1], name: names.join(' ') });
      added++;
    }
    if (!added) break;                                    // 最終ページを過ぎた
  }
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
