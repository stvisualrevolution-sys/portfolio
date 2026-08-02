/**
 * 仕事のご相談フォーム — 受け口
 *
 * ポートフォリオの contact.html から送信された内容を受け取って
 *   1. スプレッドシートに1行記録し（初回に自動で作られる）
 *   2. 自分あてに通知メールを送り
 *   3. 送信してくれた人にも自動返信を出す
 *
 * 設置手順は apps-script/README.md を参照。
 * このファイルは中身を全部コピーして Apps Script に貼るだけでよい。
 */

/* ── 設定 ────────────────────────────────────
   変えたくなったらここだけ触れば済むようにしてある
   ─────────────────────────────────────────── */

var TO_EMAIL   = 'stvisualrevolution@gmail.com';  // 通知メールの宛先
var FROM_NAME  = '医療 × AI 実装事例 / お問い合わせ';
var SS_NAME    = 'ポートフォリオ 問い合わせ';      // 自動で作られるシートの名前
var SHEET_NAME = '問い合わせ';
var TIMEZONE   = 'Asia/Tokyo';

var AUTO_REPLY   = true;  // 送信者への自動返信。要らなければ false
var MAX_PER_HOUR = 20;    // 1時間あたりの受付上限（いたずら・連投よけ）

var LIMITS = { name: 100, org: 120, email: 200, type: 60, budget: 40, when: 40, message: 5000 };
var HEADERS = ['受信日時', '名前', '所属', 'メール', '種類', '予算', '時期', '内容', '送信元'];

/* ── 受信 ────────────────────────────────────
   フォームから送信されると、この doPost が呼ばれる
   ─────────────────────────────────────────── */

function doPost(e) {
  try {
    var body = readBody_(e);
    if (!body) return json_({ ok: false, error: '内容を読み取れませんでした。' });

    // ハニーポット。人間には見えない欄が埋まっていたら自動投稿とみなし、
    // 何もせずに成功だけ返す（弾いたことを相手に悟らせない）
    if (String(body.website || '').trim() !== '') return json_({ ok: true });

    var f = {
      name:    clip_(body.name,    LIMITS.name),
      org:     clip_(body.org,     LIMITS.org),
      email:   clip_(body.email,   LIMITS.email),
      type:    clip_(body.type,    LIMITS.type),
      budget:  clip_(body.budget,  LIMITS.budget),
      when:    clip_(body.when,    LIMITS.when),
      message: clip_(body.message, LIMITS.message),
      page:    clip_(body.page,    300)
    };

    // ブラウザ側でも検査しているが、直接叩かれる場合があるのでここでも見る
    if (!f.name)    return json_({ ok: false, error: 'お名前が入力されていません。' });
    if (!f.message) return json_({ ok: false, error: 'ご相談内容が入力されていません。' });
    if (!isEmail_(f.email)) return json_({ ok: false, error: 'メールアドレスの形式をご確認ください。' });

    if (!underRateLimit_()) {
      return json_({ ok: false, error: '受付が混み合っています。時間をおいてお試しください。' });
    }

    var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

    // 記録に失敗しても通知は止めない。届くことのほうが大事
    var sheetUrl = '';
    try { sheetUrl = logToSheet_(stamp, f); } catch (err) { console.error('sheet: ' + err); }

    notify_(stamp, f, sheetUrl);

    if (AUTO_REPLY) {
      try { autoReply_(f); } catch (err) { console.error('autoreply: ' + err); }
    }

    return json_({ ok: true });

  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: '送信できませんでした。時間をおいてお試しください。' });
  }
}

/** ブラウザで URL を直接開いたときの応答。生きているかの確認用 */
function doGet() {
  return json_({ ok: true, message: 'contact endpoint is alive' });
}

/* ── 記録 ────────────────────────────────────
   保存先のスプレッドシートは初回に自動で作られ、
   その ID を覚えておいて2回目以降は同じものを使う
   ─────────────────────────────────────────── */

function logToSheet_(stamp, f) {
  var ss = getSpreadsheet_();

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (err) { return ss.getUrl(); }

  try {
    var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sh.setFrozenRows(1);
      sh.setColumnWidth(1, 140);
      sh.setColumnWidth(8, 480);
    }
    sh.appendRow([stamp, f.name, f.org, f.email, f.type, f.budget, f.when, f.message, f.page]);
  } finally {
    lock.releaseLock();
  }

  return ss.getUrl();
}

function getSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SS_ID');

  if (id) {
    // 手で消された場合に備えて、開けなかったら作り直す
    try { return SpreadsheetApp.openById(id); } catch (err) { /* 下へ */ }
  }

  var ss = SpreadsheetApp.create(SS_NAME);
  props.setProperty('SS_ID', ss.getId());

  var first = ss.getSheets()[0];
  first.setName(SHEET_NAME);

  return ss;
}

/* ── 通知 ───────────────────────────────────── */

function notify_(stamp, f, sheetUrl) {
  var subject = '【依頼】' + (f.type || 'お問い合わせ') + ' — ' + f.name + ' 様';

  var lines = [
    'ポートフォリオのフォームから相談が届きました。',
    'このメールにそのまま返信すれば、送信者に届きます。',
    '',
    '──────────────────────────────',
    '受信日時 : ' + stamp,
    'お名前   : ' + f.name,
    '所属     : ' + (f.org    || '（未記入）'),
    'メール   : ' + f.email,
    '種類     : ' + (f.type   || '（未選択）'),
    '予算     : ' + (f.budget || '未定'),
    '時期     : ' + (f.when   || '未定'),
    '──────────────────────────────',
    '',
    '【ご相談内容】',
    f.message,
    '',
    '──────────────────────────────',
    '送信元 : ' + (f.page || '-'),
    '記録   : ' + (sheetUrl || '（記録に失敗しました）')
  ];

  MailApp.sendEmail({
    to: TO_EMAIL,
    subject: subject,
    body: lines.join('\n'),
    replyTo: f.email,   // 返信すると相談者に届く
    name: FROM_NAME
  });
}

function autoReply_(f) {
  // 送信枠が尽きかけているときは、自分あての通知を優先して自動返信は止める
  if (MailApp.getRemainingDailyQuota() < 5) return;

  var lines = [
    f.name + ' 様',
    '',
    'お問い合わせありがとうございます。以下の内容を受け付けました。',
    '内容を確認のうえ、改めてこのアドレスから返信いたします。',
    '本業が診療のため、返信までに数日いただくことがあります。',
    '',
    '──────────────────────────────',
    '種類 : ' + (f.type   || '（未選択）'),
    '予算 : ' + (f.budget || '未定'),
    '時期 : ' + (f.when   || '未定'),
    '',
    f.message,
    '──────────────────────────────',
    '',
    '※このメールは自動送信です。ご返信いただいても問題ありません。',
    '',
    '医療 × AI 実装事例',
    'https://stvisualrevolution-sys.github.io/portfolio/'
  ];

  MailApp.sendEmail({
    to: f.email,
    subject: '【自動返信】お問い合わせを受け付けました',
    body: lines.join('\n'),
    replyTo: TO_EMAIL,
    name: FROM_NAME
  });
}

/* ── 道具 ───────────────────────────────────── */

function readBody_(e) {
  if (!e) return null;
  if (e.postData && e.postData.contents) {
    try { return JSON.parse(e.postData.contents); } catch (err) { /* 下へ */ }
  }
  if (e.parameter && Object.keys(e.parameter).length) return e.parameter;
  return null;
}

function clip_(v, max) {
  return String(v == null ? '' : v).trim().slice(0, max);
}

function isEmail_(s) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s);
}

function underRateLimit_() {
  var props = PropertiesService.getScriptProperties();
  var key = 'rate_' + Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHH');
  var n = Number(props.getProperty(key) || 0) + 1;
  props.setProperty(key, String(n));
  return n <= MAX_PER_HOUR;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── 動作確認 ─────────────────────────────────
   エディタ上部の関数プルダウンで testSend を選んで「実行」。
   スプレッドシートが作られ、テストのメールが2通届けば設定完了。
   ─────────────────────────────────────────── */

function testSend() {
  var res = doPost({
    postData: {
      contents: JSON.stringify({
        name: 'テスト太郎',
        org: 'テストクリニック',
        email: TO_EMAIL,
        type: '新規開発のご依頼',
        budget: '30〜100万円',
        when: '3か月以内',
        message: 'これは動作確認の送信です。届いていれば設定は完了しています。',
        page: 'testSend'
      })
    }
  });
  console.log(res.getContent());
  console.log('保存先: ' + getSpreadsheet_().getUrl());
}
