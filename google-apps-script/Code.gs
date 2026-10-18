/**
 * AI 自走車常態班 報名 → Google 試算表 接收程式＋確認信
 * 使用方式見同資料夾「設定說明.md」
 *
 * 網站會分兩次送資料：
 *  1. stage = 'register'：家長送出報名表 → 先寫入一列，處理狀態為「未完成」
 *  2. stage = 'finalize'：整期填完轉帳後 5 碼（或按「稍後再轉帳」）／試聽選完聯繫時段
 *     → 更新同一列，寄確認信給家長、寄通知信給 ADMIN_EMAIL
 * 只按到第 1 步就關掉網頁的報名，會留在試算表裡顯示「未完成」，不會寄信。
 */

const SHEET_NAME = '報名資料';
// 早鳥以伺服器時間判斷（不信任訪客電腦的時間）
const EARLY_BIRD_DEADLINE = new Date('2026-11-30T23:59:59+08:00');
const PRICE_ORIGINAL = 4800;
const PRICE_EARLY = 4200;
// 有人報名時寄通知信給這個信箱（確認信也從部署這支程式的帳號寄出）
const ADMIN_EMAIL = 'randdlaboratary@gmail.com';
const SENDER_NAME = '微研 Weiyan 科學教育';
const COURSE_NAME = 'AI 自走車常態班｜小小駕訓班';
const COURSE_TIME = '每週三 16:00–17:30（一期 8 堂，每堂 90 分鐘）';
const COURSE_PLACE = '強納森美語教室（高科大旁）';
const BANK_INFO = '中國信託 (822) 營業部｜帳號 901890457506';
const LINE_URL = 'https://lin.ee/rZW93KN';

// 欄位順序＝試算表欄位順序；key 對應網站送來的資料欄位
const COLUMNS = [
  ['收到時間', null], ['報名編號', 'refCode'], ['報名類型', 'type'], ['課程', 'course'],
  ['學生姓名', 'studentName'], ['年級', 'grade'], ['就讀學校', 'school'], ['程式基礎', 'codingLevel'],
  ['家長姓名', 'parentName'], ['家長電話', 'parentPhone'], ['家長Email', 'parentEmail'],
  ['緊急聯絡人', 'emergencyName'], ['緊急聯絡電話', 'emergencyPhone'],
  ['應繳金額', 'amount'], ['早鳥', 'earlyBird'], ['轉帳後五碼', 'last5'], ['方便聯繫時間', 'contactTimes'],
  ['從哪裡知道', 'source'], ['家長備註', 'note'], ['照片公開', 'photo'], ['個資同意', 'consent'],
  ['完成時間', null], ['處理狀態', null], ['備註', null]
];
const HEADERS = COLUMNS.map(c => c[0]);
const col_ = name => HEADERS.indexOf(name) + 1;

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  const first = sh.getLastRow() === 0 ? [] : sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (first.join('|') !== HEADERS.join('|')) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight('bold').setBackground('#0B2A66').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

// 防止以 = + - @ 開頭的內容被試算表當成公式執行
function safe_(v) {
  const s = String(v == null ? '' : v).slice(0, 500);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function now_() {
  return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
}

function findRow_(sh, refCode) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const ids = sh.getRange(2, col_('報名編號'), n - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === refCode) return i + 2;
  return 0;
}

function appendRow_(sh, d, status, doneAt) {
  const row = COLUMNS.map(([h, k]) => {
    if (h === '收到時間') return now_();
    if (h === '完成時間') return doneAt;
    if (h === '處理狀態') return status;
    return k ? safe_(d[k]) : '';
  });
  // 不用 appendRow：它會無視純文字格式，把 09xx 電話轉成數字吃掉開頭的 0
  sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setNumberFormat('@').setValues([row]);
}

function setCell_(sh, r, name, value) {
  sh.getRange(r, col_(name)).setNumberFormat('@').setValue(safe_(value));
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d.refCode || !d.studentName || !d.parentPhone) throw new Error('缺少必要欄位');
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(d.parentEmail || ''))) throw new Error('Email 格式不正確');
    d.refCode = String(d.refCode).slice(0, 40);
    const trial = d.type === '試聽一堂';
    if (trial) {
      d.amount = '試聽（專人聯繫）'; d.earlyBird = '';
    } else {
      const early = new Date() <= EARLY_BIRD_DEADLINE;
      d.earlyBird = early ? '是' : '否';
      d.amount = String(early ? PRICE_EARLY : PRICE_ORIGINAL);
    }
    d.consent = '同意';
    d.photo = d.photo === '不同意公開' ? '不同意公開' : '可公開';
    const sh = getSheet_();
    const props = PropertiesService.getScriptProperties();

    if (d.stage === 'finalize') return json_(finalize_(sh, props, d, trial));

    d.last5 = ''; d.contactTimes = '';
    appendRow_(sh, d, trial ? '未完成（未選聯繫時段）' : '未完成（未填轉帳資訊）', '');
    // 第 2 步要帶同一組 token 才能更新這一列，避免別人亂改或重複寄信
    if (d.token) props.setProperty('t_' + d.refCode, String(d.token).slice(0, 64));
    return json_({ ok: true, refCode: d.refCode });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function finalize_(sh, props, d, trial) {
  const key = 't_' + d.refCode;
  const r = findRow_(sh, d.refCode);
  if (r && props.getProperty(key) !== String(d.token)) throw new Error('驗證失敗或已完成');
  d.last5 = trial ? '' : String(d.last5 || '').replace(/\D/g, '').slice(0, 5);
  d.contactTimes = trial ? String(d.contactTimes || '').slice(0, 60) : '';
  const status = trial ? '待聯繫' : (d.last5 ? '待對帳' : '待轉帳');
  const doneAt = now_();
  if (r) {
    setCell_(sh, r, '轉帳後五碼', d.last5);
    setCell_(sh, r, '方便聯繫時間', d.contactTimes);
    setCell_(sh, r, '完成時間', doneAt);
    setCell_(sh, r, '處理狀態', status);
  } else {
    // 第 1 步沒寫進來（例如網路中斷）就整列補上
    appendRow_(sh, d, status, doneAt);
  }
  props.deleteProperty(key);

  let mailNote = '';
  try {
    sendParentMail_(d, trial);
  } catch (err) {
    mailNote = '家長確認信寄送失敗：' + err;
  }
  sendAdminMail_(d, trial, status, mailNote);
  return { ok: true, refCode: d.refCode };
}

/* ================== Email ================== */

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function table_(rows) {
  return '<table style="width:100%;border-collapse:collapse;font-size:15px;margin:14px 0">' +
    rows.filter(r => r[1] !== '' && r[1] != null).map(([k, v]) =>
      `<tr><td style="padding:9px 0;border-bottom:1px solid #eef0f5;color:#6b7280;width:120px;vertical-align:top">${esc_(k)}</td>` +
      `<td style="padding:9px 0;border-bottom:1px solid #eef0f5;font-weight:600;color:#1f2430">${esc_(v)}</td></tr>`).join('') +
    '</table>';
}

function shell_(title, bodyHtml) {
  return '<div style="font-family:\'Noto Sans TC\',\'Microsoft JhengHei\',\'PingFang TC\',sans-serif;max-width:580px;margin:0 auto;color:#1f2430;line-height:1.75">' +
    '<div style="background:#0B2A66;color:#fff;padding:22px 26px;border-radius:14px 14px 0 0">' +
    `<div style="font-size:13px;opacity:.75">${esc_(SENDER_NAME)}</div><div style="font-size:22px;font-weight:700;margin-top:2px">${esc_(title)}</div></div>` +
    `<div style="border:1px solid #e3e7f0;border-top:0;border-radius:0 0 14px 14px;padding:22px 26px">${bodyHtml}</div></div>`;
}

function sendParentMail_(d, trial) {
  const hi = `<p>${esc_(d.parentName)} 您好：</p>`;
  const sign = `<p style="margin-top:18px">有任何問題，歡迎直接回覆這封信，或加入<a href="${LINE_URL}" style="color:#0A63D6">微研官方 LINE</a>。</p>` +
    `<p style="color:#6b7280;font-size:14px">${esc_(SENDER_NAME)} 敬上</p>`;
  let subject, title, body;
  if (trial) {
    subject = `【預約成功】${COURSE_NAME} 試聽｜${d.studentName}`;
    title = '試聽預約成功！期待相見～';
    body = hi + `<p>已收到 <b>${esc_(d.studentName)}</b> 的試聽預約，謝謝您！我們會在您方便的時段與您聯繫，安排試聽日期。</p>` +
      table_([['報名編號', d.refCode], ['學生', `${d.studentName}（${d.grade}）`], ['方便聯繫時間', d.contactTimes],
        ['課程時間', COURSE_TIME], ['上課地點', COURSE_PLACE]]) + sign;
  } else if (d.last5) {
    subject = `【報名成功】${COURSE_NAME}｜${d.studentName}`;
    title = '報名成功！期待相見～';
    body = hi + `<p>已收到 <b>${esc_(d.studentName)}</b> 的報名與轉帳資訊，謝謝您！我們會盡快核對入帳。</p>` +
      table_([['報名編號', d.refCode], ['學生', `${d.studentName}（${d.grade}）`], ['上課時間', COURSE_TIME],
        ['上課地點', COURSE_PLACE], ['應繳金額', `NT$${Number(d.amount).toLocaleString('en-US')}`], ['轉帳後 5 碼', d.last5]]) +
      '<p>開班日期與上課細節，我們會再與您聯繫確認。上課請讓孩子自備 <b>Windows 筆記型電腦</b>。</p>' + sign;
  } else {
    subject = `【報名已收到】${COURSE_NAME}｜${d.studentName}｜請完成轉帳`;
    title = '已收到報名，請完成轉帳';
    body = hi + `<p>已收到 <b>${esc_(d.studentName)}</b> 的報名，期待相見～請在方便時完成轉帳：</p>` +
      table_([['報名編號', d.refCode], ['轉帳帳號', BANK_INFO], ['應繳金額', `NT$${Number(d.amount).toLocaleString('en-US')}`],
        ['上課時間', COURSE_TIME], ['上課地點', COURSE_PLACE]]) +
      '<p style="background:#FFF6DD;border-radius:10px;padding:12px 14px"><b>轉帳完成後，請直接回覆這封信，告訴我們轉帳帳號後 5 碼</b>，方便我們核對。</p>' + sign;
  }
  MailApp.sendEmail({ to: d.parentEmail, subject, htmlBody: shell_(title, body), name: SENDER_NAME, replyTo: ADMIN_EMAIL });
}

function sendAdminMail_(d, trial, status, mailNote) {
  const amount = trial ? '' : `NT$${Number(d.amount).toLocaleString('en-US')}`;
  const subject = trial
    ? `【新報名｜試聽】${d.studentName}（${d.grade}）｜聯繫：${d.contactTimes}`
    : `【新報名｜整期】${d.studentName}（${d.grade}）｜${amount}｜${d.last5 ? '後五碼 ' + d.last5 : '尚未轉帳'}`;
  const key = trial
    ? [['類型', '試聽一堂'], ['方便聯繫時間', d.contactTimes], ['家長電話', d.parentPhone], ['處理狀態', status]]
    : [['類型', '報名整期'], ['應繳金額', `${amount}（早鳥：${d.earlyBird}）`], ['轉帳後 5 碼', d.last5 || '尚未轉帳（家長選擇稍後轉帳）'], ['處理狀態', status]];
  const body =
    (mailNote ? `<p style="background:#FDE2E4;border-radius:10px;padding:12px 14px;color:#9B1C2C"><b>⚠️ ${esc_(mailNote)}</b></p>` : '') +
    `<div style="background:#FFF6DD;border:1px solid #F7B21A;border-radius:12px;padding:4px 16px">${table_(key)}</div>` +
    '<div style="font-weight:700;color:#0B2A66;margin-top:18px">學生</div>' +
    table_([['學生姓名', d.studentName], ['年級', d.grade], ['就讀學校', d.school], ['程式基礎', d.codingLevel]]) +
    '<div style="font-weight:700;color:#0B2A66">家長與緊急聯絡</div>' +
    table_([['家長姓名', d.parentName], ['家長電話', d.parentPhone], ['家長 Email', d.parentEmail],
      ['緊急聯絡人', d.emergencyName], ['緊急聯絡電話', d.emergencyPhone]]) +
    '<div style="font-weight:700;color:#0B2A66">其他</div>' +
    table_([['從哪裡知道', d.source || '（未填）'], ['家長備註', d.note || '（未填）'], ['照片公開', d.photo], ['報名編號', d.refCode]]) +
    `<p><a href="${SpreadsheetApp.getActiveSpreadsheet().getUrl()}" style="display:inline-block;background:#0A63D6;color:#fff;text-decoration:none;padding:10px 18px;border-radius:10px">打開報名試算表</a></p>`;
  MailApp.sendEmail({ to: ADMIN_EMAIL, subject, htmlBody: shell_('有新報名！', body), name: SENDER_NAME, replyTo: d.parentEmail });
}

// 用瀏覽器開啟部署網址時顯示，方便確認有部署成功
function doGet() {
  return json_({ ok: true, message: 'AI 自走車常態班報名接收程式運作中' });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
