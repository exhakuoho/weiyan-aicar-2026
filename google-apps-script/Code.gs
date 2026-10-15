/**
 * AI 自走車常態班 報名 → Google 試算表 接收程式
 * 使用方式見同資料夾「設定說明.md」
 */

const SHEET_NAME = '報名資料';
// 早鳥以伺服器時間判斷（不信任訪客電腦的時間）
const EARLY_BIRD_DEADLINE = new Date('2026-11-30T23:59:59+08:00');
const PRICE_ORIGINAL = 4800;
const PRICE_EARLY = 4200;
// 有人報名時寄通知信給這個信箱；不需要就留空字串 ''
const NOTIFY_EMAIL = '';

const HEADERS = [
  '收到時間', '報名編號', '報名類型', '課程', '學生姓名', '年級', '就讀學校', '程式基礎',
  '家長姓名', '家長電話', '緊急聯絡人', '緊急聯絡電話', '付款方式', '應繳金額', '早鳥', '個資同意',
  '處理狀態', '備註'
];
const KEYS = [
  null, 'refCode', 'type', 'course', 'studentName', 'grade', 'school', 'codingLevel',
  'parentName', 'parentPhone', 'emergencyName', 'emergencyPhone', 'payment', 'amount', 'earlyBird', 'consent',
  null, null
];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#0B2A66').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    // 全部設為純文字，避免電話開頭的 0 被吃掉
    sh.getRange(2, 1, sh.getMaxRows() - 1, HEADERS.length).setNumberFormat('@');
  }
  return sh;
}

// 防止以 = + - @ 開頭的內容被試算表當成公式執行
function safe_(v) {
  const s = String(v == null ? '' : v).slice(0, 200);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    if (!data.studentName || !data.parentPhone) throw new Error('缺少必要欄位');
    if (data.type === '報名整期') {
      const early = new Date() <= EARLY_BIRD_DEADLINE;
      data.earlyBird = early ? '是' : '否';
      data.amount = String(early ? PRICE_EARLY : PRICE_ORIGINAL);
    }
    const sh = getSheet_();
    const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
    const row = KEYS.map((k, i) => {
      if (i === 0) return now;
      if (HEADERS[i] === '處理狀態') return '待聯繫';
      return k ? safe_(data[k]) : '';
    });
    sh.appendRow(row);

    if (NOTIFY_EMAIL) {
      MailApp.sendEmail(NOTIFY_EMAIL,
        `【新報名】AI 自走車常態班｜${data.type}｜${data.studentName}`,
        HEADERS.map((h, i) => `${h}：${row[i]}`).join('\n'));
    }
    return json_({ ok: true, refCode: data.refCode });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 用瀏覽器開啟部署網址時顯示，方便確認有部署成功
function doGet() {
  return json_({ ok: true, message: 'AI 自走車常態班報名接收程式運作中' });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
