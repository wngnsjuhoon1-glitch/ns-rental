/**
 * NS렌탈서비스 상담 신청 → 구글 시트 저장 스크립트
 * 구글 시트 > 확장 프로그램 > Apps Script 에 이 내용을 통째로 붙여넣고
 * "배포 > 새 배포 > 웹 앱"으로 배포하세요. (실행: 나, 액세스: 모든 사용자)
 */

// 새 신청이 들어오면 알림 메일을 받을 주소 (비워두면 메일 알림 없음)
const NOTIFY_EMAIL = '';
const SHEET_NAME = '상담신청';
const HEADERS = ['접수일시', '이름', '연락처', '상호명', '이메일', '관심분야', '예상예산', '문의내용', '개인정보동의', '접수페이지'];

function doPost(e) {
  const p = (e && e.parameter) || {};
  if (p.website) return json({ ok: true }); // 스팸봇 차단용 숨은 칸

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#eaf2ff');
      sh.setFrozenRows(1);
    }
    const fields = [p.name, p.tel, p.biz, p.email, p.cat, p.budget, p.msg, p.agree, p.page].map(clean);
    sh.appendRow([new Date()].concat(fields));

    if (NOTIFY_EMAIL) {
      const body = HEADERS.slice(1).map((h, i) => h + ': ' + fields[i]).join('\n');
      MailApp.sendEmail(NOTIFY_EMAIL, '[NS렌탈] 새 상담 신청 - ' + fields[0], body + '\n\n시트: ' + ss.getUrl());
    }
    return json({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// 긴 입력 자르기 + 시트 수식으로 실행되지 않도록 처리
function clean(v) {
  v = String(v || '').slice(0, 2000);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
