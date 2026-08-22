// 방문예약 폼 데이터 수신 → 구글시트 저장 + 이메일 알림
// haepal7979@gmail.com 앞으로 알림이 갑니다. 다른 메일로 받고 싶으면 NOTIFY_EMAIL만 바꾸세요.

const SHEET_NAME = '방문예약';
const NOTIFY_EMAIL = 'haepal7979@gmail.com';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet_();

    sheet.appendRow([
      new Date(),
      data.name || '',
      data.phone || '',
      data.utm_source || '',
      data.utm_medium || '',
      data.utm_campaign || '',
      data.gclid || '',
      data.source_slot || ''
    ]);

    sendNotification_(data);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['접수시각', '이름', '연락처', 'UTM_소스', 'UTM_매체', 'UTM_캠페인', 'GCLID', '유입폼위치']);
  }
  return sheet;
}

function sendNotification_(data) {
  const subject = '[방문예약 접수] ' + (data.name || '이름없음') + '님';
  const body =
    '새로운 방문예약이 접수됐습니다.\n\n' +
    '이름: ' + (data.name || '') + '\n' +
    '연락처: ' + (data.phone || '') + '\n' +
    '유입경로: ' + (data.utm_source || '-') + ' / ' + (data.utm_medium || '-') + ' / ' + (data.utm_campaign || '-') + '\n' +
    '접수시각: ' + new Date().toLocaleString('ko-KR');

  MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
}
