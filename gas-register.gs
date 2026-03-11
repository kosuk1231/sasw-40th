/**
 * 서울특별시사회복지사협회 창립 40주년 기념행사 - 사전 신청 Google Apps Script
 *
 * [사용 방법]
 * 1. Google Apps Script (script.google.com)에서 새 프로젝트 생성
 * 2. 이 코드를 붙여넣기
 * 3. 배포 > 새 배포 > 웹 앱 선택
 *    - 실행 대상: 본인
 *    - 액세스 권한: 모든 사용자
 * 4. 배포 후 생성된 URL을 register.html의 GAS_URL에 붙여넣기
 */

const SPREADSHEET_ID = '1xixpkKen7Ozky0carX6ZYhRSF6uRl5wpu10qrpM_o2s';

function doPost(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Sheet1') || ss.getSheets()[0];
    const data = JSON.parse(e.postData.contents);

    // 연번 자동 생성
    const lastRow = sheet.getLastRow();
    const nextNumber = lastRow <= 1 ? 1 : lastRow; // 헤더 행 고려

    sheet.appendRow([
      nextNumber,       // A: 연번
      data.name,        // B: 성함
      data.affiliation, // C: 소속
      data.phone,       // D: 연락처
      data.email,       // E: 이메일
      data.transport,   // F: 이동방법
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: '사전 신청 API가 정상 작동 중입니다.' }))
    .setMimeType(ContentService.MimeType.JSON);
}
