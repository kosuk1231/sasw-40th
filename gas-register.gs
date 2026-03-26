/**
 * 서울특별시사회복지사협회 창립 40주년 기념행사 - 사전 신청 Google Apps Script
 * 스프레드시트 구조: 연번(A) | 성함(B) | 소속(C) | 연락처(D) | 이메일(E) | 편의제공 조사(F)
 */

// --- 스프레드시트 설정 ---
const SPREADSHEET_ID = '1xixpkKen7Ozky0carX6ZYhRSF6uRl5wpu10qrpM_o2s';

// --- Solapi 알림톡 설정 ---

// =====================================================
// 웹 앱: HTML 폼에서 POST 요청을 받아 시트 저장 + 알림톡 발송
// =====================================================
const MAX_ROWS = 231; // 헤더(1행) + 230명 = 231행까지 허용

function doPost(e) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('시트1') || ss.getSheets()[0];

    // 인원 제한 체크
    const currentLastRow = sheet.getLastRow();
    if (currentLastRow >= MAX_ROWS) {
      return ContentService
        .createTextOutput(JSON.stringify({
          result: 'closed',
          message: '죄송합니다. 선착순 접수가 마감되었습니다.'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const data  = JSON.parse(e.postData.contents);

    // 연번 자동 생성
    const lastRow    = sheet.getLastRow();
    const nextNumber = lastRow <= 1 ? 1 : lastRow; // 헤더 행 고려

    sheet.appendRow([
      nextNumber,       // A: 연번
      data.name,        // B: 성함
      data.affiliation, // C: 소속
      data.birthdate,   // D: 생년월일 (추가됨)
      data.phone,       // E: 연락처
      data.email,       // F: 이메일
      data.convenience, // G: 편의제공 조사
    ]);

    // 알림톡 발송 (성함, 연락처가 있을 때만)
    let alimtalkResult = 'skipped';
    if (data.name && data.phone) {
      const cleanMobile = data.phone.toString().replace(/-/g, '').replace(/\s/g, '');
      const variables = {
        "#{성함}": data.name
        // 템플릿에 변수가 더 있으면 아래처럼 추가하세요:
        // "#{소속}": data.affiliation || '',
        // "#{편의제공}": data.convenience || '',
      };
      const success = sendAlimtalkUsingSolapi({ variables, mobile: cleanMobile });
      alimtalkResult = success ? 'sent' : 'failed';
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success', alimtalk: alimtalkResult }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log('[doPost ERROR] ' + error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('시트1') || ss.getSheets()[0];
    const currentLastRow = sheet.getLastRow();
    const isClosed = currentLastRow >= MAX_ROWS;

    return ContentService
      .createTextOutput(JSON.stringify({
        status: isClosed ? 'closed' : 'open',
        current: currentLastRow - 1,  // 헤더 제외 현재 신청 수
        max: MAX_ROWS - 1,            // 최대 신청 가능 수 (150)
        message: isClosed ? '선착순 접수가 마감되었습니다.' : '접수 가능합니다.'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// =====================================================
// 스프레드시트 메뉴: 수동 발송용
// =====================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📣 알림톡 발송')
    .addItem('✅ 전체 발송', 'sendAllAlimtalk')
    .addItem('🔢 선택 행 발송', 'sendSelectedAlimtalk')
    .addToUi();
}

// 전체 발송
function sendAllAlimtalk() {
  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet   = ss.getSheetByName('시트1') || ss.getSheets()[0];
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('발송할 데이터가 없습니다.');
    return;
  }

  const data   = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const result = processRows(data);

  SpreadsheetApp.getUi().alert(
    `✅ 전체 발송 완료\n\n성공: ${result.success}건\n실패: ${result.fail}건\n건너뜀: ${result.skip}건`
  );
}

// 선택 행 발송
function sendSelectedAlimtalk() {
  const sheet     = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getActiveRange();

  if (!selection) {
    SpreadsheetApp.getUi().alert('발송할 행을 먼저 선택해주세요.');
    return;
  }
  if (selection.getRow() <= 1) {
    SpreadsheetApp.getUi().alert('헤더 행은 선택할 수 없습니다. 2행부터 선택해주세요.');
    return;
  }

  const data   = sheet.getRange(selection.getRow(), 1, selection.getNumRows(), 7).getValues();
  const result = processRows(data);

  SpreadsheetApp.getUi().alert(
    `✅ 선택 발송 완료\n\n성공: ${result.success}건\n실패: ${result.fail}건\n건너뜀: ${result.skip}건`
  );
}

// =====================================================
// 공통: 행 배열을 순회하며 알림톡 발송
// =====================================================
function processRows(rows) {
  const result = { success: 0, fail: 0, skip: 0 };

  rows.forEach((row, i) => {
    const name   = row[1] ? row[1].toString().trim() : ''; // B: 성함
    const mobile = row[4] ? row[4].toString().trim() : ''; // E: 연락처 (생년월일 D열 추가로 인덱스 4번이 됨)

    if (!name || !mobile) {
      result.skip++;
      return;
    }

    const cleanMobile = mobile.replace(/-/g, '').replace(/\s/g, '');
    if (!/^\d{10,11}$/.test(cleanMobile)) {
      Logger.log(`[SKIP] 잘못된 번호: ${mobile}`);
      result.fail++;
      return;
    }

    const variables = { "#{성함}": name };
    const success   = sendAlimtalkUsingSolapi({ variables, mobile: cleanMobile });

    success ? result.success++ : result.fail++;
    Utilities.sleep(300); // API 과부하 방지
  });

  return result;
}

// =====================================================
// Solapi 알림톡 발송
// =====================================================
function sendAlimtalkUsingSolapi({ variables, mobile }) {
  const url  = "https://api.solapi.com/messages/v4/send";
  const date = new Date().toISOString();
  const salt = Utilities.getUuid();
  const sig  = createSignature(API_SECRET, date, salt);

  const payload = {
    message: {
      to: mobile,
      from: SENDER_PHONE,
      kakaoOptions: {
        pfId: PF_ID,
        templateId: TEMPLATE_ID,
        variables: variables
      }
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "Authorization": `HMAC-SHA256 apiKey=${API_KEY}, date=${date}, salt=${salt}, signature=${sig}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response     = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log(`[ALIMTALK] ${mobile} → HTTP ${responseCode}: ${responseText}`);

    if (responseCode === 200) {
      const json = JSON.parse(responseText);
      return !(json.failedMessageList && json.failedMessageList.length > 0);
    }
    return false;

  } catch (e) {
    Logger.log(`[ALIMTALK ERROR] ${mobile}: ${e.toString()}`);
    return false;
  }
}

// =====================================================
// HMAC-SHA256 서명 생성
// =====================================================
function createSignature(apiSecret, date, salt) {
  const signatureBytes = Utilities.computeHmacSignature(
    Utilities.MacAlgorithm.HMAC_SHA_256,
    date + salt,
    apiSecret
  );
  return signatureBytes.map(b => {
    const h = (b < 0 ? b + 256 : b).toString(16);
    return h.length === 1 ? '0' + h : h;
  }).join('');
}
