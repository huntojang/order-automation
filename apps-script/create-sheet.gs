/**
 * 발주도우미 - 업체 전용 시트 자동 생성 스크립트
 * Google Apps Script Web App으로 배포
 *
 * POST 요청으로 업체명을 받으면:
 * 1. 구글 시트 생성 (발주도우미 공유폴더 안에)
 * 2. 헤더 + 서식 설정
 * 3. 서비스 계정에 편집 권한 부여
 * 4. 시트 URL 반환
 */

// 설정
var SHARED_FOLDER_ID = '1eaWpMk46ghXx-hNS4KpXUOSWxK7S3Iph';
var SERVICE_ACCOUNT_EMAIL = 'balju-bot@baljoo-helper-test.iam.gserviceaccount.com';
var HEADERS = ['주문일자', '주문번호', '수취인명', '연락처', '주소', '상품명', '옵션', '수량', '택배사', '송장번호'];

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var vendorName = data.vendor_name;

    if (!vendorName) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false, error: 'vendor_name is required'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 시트 생성
    var ss = SpreadsheetApp.create('[발주도우미] ' + vendorName + '_발주서');
    var sheet = ss.getSheets()[0];

    // 헤더 입력
    var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setValues([HEADERS]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#e6e6e6');

    // 택배사/송장번호 칸 노란색 (I, J 열 2~100행)
    sheet.getRange('I2:J100').setBackground('#ffffcc');

    // 열 너비 조정
    sheet.setColumnWidth(1, 100);  // 주문일자
    sheet.setColumnWidth(2, 120);  // 주문번호
    sheet.setColumnWidth(3, 80);   // 수취인명
    sheet.setColumnWidth(4, 120);  // 연락처
    sheet.setColumnWidth(5, 250);  // 주소
    sheet.setColumnWidth(6, 200);  // 상품명
    sheet.setColumnWidth(7, 150);  // 옵션
    sheet.setColumnWidth(8, 60);   // 수량
    sheet.setColumnWidth(9, 100);  // 택배사
    sheet.setColumnWidth(10, 130); // 송장번호

    // 공유폴더로 이동
    var file = DriveApp.getFileById(ss.getId());
    var folder = DriveApp.getFolderById(SHARED_FOLDER_ID);
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);

    // 서비스 계정에 편집 권한
    ss.addEditor(SERVICE_ACCOUNT_EMAIL);

    var url = ss.getUrl();

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      sheet_url: url,
      sheet_id: ss.getId(),
      vendor_name: vendorName
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false, error: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// GET 요청 (테스트용)
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'ok',
    message: '발주도우미 시트 생성 API가 정상 작동 중입니다.'
  })).setMimeType(ContentService.MimeType.JSON);
}
