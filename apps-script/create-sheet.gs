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

    // 링크가 있는 모든 사용자에게 편집 권한 (위탁업체 사장님이 접근 가능하도록)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);

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

// GET 요청 (테스트 + 대시보드 즉시 갱신)
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';

  if (action === 'refresh') {
    updateDashboard();
    return ContentService.createTextOutput(JSON.stringify({
      status: 'ok',
      message: '대시보드 갱신 완료'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: 'ok',
    message: '발주도우미 시트 생성 API가 정상 작동 중입니다.'
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 공유폴더 내 모든 시트를 "링크가 있는 모든 사용자 - 편집자"로 공유 설정
 * Apps Script 에디터에서 이 함수를 직접 실행하세요 (1회만)
 */
/**
 * 송장 대시보드 자동 갱신
 * 모든 업체 시트에서 주문수/송장수를 집계하여 마스터 시트의 "대시보드" 탭에 기록
 *
 * 설정 방법:
 * 1. Apps Script 에디터에서 updateDashboard 를 한 번 수동 실행 (권한 승인)
 * 2. 트리거 탭(시계 아이콘) → 트리거 추가
 *    - 실행할 함수: updateDashboard
 *    - 이벤트 소스: 시간 기반
 *    - 시간 간격: 15분마다 (10분 이하는 일일 할당량 초과 위험)
 */
var MASTER_SHEET_ID = '18_cKCFmBvUjIqlmdyidgYZ226uk4SRWYh-Rv0rWovyc';

function updateDashboard() {
  var startTime = new Date().getTime();
  var MAX_RUNTIME = 5 * 60 * 1000; // 5분 안전 타임아웃 (Apps Script 제한 6분)
  var masterSS = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var vendorSheet = masterSS.getSheetByName('시트1');
  if (!vendorSheet) {
    Logger.log('시트1 탭을 찾을 수 없습니다');
    return;
  }

  var vendorData = vendorSheet.getDataRange().getValues();
  var headers_row = vendorData[0];

  var nameIdx = headers_row.indexOf('업체명');
  var urlIdx = headers_row.indexOf('구글시트URL');
  var statusIdx = headers_row.indexOf('상태');
  if (nameIdx < 0 || urlIdx < 0) {
    Logger.log('업체명/구글시트URL 열을 찾을 수 없습니다');
    return;
  }

  var dashboard = masterSS.getSheetByName('대시보드');
  if (!dashboard) {
    dashboard = masterSS.insertSheet('대시보드');
  }

  // 기존 대시보드 데이터 캐시 (quota 에러/타임아웃 시 이전 값 유지)
  var prevData = {};
  try {
    var existingData = dashboard.getDataRange().getValues();
    for (var p = 1; p < existingData.length; p++) {
      if (existingData[p][0]) {
        prevData[existingData[p][0]] = existingData[p];
      }
    }
  } catch (e) { /* 첫 실행 시 무시 */ }

  // 활성 업체 목록 수집
  var vendors = [];
  for (var i = 1; i < vendorData.length; i++) {
    var name = vendorData[i][nameIdx];
    var url = vendorData[i][urlIdx];
    var status = statusIdx >= 0 ? vendorData[i][statusIdx] : '';
    if (!name || !url || status === '비활성') continue;

    var sheetId = extractSheetId(url);
    if (sheetId) {
      vendors.push({ name: name, sheetId: sheetId });
    }
  }

  // ===== 핵심 최적화: A열 + J열만 읽기 (batchGet) =====
  // 기존: A1:J100 (10열×100행 = 1000셀/업체) → 변경: A2:A100 + J2:J100 (2열×99행 = 198셀/업체)
  // 대역폭 80% 절감
  var token = ScriptApp.getOAuthToken();
  var requests = [];
  var rangeA = encodeURIComponent('시트1!A2:A100');
  var rangeJ = encodeURIComponent('시트1!J2:J100');
  for (var v = 0; v < vendors.length; v++) {
    var apiUrl = 'https://sheets.googleapis.com/v4/spreadsheets/' + vendors[v].sheetId
      + '/values:batchGet?ranges=' + rangeA + '&ranges=' + rangeJ
      + '&fields=valueRanges/values';
    requests.push({
      url: apiUrl,
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
  }

  // 3개씩 배치, 10초 대기 (대역폭 80% 감소로 대기시간도 축소 가능)
  var BATCH_SIZE = 3;
  var BATCH_DELAY = 10000;
  var responses = [];
  var totalBatches = Math.ceil(requests.length / BATCH_SIZE);
  var timedOut = false;
  Logger.log('병렬 요청 시작: ' + requests.length + '개 업체, ' + totalBatches + '배치 (batchGet 최적화)');

  for (var b = 0; b < totalBatches; b++) {
    // 안전 타임아웃 체크
    if (new Date().getTime() - startTime > MAX_RUNTIME) {
      Logger.log('⏱️ 5분 안전 제한 도달, 배치 ' + b + '/' + totalBatches + '에서 중단. 나머지는 캐시 사용');
      timedOut = true;
      break;
    }

    var batch = requests.slice(b * BATCH_SIZE, (b + 1) * BATCH_SIZE);
    try {
      var batchResp = UrlFetchApp.fetchAll(batch);
      for (var br = 0; br < batchResp.length; br++) {
        responses.push(batchResp[br]);
      }
    } catch (e) {
      Logger.log('⚠️ 배치 ' + b + ' 실패: ' + e.toString() + ' — 30초 대기 후 재시도');
      Utilities.sleep(30000);
      // 재시도 1회
      try {
        var retryResp = UrlFetchApp.fetchAll(batch);
        for (var rr = 0; rr < retryResp.length; rr++) {
          responses.push(retryResp[rr]);
        }
      } catch (e2) {
        Logger.log('❌ 배치 ' + b + ' 재시도 실패: ' + e2.toString() + ' — 캐시 사용');
        // 이 배치의 업체들은 null로 채워서 캐시 처리
        for (var skip = 0; skip < batch.length; skip++) {
          responses.push(null);
        }
      }
    }

    if (b < totalBatches - 1) {
      Utilities.sleep(BATCH_DELAY);
    }
  }

  var fetchTime = Math.round((new Date().getTime() - startTime) / 1000);
  Logger.log('배치 fetch 완료: ' + fetchTime + '초, 응답 ' + responses.length + '/' + vendors.length);

  // 응답 처리
  var dashHeaders = ['업체명', '전체주문', '송장완료', '미입력', '완료율', '최종갱신'];
  var results = [dashHeaders];
  var now = new Date();
  var quotaErrors = 0;
  var cacheUsed = 0;

  for (var v = 0; v < vendors.length; v++) {
    var vendorName = vendors[v].name;

    // 타임아웃으로 fetch 안 된 업체 → 캐시 사용
    if (v >= responses.length || responses[v] === null) {
      cacheUsed++;
      if (prevData[vendorName]) {
        results.push(prevData[vendorName]);
      } else {
        results.push([vendorName, 0, 0, 0, 0, now]);
      }
      continue;
    }

    try {
      var resp = responses[v];
      var code = resp.getResponseCode();

      // quota 에러 시 이전 캐시 값 유지
      if (code === 429 || (code >= 500 && code < 600)) {
        quotaErrors++;
        if (prevData[vendorName]) {
          results.push(prevData[vendorName]);
        } else {
          results.push([vendorName, -1, 0, 0, 0, now]);
        }
        continue;
      }

      if (code !== 200) {
        Logger.log('읽기 실패: ' + vendorName + ' - HTTP ' + code);
        results.push([vendorName, -1, 0, 0, 0, now]);
        continue;
      }

      var json = JSON.parse(resp.getContentText());
      var valueRanges = json.valueRanges || [];

      // A열 (주문 존재 여부) — 빈 행에서 멈춤
      var aValues = (valueRanges[0] && valueRanges[0].values) || [];
      var totalOrders = 0;
      for (var r = 0; r < aValues.length; r++) {
        if (aValues[r] && aValues[r][0] && String(aValues[r][0]).trim() !== '') {
          totalOrders++;
        } else {
          break;
        }
      }

      // J열 (송장번호) — totalOrders 범위 내에서만 카운트
      var jValues = (valueRanges[1] && valueRanges[1].values) || [];
      var invoiced = 0;
      for (var r = 0; r < Math.min(jValues.length, totalOrders); r++) {
        if (jValues[r] && jValues[r][0] && String(jValues[r][0]).trim() !== '') {
          invoiced++;
        }
      }

      var pending = totalOrders - invoiced;
      var rate = totalOrders > 0 ? Math.round(invoiced / totalOrders * 100) : 0;
      results.push([vendorName, totalOrders, invoiced, pending, rate, now]);

    } catch (e) {
      Logger.log('처리 실패: ' + vendorName + ' - ' + e.toString());
      if (prevData[vendorName]) {
        results.push(prevData[vendorName]);
      } else {
        results.push([vendorName, -1, 0, 0, 0, now]);
      }
    }
  }

  dashboard.clear();
  dashboard.getRange(1, 1, results.length, dashHeaders.length).setValues(results);
  var elapsed = Math.round((new Date().getTime() - startTime) / 1000);
  Logger.log('대시보드 갱신 완료: ' + (results.length - 1) + '개 업체, quota에러 ' + quotaErrors + '건, 캐시사용 ' + cacheUsed + '건, ' + elapsed + '초 소요'
    + (timedOut ? ' (타임아웃으로 일부 캐시)' : ''));
}

function extractSheetId(url) {
  var match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}


function shareAllSheets() {
  var folder = DriveApp.getFolderById(SHARED_FOLDER_ID);
  var files = folder.getFiles();
  var count = 0;
  var errors = [];

  while (files.hasNext()) {
    var file = files.next();
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
      count++;
      Logger.log('공유 완료: ' + file.getName());
    } catch (err) {
      errors.push(file.getName() + ': ' + err.toString());
      Logger.log('공유 실패: ' + file.getName() + ' - ' + err.toString());
    }
  }

  Logger.log('=== 완료: ' + count + '개 파일 공유 설정 ===');
  if (errors.length > 0) {
    Logger.log('실패: ' + errors.length + '개');
  }
}
