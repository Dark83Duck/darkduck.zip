/**
 * 1. 시트가 열릴 때 상단 메뉴를 자동으로 생성합니다.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🍎 초등 생활지도 도구') // 메뉴 이름
      .addItem('✅ 1단계: 시트 자동 준비하기', 'setupSheets') 
      .addSeparator()
      .addItem('🔗 2단계: 웹앱 주소 등록하기 (필수)', 'setWebAppUrl') // 수동 등록 메뉴 추가
      .addItem('📱 3단계: 웹앱 화면 열기', 'showWebAppUrl')
      .addToUi();
}

/**
 * 2. [명단] 및 [기록] 시트를 생성하고 제목 행을 설정합니다.
 */
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  // --- [명단] 시트 준비 ---
  var listSheet = ss.getSheetByName("명단");
  if (!listSheet) {
    listSheet = ss.insertSheet("명단");
  }
  listSheet.getRange("A1").setValue("이름");
  listSheet.getRange("A1").setBackground("#E1F5FE").setFontWeight("bold");
  
  // --- [기록] 시트 준비 ---
  var recordSheet = ss.getSheetByName("기록");
  if (!recordSheet) {
    recordSheet = ss.insertSheet("기록");
  }
  var headers = ["일시", "학생이름", "구분", "행동내용", "메모"];
  recordSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  recordSheet.getRange(1, 1, 1, headers.length).setBackground("#FFF9C4").setFontWeight("bold");
  
  ui.alert('시트 준비가 완료되었습니다! [명단] 시트에 학생 이름을 입력해 주세요.');
}

/**
 * 3. 발급받은 웹앱 주소를 구글 시트 백그라운드에 안전하게 저장합니다. (오류 해결의 핵심)
 */
function setWebAppUrl() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getDocumentProperties();
  var currentUrl = props.getProperty('WEB_APP_URL') || '';
  
  var response = ui.prompt(
    '웹앱 주소 등록',
    '가장 최근에 배포하여 복사한 웹앱 주소(URL)를 아래에 붙여넣어 주세요.\n\n(현재 등록된 주소: ' + (currentUrl ? currentUrl : '없음') + ')',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() == ui.Button.OK) {
    var inputUrl = response.getResponseText().trim();
    
    // 올바른 구글 스크립트 주소인지 검사
    if (inputUrl.indexOf('https://script.google.com/macros/s/') === 0) {
      props.setProperty('WEB_APP_URL', inputUrl); // 저장
      ui.alert('✅ 성공적으로 주소가 등록되었습니다!\n\n이제 [3단계: 웹앱 화면 열기] 메뉴를 클릭하면 정상적으로 연결됩니다.');
    } else {
      ui.alert('❌ 올바른 구글 웹앱 주소가 아닙니다. 주소를 다시 확인해 주세요.');
    }
  }
}

/**
 * 4. 등록된 웹앱 주소를 팝업창으로 예쁘게 띄워주는 안내 함수
 */
function showWebAppUrl() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getDocumentProperties();
  var url = props.getProperty('WEB_APP_URL'); // 저장된 주소를 가져옴
  
  if (url) {
    // 예쁜 팝업창(모달)을 만들기 위한 HTML 코드
    var htmlString = 
      '<div style="font-family: \'Noto Sans KR\', sans-serif; padding: 10px; text-align: center;">' +
      '  <p style="color: #4b5563; font-size: 14px; margin-bottom: 15px;">아래 버튼을 눌러 웹앱을 실행하거나 주소를 복사하세요.</p>' +
      '  <input type="text" id="urlField" value="' + url + '" style="width: 90%; padding: 10px; margin-bottom: 15px; border: 1px solid #d1d5db; border-radius: 6px; font-size: 12px; color: #374151; background-color: #f9fafb;" readonly>' +
      '  <div style="display: flex; gap: 10px; justify-content: center;">' +
      '    <button onclick="window.open(\'' + url + '\', \'_blank\')" style="flex: 1; padding: 10px; background-color: #2563eb; color: white; border: none; border-radius: 6px; font-weight: bold; cursor: pointer;">🌐 바로 열기</button>' +
      '    <button onclick="copyUrl()" style="flex: 1; padding: 10px; background-color: #f3f4f6; color: #374151; border: 1px solid #d1d5db; border-radius: 6px; font-weight: bold; cursor: pointer;">📋 주소 복사</button>' +
      '  </div>' +
      '  <p id="copyMsg" style="color: #059669; font-size: 13px; font-weight: bold; margin-top: 15px; display: none;">✅ 주소가 안전하게 복사되었습니다!</p>' +
      '  <script>' +
      '    function copyUrl() {' +
      '      var copyText = document.getElementById("urlField");' +
      '      copyText.select();' +
      '      document.execCommand("copy");' +
      '      document.getElementById("copyMsg").style.display = "block";' +
      '      setTimeout(function() { document.getElementById("copyMsg").style.display = "none"; }, 2000);' +
      '    }' +
      '  </script>' +
      '</div>';

    var htmlOutput = HtmlService.createHtmlOutput(htmlString)
        .setWidth(380)
        .setHeight(200);

    ui.showModalDialog(htmlOutput, '📱 딸깍학급메모 접속');
  } else {
    ui.alert('등록된 웹앱 주소가 없습니다.\n\n먼저 메뉴에서 [🔗 2단계: 웹앱 주소 등록하기]를 클릭하여 주소를 붙여넣어 주세요.');
  }
}

// --- 아래부터는 웹앱(클라이언트)과 통신하는 함수들입니다 ---

// 5. 웹앱 초기 화면을 로드
function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setTitle('딸깍학급메모');
}

// 6. 명단 시트에서 학생 이름 리스트 가져오기
function getStudentNames() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("명단");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange("A2:A" + lastRow).getValues();
  return data.map(function(row) { return row[0]; }).filter(String);
}

// 7. 기록 시트에서 전체 데이터 가져오기 (기록보기용)
function getAllRecords() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("기록");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];
  
  var records = [];
  for (var i = 1; i < data.length; i++) {
    records.push({
      date: data[i][0],
      student: data[i][1],
      type: data[i][2],
      event: data[i][3],
      memo: data[i][4]
    });
  }
  return records;
}

// 8. 새로운 기록 저장하기
function saveRecord(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("기록") || ss.insertSheet("기록");
  
  var timeZone = Session.getScriptTimeZone();
  var currentTime = Utilities.formatDate(new Date(), timeZone, "HH:mm:ss");
  var recordDateTime = data.date + " " + currentTime;
  
  sheet.appendRow([recordDateTime, data.student, data.type, data.event, data.memo]);
  return "success";
}

// 9. 기존 기록을 찾아서 수정하는 함수
function updateRecord(oldData, newData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("기록");
  
  var data = sheet.getDataRange().getDisplayValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === oldData.date && data[i][1] === oldData.student) {
      sheet.getRange(i + 1, 1, 1, 5).setValues([
        [newData.date, newData.student, newData.type, newData.event, newData.memo]
      ]);
      return "success";
    }
  }
  return "error"; 
}

// 10. 기존 기록을 완전히 삭제하는 함수
function deleteRecord(recordToDelete) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("기록");
  
  var data = sheet.getDataRange().getDisplayValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === recordToDelete.date && data[i][1] === recordToDelete.student) {
      sheet.deleteRow(i + 1);
      return "success";
    }
  }
  return "error"; 
}
