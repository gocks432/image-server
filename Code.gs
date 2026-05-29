// 1. 설정 (본인의 정보로 수정하세요)
const REPO_OWNER = "gocks432"; 
const REPO_NAME = "image-server";
const GITHUB_TOKEN = PropertiesService.getScriptProperties().getProperty('timer-app');

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Student Image Hosting')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 깃허브 업로드 및 시트 기록 함수 (파일명 강제 변환 버전)
function processUpload(obj) {
  try {
    const folderPath = "student_images";
    
    // 1. 파일 확장자 추출 (예: .jpg, .png)
    const fileExt = obj.fileName.split('.').pop();
    
    // 2. 파일명을 영어/숫자로 강제 변환
    // 형식: student_타임스탬프_랜덤값.확장자 (예: student_1715832000_123.png)
    const safeFileName = `student_${Date.now()}_${Math.floor(Math.random() * 1000)}.${fileExt}`;
    
    const url = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/${folderPath}/${safeFileName}`;
    
    const payload = {
      "message": "Upload by: " + obj.studentName,
      "content": obj.base64,
      "branch": "main"
    };

    const options = {
      "method": "put",
      "headers": { "Authorization": "token " + GITHUB_TOKEN },
      "contentType": "application/json",
      "payload": JSON.stringify(payload)
    };

    UrlFetchApp.fetch(url, options);

    // 3. Raw 이미지 링크 생성 (영문 파일명이라 절대 깨지지 않음)
    const rawUrl = `https://raw.githubusercontent.com/${REPO_OWNER}/${REPO_NAME}/main/${folderPath}/${safeFileName}`;

    // 4. 구글 시트에 기록 (원본 파일명은 참고용으로 기록)
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('List');
    sheet.appendRow([new Date(), obj.studentName, safeFileName, rawUrl, obj.fileName]);

    return { success: true, url: rawUrl };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// 5. 최근 업로드 목록 가져오기
function getHistory() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('List');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  // 최근 10개 기록 가져오기
  const startRow = Math.max(2, lastRow - 9);
  const data = sheet.getRange(startRow, 1, (lastRow - startRow + 1), 4).getValues();
  
  return data.reverse().map(row => ({
    student: row[1],
    fileName: row[2],
    url: row[3]
  }));
}

// 파일 삭제 함수
function deleteFile(obj) {
  const adminPw = PropertiesService.getScriptProperties().getProperty('delete-pw');
  
  // 1. 비밀번호 체크
  if (obj.inputPw !== adminPw) {
    return { success: false, error: "비밀번호가 일치하지 않습니다." };
  }

  try {
    const folderPath = "student_images";
    const apiUrl = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/${folderPath}/${obj.fileName}`;
    
    // 2. 깃허브에서 파일의 SHA 값 가져오기 (삭제 시 필요)
    const getOptions = {
      "method": "get",
      "headers": { "Authorization": "token " + GITHUB_TOKEN }
    };
    const response = UrlFetchApp.fetch(apiUrl, getOptions);
    const fileInfo = JSON.parse(response.getContentText());
    const sha = fileInfo.sha;

    // 3. 깃허브 파일 삭제 실행
    const deletePayload = {
      "message": "Delete file by admin",
      "sha": sha,
      "branch": "main"
    };
    const deleteOptions = {
      "method": "delete",
      "headers": { "Authorization": "token " + GITHUB_TOKEN },
      "contentType": "application/json",
      "payload": JSON.stringify(deletePayload)
    };
    UrlFetchApp.fetch(apiUrl, deleteOptions);

    // 4. 구글 시트에서 해당 행 삭제
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('List');
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][2] === obj.fileName) { // 파일명이 일치하는 행 찾기
        sheet.deleteRow(i + 1);
        break;
      }
    }

    return { success: true };
  } catch (e) {
    return { success: false, error: "삭제 실패: " + e.toString() };
  }
}
