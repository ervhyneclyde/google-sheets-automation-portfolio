const CONFIG = {
  TARGET_SHEET_NAME: "YOUR_SHEET_NAME_HERE",
  WEB_APP_URL: "PASTE_YOUR_DEPLOYED_WEB_APP_URL_HERE",
  
  MAIN_RECIPIENT: "primary@example.com",
  CC_LIST: "user1@example.com, user2@example.com",
  
  VC_NAME: "VC_FULL_NAME",
  AVC_LIST: ["Name 1", "Name 2", "Name 3"], 
  
  HEADER_IMG: "YOUR_HEADER_IMAGE_URL",
  FOOTER_IMG: "YOUR_FOOTER_IMAGE_URL"
};

function notifyOnNewRow(e) {
  Utilities.sleep(1000); 

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME); 
  if (!sheet) return;

  const colA = sheet.getRange("A:A").getValues();
  let actualLastRow = 0;
  for (let i = colA.length - 1; i >= 0; i--) {
    if (colA[i][0] !== "" && colA[i][0] !== null) {
      actualLastRow = i + 1;
      break;
    }
  }

  if (actualLastRow <= 1) return;

  const startRow = Math.max(2, actualLastRow - 2); 
  const numRows = (actualLastRow - startRow) + 1;
  
  const fullDataRange = sheet.getRange(startRow, 1, numRows, 13).getValues();
  const scriptProperties = PropertiesService.getScriptProperties();

  for (let i = 0; i < fullDataRange.length; i++) {
    let rowNum = startRow + i;
    let timestamp = fullDataRange[i][0].toString();
    let columnMValue = fullDataRange[i][12] ? fullDataRange[i][12].toString().trim() : "";
    
    if (timestamp.trim() === "") continue;

    let rowId = "SECURE_V1_" + rowNum + "_" + timestamp;
    if (columnMValue === "" && !scriptProperties.getProperty(rowId)) {
      sendNotificationEmail(rowNum, rowId);
    }
  }
}

function sendNotificationEmail(activeRow, rowId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  const scriptProperties = PropertiesService.getScriptProperties();
  
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const newRowData = sheet.getRange(activeRow, 1, 1, lastColumn).getValues()[0];

  const deepLink = ss.getUrl() + "#gid=" + sheet.getSheetId() + "&range=" + activeRow + ":" + activeRow;
  
  let signOffHtml = `<p style="font-size: 12px; color: #555; margin-bottom: 5px;"><b>Reserve this Submission:</b></p>
    <a href="${CONFIG.WEB_APP_URL}?row=${activeRow}&name=${encodeURIComponent(CONFIG.VC_NAME)}" 
       style="background-color: #2d6a4f; color: white; padding: 5px 8px; text-decoration: none; border-radius: 3px; font-size: 10px; font-weight: bold; display: inline-block; margin: 2px;">VC: ${CONFIG.VC_NAME}</a><br>`;
  
  CONFIG.AVC_LIST.forEach(name => {
    signOffHtml += `<a href="${CONFIG.WEB_APP_URL}?row=${activeRow}&name=${encodeURIComponent(name)}" 
       style="background-color: #f1f1f1; color: #2d6a4f; padding: 5px 8px; text-decoration: none; border-radius: 3px; font-size: 10px; border: 1px solid #2d6a4f; display: inline-block; margin: 2px;">AVC: ${name}</a>`;
  });

  let htmlTableRows = "";
  for (let i = 0; i < 7; i++) {
    htmlTableRows += `<tr>
      <td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold; color: #333; width: 40%;">${headers[i] || 'Field'}</td>
      <td style="padding: 10px; border-bottom: 1px solid #eee; color: #555;">${newRowData[i] || 'N/A'}</td>
    </tr>`;
  }

  const htmlBody = `<div style="font-family: Arial, sans-serif; max-width: 600px; margin: auto; border: 1px solid #ddd; border-radius: 8px; overflow: hidden;">
      <img src="${CONFIG.HEADER_IMG}" style="width: 100%; display: block;">
      <div style="padding: 20px;">
        <h2 style="color: #2d6a4f; margin-top: 0;">New Submission Detected</h2>
        <table style="width: 100%; border-collapse: collapse; margin-top: 10px;">${htmlTableRows}</table>
        <div style="margin-top: 20px; padding: 15px; background-color: #f9f9f9; border-radius: 5px; border: 1px dashed #ccc;">${signOffHtml}</div>
        <div style="margin-top: 25px; text-align: center;"> 
          <a href="${deepLink}" style="background-color: #2d6a4f; color: white; padding: 12px 25px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;">Review in Spreadsheet</a> 
        </div> 
      </div> 
      <img src="${CONFIG.FOOTER_IMG}" style="width: 100%; display: block;"> 
    </div>`;

  try {
    GmailApp.sendEmail(CONFIG.MAIN_RECIPIENT, "Alert: Row " + activeRow, "", {
      name: "CSO Notification Bot",
      cc: CONFIG.CC_LIST,
      htmlBody: htmlBody
    });
    scriptProperties.setProperty(rowId, "true");
  } catch (e) {
    console.error("Gmail Error: " + e.message);
  }
}

function doGet(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
    const row = e.parameter.row;
    const name = e.parameter.name;
    const cellM = sheet.getRange(row, 13);
    const cellN = sheet.getRange(row, 14);

    const deepLink = ss.getUrl() + "#gid=" + sheet.getSheetId() + "&range=" + row + ":" + row;
    
    let currentVal = cellM.getValue().toString().trim();

    if (currentVal === "") {
      cellM.setValue(name);
      cellN.setValue(new Date());
      SpreadsheetApp.flush(); 
      
      return HtmlService.createHtmlOutput(`
        <html>
          <body style="font-family:
