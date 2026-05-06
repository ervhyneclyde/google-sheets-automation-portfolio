const CONFIG = {
  SHEET_NAME: 'Tab Name',
  SHEET_ID: SpreadsheetApp.getActiveSpreadsheet().getId(),
  WEB_APP_URL: "PASTE_YOUR_DEPLOYED_WEB_APP_URL_HERE",
  
  MAIN_RECEIVER: "email@gmail.com",
  CC_LIST: "member1@gmail.com",
  BCC_LIST: "member2@gmail.com",

  VC: ["Name"],
  AVCS: ["Name 1"],
  ASSOCS: ["Name 2"],

  HEADER_IMG: "IMAGE_URL_HERE",
  FOOTER_IMG: "IMAGE_URL_HERE"}

function sendHourlySummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
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

  const scanCount = Math.min(actualLastRow - 1, 50);
  const startRow = actualLastRow - scanCount + 1;
  const data = sheet.getRange(startRow, 1, scanCount, 25).getValues(); 
  const props = PropertiesService.getScriptProperties();
  let newSubmissions = [];

  for (let i = 0; i < data.length; i++) {
    let rowNum = startRow + i;
    let timestamp = data[i][0].toString();
    let secondCheckStatus = data[i][24] ? data[i][24].toString().trim() : ""; 

    if (timestamp.trim() === "") continue;

    let rowId = "PROD_V6_" + rowNum + "_" + timestamp;
    
    if (secondCheckStatus === "" && !props.getProperty(rowId)) {
      newSubmissions.push({
        row: rowNum,
        org: data[i][2],
        refNum: data[i][3],
        title: data[i][4],
        type: data[i][5],
        id: rowId
      });
    }
  }

  if (newSubmissions.length > 0) {
    compileAndSendSummary(newSubmissions);
  }
}

function compileAndSendSummary(submissions) {
  let tableRows = "";
  
  submissions.forEach(sub => {
    let assocButtons = CONFIG.ASSOCS.map(name => 
      `<a href="${CONFIG.WEB_APP_URL}?row=${sub.row}&name=${encodeURIComponent(name)}&role=ASSOC" style="background-color: #f1f1f1; color: #333; padding: 4px 8px; text-decoration: none; border-radius: 3px; font-size: 10px; margin: 2px 1px; display: inline-block; white-space: nowrap;">${name}</a>`
    ).join("");

    let vcButton = CONFIG.VC.map(name => 
      `<a href="${CONFIG.WEB_APP_URL}?row=${sub.row}&name=${encodeURIComponent(name)}&role=VC" style="background-color: #ffd700; color: #000; padding: 4px 8px; text-decoration: none; border-radius: 3px; font-size: 10px; margin: 2px 1px; display: inline-block; white-space: nowrap; font-weight: bold;">${name}</a>`
    ).join("");

    let avcButtons = CONFIG.AVCS.map(name => 
      `<a href="${CONFIG.WEB_APP_URL}?row=${sub.row}&name=${encodeURIComponent(name)}&role=AVC" style="background-color: #004d40; color: white; padding: 4px 8px; text-decoration: none; border-radius: 3px; font-size: 10px; margin: 2px 1px; display: inline-block; white-space: nowrap;">${name}</a>`
    ).join("");

    tableRows += `
      <tr>
        <td style="padding: 10px; border-bottom: 1px solid #eee; text-align: center; vertical-align: top;"><b>${sub.row}</b></td>
        <td style="padding: 10px; border-bottom: 1px solid #eee; vertical-align: top;">
          <b style="color: #004d40;">${sub.title}</b><br>
          <span style="font-size: 11px; color: #666;">${sub.org} | Ref: ${sub.refNum}</span><br>
          <span style="font-size: 11px; color: #999;">Type: ${sub.type}</span>
        </td>
        <td style="padding: 10px; border-bottom: 1px solid #eee; background-color: #fafafa; vertical-align: top;">
          <div style="font-size: 9px; font-weight: bold; margin-bottom: 5px; color: #555;">1ST CHECK (ASSOCS)</div>
          ${assocButtons}
        </td>
        <td style="padding: 10px; border-bottom: 1px solid #eee; vertical-align: top;">
          <div style="font-size: 9px; font-weight: bold; margin-bottom: 5px; color: #004d40;">2ND CHECK (VC/AVCS)</div>
          ${vcButton}
          <div style="margin-top: 5px;">${avcButtons}</div>
        </td>
      </tr>`;
  });

  const htmlBody = `
    <div style="max-width: 1000px; margin: auto; border: 1px solid #ddd; border-radius: 8px; font-family: Arial, sans-serif;">
      <img src="${CONFIG.HEADER_IMG}" style="width: 100%; display: block;">
      <div style="padding: 20px;">
        <h2 style="color: #004d40;">Submission Queue</h2>
        <table style="width: 100%; border-collapse: collapse; font-size: 12px;">
          <tr style="background-color: #f8f8f8;">
            <th style="padding: 10px; border-bottom: 2px solid #004d40; text-align: center;">Row</th>
            <th style="padding: 10px; border-bottom: 2px solid #004d40; text-align: left;">Activity Details</th>
            <th style="padding: 10px; border-bottom: 2px solid #004d40; text-align: left;">1st Check</th>
            <th style="padding: 10px; border-bottom: 2px solid #004d40; text-align: left;">2nd Check</th>
          </tr>
          ${tableRows}
        </table>
      </div>
      <img src="${CONFIG.FOOTER_IMG}" style="width: 100%; display: block;">
    </div>`;

  GmailApp.sendEmail(CONFIG.MAIN_RECEIVER, `Update: ${submissions.length} New Items`, "", {
    name: "Bot",
    cc: CONFIG.CC_LIST,
    bcc: CONFIG.BCC_LIST,
    htmlBody: htmlBody
  });

  const props = PropertiesService.getScriptProperties();
  submissions.forEach(sub => props.setProperty(sub.id, "true"));
}

function doGet(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const gid = sheet.getSheetId();
    const row = e.parameter.row;
    const name = e.parameter.name;
    const role = e.parameter.role;
    
    let targetColName, targetColDate;
    
    if (role === "ASSOC") {
      targetColName = 23;
      targetColDate = 24;
    } else {
      targetColName = 25;
      targetColDate = 26;
    }

    const cellName = sheet.getRange(row, targetColName);
    const cellDate = sheet.getRange(row, targetColDate);

    if (cellName.getValue().toString().trim() === "") {
      cellName.setValue(name);
      cellDate.setValue(new Date());
      
      const redirectUrl = `docs_link#gid=${gid}&range=A${row}`;
      
      return HtmlService.createHtmlOutput(`
        <html>
          <body style="font-family: Arial; text-align: center; padding-top: 50px;">
            <h2 style="color: #004d40;">SUCCESS!</h2>
            <p>Row ${row} has been reserved for <b>${name}</b>.</p>
            <p>Redirecting you to the sheet now...</p>
            <script>
              window.location.replace("${redirectUrl}");
            </script>
          </body>
        </html>
      `);
    } else {
      return HtmlService.createHtmlOutput("<h2 style='color:red; text-align:center;'>FAILED: This row is already claimed.</h2>");
    }
  } catch (err) {
    return HtmlService.createHtmlOutput("<h2>ERROR: System busy. Please try again.</h2>");
  } finally {
    lock.releaseLock();
  }
}

function resetMemory() {
  const props = PropertiesService.getScriptProperties();
  const keys = props.getKeys();
  keys.forEach(key => { if(key.includes("PROD_V6_")) props.deleteProperty(key); });
}
