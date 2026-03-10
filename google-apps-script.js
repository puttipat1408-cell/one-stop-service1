/**
 * Google Apps Script for Provincial Treasury Service Statistics System
 * Optimized for Table-based UI and full field synchronization.
 */

const SHEET_NAME = 'Records';

/**
 * Initializes the spreadsheet and returns the active sheet.
 * Creates headers if the sheet is empty or updates them.
 */
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  
  // เพิ่ม Headers ลายเซ็นไว้ด้านหลัง ป้องกันคอลัมน์เก่าเคลื่อน
  const headers = [
    'ID', 'Service Date', 'Service ID', 
    'Counter Count', 'Phone Count', 'One Stop Service', 'Send to Expert',
    'Staff 1', 'Staff 2', 'Group Head', 'Provincial Treasury',
    'Satisfaction 1', 'Satisfaction 2', 'Satisfaction 3', 'Satisfaction 4', 'Satisfaction 5',
    'Other Text',
    'Recipient Civil Servant', 'Recipient Contractor', 'Recipient Pensioner', 'Recipient Welfare Card', 'Recipient General',
    'Sign Staff 1', 'Sign Staff 2', 'Sign Group Head', 'Sign Provincial Treasury',
    'Created At'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

/**
 * SERVING UI
 */
function doGet(e) {
  try {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('ระบบสรุปสถิติบริการคลังจังหวัด')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return ContentService.createTextOutput("Error loading UI: " + err.toString());
  }
}

/**
 * RPC FUNCTIONS (Backend logic)
 */

function getRecords() {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const headers = data[0];
    const rows = data.slice(1);
    
    return rows.map(row => {
      const obj = {};
      // Direct numeric mapping for safety
      obj.__backendId = row[0] ? row[0].toString() : '';
      obj.id = obj.__backendId;
      obj.service_date = formatDateForFrontend(row[1]);
      obj.service_id = row[2] ? row[2].toString() : '';
      obj.counter_count = Number(row[3]) || 0;
      obj.phone_count = Number(row[4]) || 0;
      obj.one_stop_service = Number(row[5]) || 0;
      obj.send_to_expert = Number(row[6]) || 0;
      obj.staff_1 = row[7] ? row[7].toString() : '';
      obj.staff_2 = row[8] ? row[8].toString() : '';
      obj.group_head = row[9] ? row[9].toString() : '';
      obj.provincial_treasury = row[10] ? row[10].toString() : '';
      obj.satisfaction_1 = Number(row[11]) || 0;
      obj.satisfaction_2 = Number(row[12]) || 0;
      obj.satisfaction_3 = Number(row[13]) || 0;
      obj.satisfaction_4 = Number(row[14]) || 0;
      obj.satisfaction_5 = Number(row[15]) || 0;
      obj.other_text = row[16] ? row[16].toString() : '';
      obj.recipient_civil_servant = Number(row[17]) || 0;
      obj.recipient_contractor = Number(row[18]) || 0;
      obj.recipient_pensioner = Number(row[19]) || 0;
      obj.recipient_welfare_card = Number(row[20]) || 0;
      obj.recipient_general = Number(row[21]) || 0;
      
      // ดึงข้อมูลลายเซ็นจากคอลัมน์ใหม่
      obj.sign_staff_1 = row[22] ? row[22].toString() : '';
      obj.sign_staff_2 = row[23] ? row[23].toString() : '';
      obj.sign_group_head = row[24] ? row[24].toString() : '';
      obj.sign_provincial_treasury = row[25] ? row[25].toString() : '';
      
      obj.created_at = row[26] ? row[26].toString() : '';
      return obj;
    });
  } catch (err) {
    console.error("getRecords Error:", err);
    return [];
  }
}

function createRecord(record) {
  try {
    const sheet = getSheet();
    const rowData = [
      record.id || Date.now().toString(),
      record.service_date,
      record.service_id,
      record.counter_count || 0,
      record.phone_count || 0,
      record.one_stop_service || 0,
      record.send_to_expert || 0,
      record.staff_1 || '',
      record.staff_2 || '',
      record.group_head || '',
      record.provincial_treasury || '',
      record.satisfaction_1 || 0,
      record.satisfaction_2 || 0,
      record.satisfaction_3 || 0,
      record.satisfaction_4 || 0,
      record.satisfaction_5 || 0,
      record.other_text || '',
      record.recipient_civil_servant || 0,
      record.recipient_contractor || 0,
      record.recipient_pensioner || 0,
      record.recipient_welfare_card || 0,
      record.recipient_general || 0,
      
      // บันทึกลายเซ็นลงคอลัมน์ใหม่
      record.sign_staff_1 || '',
      record.sign_staff_2 || '',
      record.sign_group_head || '',
      record.sign_provincial_treasury || '',
      
      record.created_at || new Date().toISOString()
    ];
    sheet.appendRow(rowData);
    return { isOk: true };
  } catch (err) {
    return { isOk: false, error: err.toString() };
  }
}

function updateRecord(record) {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const idToUpdate = record.id || record.__backendId;
    
    for (let i = 1; i < data.length; i++) {
       if (data[i][0].toString() === idToUpdate.toString()) {
         const rowIdx = i + 1;
         const rowData = [
            idToUpdate,
            record.service_date,
            record.service_id,
            record.counter_count || 0,
            record.phone_count || 0,
            record.one_stop_service || 0,
            record.send_to_expert || 0,
            record.staff_1 || '',
            record.staff_2 || '',
            record.group_head || '',
            record.provincial_treasury || '',
            record.satisfaction_1 || 0,
            record.satisfaction_2 || 0,
            record.satisfaction_3 || 0,
            record.satisfaction_4 || 0,
            record.satisfaction_5 || 0,
            record.other_text || '',
            record.recipient_civil_servant || 0,
            record.recipient_contractor || 0,
            record.recipient_pensioner || 0,
            record.recipient_welfare_card || 0,
            record.recipient_general || 0,
            
            // อัปเดตลายเซ็น
            record.sign_staff_1 || '',
            record.sign_staff_2 || '',
            record.sign_group_head || '',
            record.sign_provincial_treasury || '',
            
            new Date().toISOString() // updated_at / created_at
         ];
         sheet.getRange(rowIdx, 1, 1, rowData.length).setValues([rowData]);
         return { isOk: true };
       }
    }
    return { isOk: false, error: "Record not found" };
  } catch (err) {
    return { isOk: false, error: err.toString() };
  }
}

function deleteRecord(record) {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const idToDelete = record.id || record.__backendId;
    
    for (let i = 1; i < data.length; i++) {
       if (data[i][0].toString() === idToDelete.toString()) {
         sheet.deleteRow(i + 1);
         return { isOk: true };
       }
    }
    return { isOk: false, error: "Record not found" };
  } catch (err) {
    return { isOk: false, error: err.toString() };
  }
}

// Helper to format Date objects as YYYY-MM-DD
function formatDateForFrontend(val) {
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = ('0' + (val.getMonth() + 1)).slice(-2);
    const d = ('0' + val.getDate()).slice(-2);
    return `${y}-${m}-${d}`;
  }
  return val;
}
