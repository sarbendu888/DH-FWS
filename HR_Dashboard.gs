/**
 * HRMIS Portal - Web App Backend
 * Target Spreadsheet: HRMIS Lock
 */

const SPREADSHEET_ID = "1ZiIFxFHYgc6qqnDw5np8tro6NyA0Ws9iS5iH9bQPcj8"; // Active Sheet ID
const SHEET_NAME = "Sheet1";
const ADMIN_PIN = "3669"; // Required for Edit and Delete operations

function doGet() {
  return HtmlService.createTemplateFromFile("Index_HR")
    .evaluate()
    .setTitle("NHM HR , Hooghly")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Fetch all employee records formatted for UI consumption
 */
function getEmployeeData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return { success: true, headers: [], rows: [] };
    
    const headers = data[0].map(h => String(h).trim());
    const splIncIdx = headers.findIndex(h => h.toLowerCase() === "salary with spl. increment");
    const salarySplIncCol = splIncIdx !== -1 ? splIncIdx : 7;

    const rows = [];
    const today = new Date();

    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      // Skip completely blank rows
      if (!r[1] && !r[2]) continue;

      // Calculate Years of Service floating value
      let yearsWorking = 0;
      let dojStr = "";
      if (r[11]) { // Col L: Date of Joining
        const doj = new Date(r[11]);
        if (!isNaN(doj.getTime())) {
          const diffMs = today - doj;
          yearsWorking = Math.max(0, (diffMs / (1000 * 60 * 60 * 24 * 365.25))).toFixed(1);
          dojStr = doj.toISOString().split("T")[0];
        }
      }

      let dobStr = "";
      if (r[10]) { // Col K: DOB
        const dob = new Date(r[10]);
        if (!isNaN(dob.getTime())) {
          dobStr = dob.toISOString().split("T")[0];
        }
      }

      rows.push({
        rowIndex: i + 1, // 1-based row index in Google Sheets
        slNo: r[0],
        name: r[1] || "",
        code: r[2] || "",
        designation: r[3] || "",
        salaryMarch26: Number(r[4]) || 0,
        increment: Number(r[5]) || 0,
        revisedSalary: Number(r[6]) || 0,
        salarySplInc: Number(r[salarySplIncCol]) || 0,
        postingPlace: r[8] || "",
        mobile: r[9] ? String(r[9]).trim() : "",
        dob: dobStr,
        doj: dojStr,
        serviceCategory: r[12] || "",
        majorProgHead: r[13] || "",
        progHead: r[14] || "",
        postingLevel: r[15] || "",
        subDesignation: r[16] || "",
        tallyCode: r[17] || "",
        selfFund: Number(r[18]) || 0,
        childFund: Number(r[19]) || 0,
        arrear: Number(r[20]) || 0,
        remarks: r[21] || "",
        pTax: Number(r[22]) || 0,
        yearsWorking: Number(yearsWorking)
      });
    }

    return { success: true, rows: rows };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Updates an existing employee row (Protected by PIN)
 */
function updateEmployeeRecord(payload, pin) {
  if (String(pin).trim() !== ADMIN_PIN) {
    return { success: false, message: "Unauthorized: Invalid PIN code entered." };
  }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) {
    return { success: false, message: "Sheet is currently locked by another operation. Try again." };
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const rowIndex = Number(payload.rowIndex);

    if (!rowIndex || rowIndex < 2 || rowIndex > sheet.getLastRow()) {
      return { success: false, message: "Invalid row index." };
    }

    // Verify employee code match as safety check
    const currentCode = sheet.getRange(rowIndex, 3).getValue();
    if (currentCode && String(currentCode).trim() !== String(payload.code).trim()) {
      return { success: false, message: "Data conflict: Employee code has changed." };
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const splIncIdx = headers.findIndex(h => h.toLowerCase() === "salary with spl. increment");
    const salarySplIncCol1Based = splIncIdx !== -1 ? splIncIdx + 1 : 8;

    // Update mapped columns
    sheet.getRange(rowIndex, 2).setValue(payload.name);
    sheet.getRange(rowIndex, 4).setValue(payload.designation);
    sheet.getRange(rowIndex, 7).setValue(payload.revisedSalary);
    sheet.getRange(rowIndex, salarySplIncCol1Based).setValue(payload.salarySplInc);
    sheet.getRange(rowIndex, 9).setValue(payload.postingPlace);
    sheet.getRange(rowIndex, 10).setValue(payload.mobile);
    sheet.getRange(rowIndex, 14).setValue(payload.majorProgHead);
    sheet.getRange(rowIndex, 15).setValue(payload.progHead);
    sheet.getRange(rowIndex, 16).setValue(payload.postingLevel);
    sheet.getRange(rowIndex, 18).setValue(payload.tallyCode);
    sheet.getRange(rowIndex, 19).setValue(payload.selfFund);
    sheet.getRange(rowIndex, 20).setValue(payload.childFund);
    sheet.getRange(rowIndex, 21).setValue(payload.arrear);
    sheet.getRange(rowIndex, 22).setValue(payload.remarks);

    return { success: true, message: "Employee updated successfully!" };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes an employee row (Protected by PIN)
 */
function deleteEmployeeRecord(rowIndex, empCode, pin) {
  if (String(pin).trim() !== ADMIN_PIN) {
    return { success: false, message: "Unauthorized: Invalid PIN code entered." };
  }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) {
    return { success: false, message: "Sheet is currently busy. Please retry." };
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const targetRow = Number(rowIndex);

    if (targetRow < 2 || targetRow > sheet.getLastRow()) {
      return { success: false, message: "Row out of bounds." };
    }

    const checkCode = sheet.getRange(targetRow, 3).getValue();
    if (String(checkCode).trim() !== String(empCode).trim()) {
      return { success: false, message: "Safety check failed: Row code mismatch." };
    }

    sheet.deleteRow(targetRow);
    return { success: true, message: `Employee (${empCode}) deleted successfully.` };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}
