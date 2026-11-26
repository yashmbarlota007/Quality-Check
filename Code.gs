// ==============================================================================
// YASH MARKETING - QUALITY CHECK (QC) ENTERPRISE ENGINE (Code.gs)
// Features: Attendance QC, Dual-Photo Stock Inward QC (Invoice vs Stock),
//           Task QC, Auto Rejections Vault & Multi-Channel Staff Alert System
// ==============================================================================

/**
 * Web App Entry Point
 */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  
  if (e && e.parameter) {
    template.empId = e.parameter.empId || '';
    template.role = e.parameter.role || '';
    template.sheet = e.parameter.sheet || '';
  } else {
    template.empId = '';
    template.role = '';
    template.sheet = '';
  }
  
  return template.evaluate()
    .setTitle('Yash Marketing - Master QC')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * One-time Permission Granter for Google Services
 */
function GRANT_PERMISSIONS_RUN_ME_FIRST() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var user = Session.getActiveUser().getEmail();
  Logger.log("Permissions active for: " + user + " on " + ss.getName());
  return "OK";
}

/**
 * Helper to safely get sheet by multiple possible aliases case-insensitively
 */
function getSheetSafely(ss, targetNames) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (var t = 0; t < targetNames.length; t++) {
    var search = targetNames[t].trim().toLowerCase();
    for (var i = 0; i < sheets.length; i++) {
      var sName = sheets[i].getName().trim().toLowerCase();
      if (sName === search || sName.replace(/[\s_]+/g, "") === search.replace(/[\s_]+/g, "")) {
        return sheets[i];
      }
    }
  }
  return null;
}

/**
 * Format timestamps into human-readable IST strings
 */
function formatTime(val) {
  if (!val) return "N/A";
  try {
    if (val instanceof Date) {
      return Utilities.formatDate(val, "Asia/Kolkata", "dd MMM yyyy, hh:mm a");
    }
    var d = new Date(val);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, "Asia/Kolkata", "dd MMM yyyy, hh:mm a");
    }
    return String(val);
  } catch (e) {
    return String(val);
  }
}

/**
 * Converts Google Drive share links into direct downloadable / high-res preview links
 */
function processPhotoField(val) {
  if (!val || val === "-" || val === "No Photo") return "";
  var str = String(val).trim();
  
  // If multiple URLs separated by newline or comma, take the first one
  if (str.indexOf("\n") !== -1) {
    str = str.split("\n")[0].trim();
  } else if (str.indexOf(",") !== -1) {
    str = str.split(",")[0].trim();
  }
  
  // Extract Google Drive ID if present
  var driveIdMatch = str.match(/\/d\/([a-zA-Z0-9_-]+)/) || str.match(/id=([a-zA-Z0-9_-]+)/);
  if (driveIdMatch && driveIdMatch[1]) {
    var fileId = driveIdMatch[1];
    return "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w1000";
  }
  
  return str;
}

/**
 * Helper to get staff directory from Staff Master
 */
function getStaffDirectory(ss) {
  var staffList = [];
  try {
    var sheet = getSheetSafely(ss, ["Staff Master", "Staff_Master", "Staff", "Employees", "Users"]);
    if (!sheet) return staffList;
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return staffList;
    
    var headers = data[0];
    var nameCol = -1;
    var phoneCol = -1;
    var idCol = -1;
    
    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c]).toLowerCase().trim();
      if (nameCol === -1 && (h === "name" || h === "staff name" || h === "employee name")) nameCol = c;
      if (phoneCol === -1 && (h.indexOf("phone") !== -1 || h.indexOf("mobile") !== -1 || h.indexOf("contact") !== -1 || h.indexOf("whatsapp") !== -1)) phoneCol = c;
      if (idCol === -1 && (h === "employee id" || h === "empid" || h === "emp id" || h === "id")) idCol = c;
    }
    
    if (nameCol === -1) nameCol = 1;
    if (phoneCol === -1) phoneCol = 2;
    if (idCol === -1) idCol = 0;
    
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var name = row[nameCol] ? String(row[nameCol]).trim() : "";
      var phone = row[phoneCol] ? String(row[phoneCol]).trim() : "";
      var empId = row[idCol] ? String(row[idCol]).trim() : "";
      
      // Clean phone number (remove non-digits except +)
      var cleanPhone = phone.replace(/[^0-9]/g, "");
      if (cleanPhone.length === 10) cleanPhone = "91" + cleanPhone;
      
      if (name) {
        staffList.push({
          name: name,
          phone: cleanPhone || phone,
          empId: empId
        });
      }
    }
  } catch (e) {
    Logger.log("getStaffDirectory error: " + e.message);
  }
  return staffList;
}

/**
 * 1. GET INITIAL META (Staff List, Counts)
 */
function getInitialMeta() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var staffList = getStaffDirectory(ss);
    return {
      success: true,
      staffList: staffList
    };
  } catch (err) {
    return { success: false, error: err.message, staffList: [] };
  }
}

/**
 * Helper to ensure QC columns exist on target sheet and return their 1-based indices
 */
function ensureQCColumns(sheet) {
  var lastCol = sheet.getLastColumn();
  var headers = [];
  if (lastCol > 0) {
    headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  }
  
  var statusCol = -1;
  var remarkCol = -1;
  var timeCol = -1;
  
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).toLowerCase().trim();
    if (h === "qc status" || h === "status (qc)" || (statusCol === -1 && h === "qc_status")) statusCol = i + 1;
    if (h === "qc remark" || h === "qc remarks" || h === "reject reason" || (remarkCol === -1 && h === "qc_remark")) remarkCol = i + 1;
    if (h === "qc timestamp" || h === "qc verified at" || (timeCol === -1 && h === "qc_time")) timeCol = i + 1;
  }
  
  // If not found, append them
  if (statusCol === -1) {
    lastCol++;
    sheet.getRange(1, lastCol).setValue("QC Status").setFontWeight("bold").setBackground("#e0f2fe");
    statusCol = lastCol;
  }
  if (remarkCol === -1) {
    lastCol++;
    sheet.getRange(1, lastCol).setValue("QC Remark").setFontWeight("bold").setBackground("#e0f2fe");
    remarkCol = lastCol;
  }
  if (timeCol === -1) {
    lastCol++;
    sheet.getRange(1, lastCol).setValue("QC Timestamp").setFontWeight("bold").setBackground("#e0f2fe");
    timeCol = lastCol;
  }
  
  return { statusCol: statusCol, remarkCol: remarkCol, timeCol: timeCol };
}

/**
 * 2. GET PENDING QC ITEMS
 * Handles Attendance Logs, Stock Inward, Task Submissions, and Rejected items
 */
function getPendingQC(requestedSheet) {
  var items = [];
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var targetName = String(requestedSheet || "").trim();
    
    // CASE A: REJECTED LIST (Read from dedicated QC_Rejections or aggregate from sheets)
    if (targetName === "Rejected_List" || targetName.toLowerCase().indexOf("reject") !== -1) {
      return getRejectedItems(ss);
    }
    
    // Resolve sheet
    var aliases = [];
    if (targetName === "Attendance_Logs" || targetName.toLowerCase().indexOf("attend") !== -1) {
      aliases = ["Attendance_Logs", "Attendance Logs", "Attendance", "Punch_Logs", "Attendance_Data"];
    } else if (targetName === "Stock Inward" || targetName.toLowerCase().indexOf("stock") !== -1 || targetName.toLowerCase().indexOf("inward") !== -1) {
      aliases = ["Stock Inward", "Box Inward", "Stock_Inward", "Inward", "Inward Records"];
    } else {
      aliases = ["Task_Submissions", "Task Submissions", "Tasks", "Staff_Tasks", "Task_Logs"];
    }
    
    var sheet = getSheetSafely(ss, aliases);
    if (!sheet) {
      Logger.log("Sheet not found for: " + targetName);
      return [];
    }
    
    var actualSheetName = sheet.getName();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol <= 0) return [];
    
    var allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var headers = allData[0];
    
    // Find column mapping
    var colMap = mapSheetColumns(headers, actualSheetName);
    var staffDirectory = getStaffDirectory(ss);
    
    // Loop from bottom to top (most recent items first)
    // Limit to latest 100 rows to ensure blazing fast mobile performance
    var minRow = Math.max(1, lastRow - 120);
    for (var r = lastRow - 1; r >= minRow; r--) {
      var row = allData[r];
      var rowId = r + 1; // 1-based index
      
      // Check QC Status
      var currentStatus = "";
      if (colMap.statusCol !== -1 && row[colMap.statusCol] !== undefined) {
        currentStatus = String(row[colMap.statusCol]).trim();
      }
      
      // We only want items that are PENDING (not Approved, not Rejected)
      if (currentStatus.toLowerCase() === "approved" || currentStatus.toLowerCase() === "rejected") {
        continue;
      }
      
      var timeStr = formatTime(colMap.timeCol !== -1 ? row[colMap.timeCol] : new Date());
      var staffName = colMap.staffCol !== -1 && row[colMap.staffCol] ? String(row[colMap.staffCol]).trim() : "";
      var empId = colMap.idCol !== -1 && row[colMap.idCol] ? String(row[colMap.idCol]).trim() : "";
      
      // Match phone from staff directory if not found
      var staffPhone = "";
      for (var s = 0; s < staffDirectory.length; s++) {
        if ((staffName && staffDirectory[s].name.toLowerCase() === staffName.toLowerCase()) ||
            (empId && staffDirectory[s].empId && staffDirectory[s].empId.toLowerCase() === empId.toLowerCase())) {
          staffPhone = staffDirectory[s].phone;
          if (!staffName) staffName = staffDirectory[s].name;
          break;
        }
      }
      
      // DUAL PHOTO & DETAILED FIELDS FOR STOCK INWARD
      if (actualSheetName.toLowerCase().indexOf("stock") !== -1 || actualSheetName.toLowerCase().indexOf("inward") !== -1) {
        var invoiceNo = colMap.invoiceNoCol !== -1 && row[colMap.invoiceNoCol] ? String(row[colMap.invoiceNoCol]).trim() : "N/A";
        var invoiceDate = colMap.invoiceDateCol !== -1 && row[colMap.invoiceDateCol] ? String(row[colMap.invoiceDateCol]).trim() : "";
        var brand = colMap.brandCol !== -1 && row[colMap.brandCol] ? String(row[colMap.brandCol]).trim() : "";
        var category = colMap.categoryCol !== -1 && row[colMap.categoryCol] ? String(row[colMap.categoryCol]).trim() : "";
        var modelName = colMap.modelCol !== -1 && row[colMap.modelCol] ? String(row[colMap.modelCol]).trim() : "Item";
        
        var invoiceQty = colMap.invQtyCol !== -1 && !isNaN(parseFloat(row[colMap.invQtyCol])) ? parseFloat(row[colMap.invQtyCol]) : 0;
        var actualQty = colMap.actQtyCol !== -1 && !isNaN(parseFloat(row[colMap.actQtyCol])) ? parseFloat(row[colMap.actQtyCol]) : 0;
        var shortageQty = colMap.shortQtyCol !== -1 && !isNaN(parseFloat(row[colMap.shortQtyCol])) ? parseFloat(row[colMap.shortQtyCol]) : (invoiceQty - actualQty);
        var doaQty = colMap.doaQtyCol !== -1 && !isNaN(parseFloat(row[colMap.doaQtyCol])) ? parseFloat(row[colMap.doaQtyCol]) : 0;
        
        var invPhotoUrl = colMap.invPhotoCol !== -1 ? processPhotoField(row[colMap.invPhotoCol]) : "";
        var prodPhotoUrl = colMap.prodPhotoCol !== -1 ? processPhotoField(row[colMap.prodPhotoCol]) : "";
        var boxPhotoUrl = colMap.boxPhotoCol !== -1 ? processPhotoField(row[colMap.boxPhotoCol]) : "";
        var pavtiPhotoUrl = colMap.pavtiPhotoCol !== -1 ? processPhotoField(row[colMap.pavtiPhotoCol]) : "";
        
        items.push({
          uid: "Stock_Inward_" + rowId,
          rowId: rowId,
          sheetName: actualSheetName,
          time: timeStr,
          staffName: staffName || category || brand || "Checker",
          staffPhone: staffPhone,
          empName: modelName,
          taskType: "Inv #" + invoiceNo + (brand ? " (" + brand + ")" : ""),
          invoiceNo: invoiceNo,
          invoiceDate: invoiceDate,
          brand: brand,
          category: category,
          modelName: modelName,
          invoiceQty: invoiceQty,
          actualQty: actualQty,
          shortageQty: shortageQty,
          doaQty: doaQty,
          invoicePhotoUrl: invPhotoUrl,
          productPhotoUrl: prodPhotoUrl,
          boxPhotoUrl: boxPhotoUrl,
          pavtiPhotoUrl: pavtiPhotoUrl,
          photoUrl: invPhotoUrl || prodPhotoUrl || boxPhotoUrl,
          status: "Pending",
          remark: ""
        });
      }
      // ATTENDANCE LOGS
      else if (actualSheetName.toLowerCase().indexOf("attend") !== -1) {
        var actionType = colMap.actionCol !== -1 && row[colMap.actionCol] ? String(row[colMap.actionCol]).trim() : "Punch Log";
        var location = colMap.locCol !== -1 && row[colMap.locCol] ? String(row[colMap.locCol]).trim() : "";
        var selfieUrl = colMap.photoCol !== -1 ? processPhotoField(row[colMap.photoCol]) : "";
        
        items.push({
          uid: "Attendance_Logs_" + rowId,
          rowId: rowId,
          sheetName: actualSheetName,
          time: timeStr,
          staffName: staffName || "Staff Member",
          staffPhone: staffPhone,
          empName: actionType,
          taskType: (empId ? "ID: " + empId : "") + (location ? " | 📍 " + location : ""),
          location: location,
          empId: empId,
          photoUrl: selfieUrl,
          status: "Pending",
          remark: ""
        });
      }
      // STANDARD TASKS
      else {
        var taskTitle = colMap.taskCol !== -1 && row[colMap.taskCol] ? String(row[colMap.taskCol]).trim() : "Routine Task";
        var taskDetail = colMap.detailCol !== -1 && row[colMap.detailCol] ? String(row[colMap.detailCol]).trim() : "";
        var taskPhotoUrl = colMap.photoCol !== -1 ? processPhotoField(row[colMap.photoCol]) : "";
        
        items.push({
          uid: "Task_Submissions_" + rowId,
          rowId: rowId,
          sheetName: actualSheetName,
          time: timeStr,
          staffName: staffName || "Staff",
          staffPhone: staffPhone,
          empName: taskTitle,
          taskType: taskDetail || "Submission Proof",
          photoUrl: taskPhotoUrl,
          status: "Pending",
          remark: ""
        });
      }
    }
  } catch (err) {
    Logger.log("getPendingQC error: " + err.message);
  }
  return items;
}

/**
 * Intelligent Column Mapper for arbitrary sheet headers
 */
function mapSheetColumns(headers, sheetName) {
  var sName = sheetName.toLowerCase();
  var map = {
    timeCol: -1, staffCol: -1, idCol: -1, statusCol: -1, remarkCol: -1, photoCol: -1,
    invoiceNoCol: -1, invoiceDateCol: -1, brandCol: -1, categoryCol: -1, modelCol: -1,
    invQtyCol: -1, actQtyCol: -1, shortQtyCol: -1, doaQtyCol: -1,
    invPhotoCol: -1, prodPhotoCol: -1, boxPhotoCol: -1, pavtiPhotoCol: -1,
    actionCol: -1, locCol: -1, taskCol: -1, detailCol: -1
  };
  
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).toLowerCase().trim();
    var hClean = h.replace(/[^a-z0-9]/g, "");
    
    if (h.indexOf("qc status") !== -1 || h === "status (qc)" || h === "qc_status") map.statusCol = i;
    else if (map.statusCol === -1 && h === "status") map.statusCol = i;
    
    if (h.indexOf("qc remark") !== -1 || h.indexOf("reject reason") !== -1 || h === "qc_remark") map.remarkCol = i;
    
    if (map.timeCol === -1 && (h.indexOf("time") !== -1 || h.indexOf("date") !== -1 || h === "timestamp")) map.timeCol = i;
    
    if (map.staffCol === -1 && (h === "staff" || h === "staff name" || h === "employee name" || h === "name" || h === "staff / checker")) map.staffCol = i;
    if (map.idCol === -1 && (h === "employee id" || h === "emp id" || h === "empid" || h === "id")) map.idCol = i;
    
    // Stock Inward specific
    if (h.indexOf("invoice no") !== -1 || h === "invoiceno" || h === "bill no") map.invoiceNoCol = i;
    if (h.indexOf("invoice date") !== -1) map.invoiceDateCol = i;
    if (h.indexOf("brand") !== -1) map.brandCol = i;
    if (h.indexOf("category") !== -1) map.categoryCol = i;
    if (h.indexOf("model") !== -1) map.modelCol = i;
    
    if (h.indexOf("pieces per carton") !== -1 || h.indexOf("invoice qty") !== -1 || h === "carton qty") map.invQtyCol = i;
    if (h.indexOf("actual qty") !== -1 || h.indexOf("received qty") !== -1 || h === "found qty") map.actQtyCol = i;
    if (h.indexOf("shortage") !== -1) map.shortQtyCol = i;
    if (h.indexOf("doa") !== -1 || h.indexOf("damage") !== -1) map.doaQtyCol = i;
    
    if (h.indexOf("invoice photo") !== -1 || h === "invoicephotos") map.invPhotoCol = i;
    if (h.indexOf("product photo") !== -1 || h === "productphoto" || h.indexOf("stock photo") !== -1) map.prodPhotoCol = i;
    if (h.indexOf("box photo") !== -1 || h === "boxphotos") map.boxPhotoCol = i;
    if (h.indexOf("pavti") !== -1) map.pavtiPhotoCol = i;
    
    // Attendance specific
    if (h.indexOf("action") !== -1 || h.indexOf("punch type") !== -1 || h === "type") map.actionCol = i;
    if (h.indexOf("location") !== -1 || h.indexOf("address") !== -1) map.locCol = i;
    
    // Task specific
    if (h.indexOf("task") !== -1 && h.indexOf("status") === -1) map.taskCol = i;
    if (h.indexOf("detail") !== -1 || h.indexOf("description") !== -1 || h.indexOf("remark") !== -1) map.detailCol = i;
    
    // Generic Photo
    if (h.indexOf("photo") !== -1 || h.indexOf("image") !== -1 || h.indexOf("selfie") !== -1 || h.indexOf("url") !== -1) {
      if (map.photoCol === -1) map.photoCol = i;
    }
  }
  
  // Positional fallbacks for Attendance_Logs if headers were non-standard
  if (sName.indexOf("attend") !== -1) {
    if (map.timeCol === -1) map.timeCol = 0;
    if (map.staffCol === -1) map.staffCol = 1;
    if (map.idCol === -1) map.idCol = 2;
    if (map.actionCol === -1) map.actionCol = 3;
    if (map.locCol === -1) map.locCol = 4;
    if (map.photoCol === -1) map.photoCol = 5;
  }
  
  return map;
}

/**
 * 3. GET OR CREATE DEDICATED QC_REJECTIONS SHEET TAB
 * Solves the staff alert problem: Any app/dashboard can query this sheet
 */
function getOrCreateRejectionsSheet(ss) {
  var sheetName = "QC_Rejections";
  var sheet = ss.getSheetByName(sheetName) || ss.getSheetByName("QC Rejections");
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    var headers = [
      "Timestamp", "Date", "Source Sheet", "Row ID", "Staff Name", 
      "Staff Phone", "Item / Task", "Detail / Invoice", "Reject Reason", 
      "Resolution Status", "Notified Via"
    ];
    sheet.appendRow(headers);
    
    var hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setBackground("#991b1b").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setRowHeight(1, 35);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 4. GET REJECTED ITEMS
 */
function getRejectedItems(ss) {
  var rejected = [];
  try {
    var rejSheet = ss.getSheetByName("QC_Rejections") || ss.getSheetByName("QC Rejections");
    if (rejSheet && rejSheet.getLastRow() > 1) {
      var data = rejSheet.getDataRange().getValues();
      for (var r = data.length - 1; r >= 1; r--) {
        var row = data[r];
        rejected.push({
          uid: "Rej_" + r,
          rowId: row[3] || r,
          sheetName: row[2] || "QC_Rejections",
          time: formatTime(row[0]),
          staffName: row[4] || "Staff",
          staffPhone: row[5] || "",
          empName: row[6] || "Rejected Item",
          taskType: row[7] || "",
          remark: row[8] || "Rejected by QC",
          photoUrl: "",
          status: "Rejected"
        });
      }
    }
  } catch (e) {
    Logger.log("getRejectedItems error: " + e.message);
  }
  return rejected;
}

/**
 * 5. SUBMIT QC ACTION (Approve / Reject)
 */
function submitQC(sheetName, rowId, status, remark) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName) || getSheetSafely(ss, [sheetName]);
    if (!sheet) return { success: false, error: "Sheet " + sheetName + " not found" };
    
    rowId = parseInt(rowId, 10);
    if (isNaN(rowId) || rowId < 2) return { success: false, error: "Invalid row ID" };
    
    var qcCols = ensureQCColumns(sheet);
    var nowStr = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MM-yyyy hh:mm a");
    
    sheet.getRange(rowId, qcCols.statusCol).setValue(status);
    sheet.getRange(rowId, qcCols.remarkCol).setValue(remark || (status === "Approved" ? "OK" : "Rejected"));
    sheet.getRange(rowId, qcCols.timeCol).setValue(nowStr);
    
    // Visual indicator on target row
    try {
      var rowRange = sheet.getRange(rowId, 1, 1, Math.min(sheet.getLastColumn(), 15));
      if (status === "Approved") {
        rowRange.setBackground("#f0fdf4"); // subtle green
      } else if (status === "Rejected") {
        rowRange.setBackground("#fef2f2"); // subtle red
      }
    } catch (bgErr) {}
    
    // If Rejected, ALSO log in QC_Rejections sheet for centralized staff portal visibility!
    if (status === "Rejected") {
      try {
        var rejSheet = getOrCreateRejectionsSheet(ss);
        var rowValues = sheet.getRange(rowId, 1, 1, sheet.getLastColumn()).getValues()[0];
        
        var staffName = "";
        var itemDesc = "";
        var detail = "";
        var staffPhone = "";
        
        // Extract details for logging
        if (sheetName.toLowerCase().indexOf("stock") !== -1 || sheetName.toLowerCase().indexOf("inward") !== -1) {
          staffName = rowValues[5] || rowValues[9] || "Store Checker";
          itemDesc = rowValues[11] || "Stock Item";
          detail = "Inv #" + (rowValues[3] || "N/A") + " | Qty: " + (rowValues[12] || "0");
        } else if (sheetName.toLowerCase().indexOf("attend") !== -1) {
          staffName = rowValues[1] || "Staff";
          itemDesc = rowValues[3] || "Attendance Punch";
          detail = "ID: " + (rowValues[2] || "");
        } else {
          staffName = rowValues[1] || "Staff";
          itemDesc = rowValues[3] || "Task";
          detail = rowValues[4] || "";
        }
        
        var dateOnly = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MM-yyyy");
        rejSheet.appendRow([
          new Date(),
          dateOnly,
          sheetName,
          rowId,
          staffName,
          staffPhone,
          itemDesc,
          detail,
          remark,
          "Active (Action Required)",
          "WhatsApp Alert Sent"
        ]);
      } catch (logErr) {
        Logger.log("Logging to QC_Rejections failed: " + logErr.message);
      }
    }
    
    return { success: true, status: status, rowId: rowId };
  } catch (err) {
    Logger.log("submitQC error: " + err.message);
    return { success: false, error: err.message };
  }
}

/**
 * 6. FETCH SINGLE ROW PHOTO (Bypasses CORS / Google Drive iframe restrictions)
 */
function fetchSingleRowPhoto(sheetName, rowId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName) || getSheetSafely(ss, [sheetName]);
    if (!sheet) return { ok: false, err: "Sheet not found" };
    
    rowId = parseInt(rowId, 10);
    var rowValues = sheet.getRange(rowId, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Find photo URL in row
    var photoUrl = "";
    for (var c = 0; c < rowValues.length; c++) {
      var val = String(rowValues[c] || "").trim();
      if (val.indexOf("drive.google.com") !== -1 || val.indexOf("http") === 0) {
        photoUrl = val;
        break;
      }
    }
    
    if (!photoUrl) return { ok: false, err: "No photo link in row" };
    
    // Extract file ID
    var match = photoUrl.match(/\/d\/([a-zA-Z0-9_-]+)/) || photoUrl.match(/id=([a-zA-Z0-9_-]+)/);
    if (match && match[1]) {
      var file = DriveApp.getFileById(match[1]);
      var blob = file.getBlob();
      var b64 = Utilities.base64Encode(blob.getBytes());
      var mime = blob.getContentType() || "image/jpeg";
      return { ok: true, dataUrl: "data:" + mime + ";base64," + b64 };
    }
    
    return { ok: true, dataUrl: photoUrl };
  } catch (err) {
    return { ok: false, err: err.message };
  }
}
