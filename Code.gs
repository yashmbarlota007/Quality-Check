// ==============================================================================
// YASH MARKETING - QUALITY CHECK (QC) ENTERPRISE SYSTEM (Code.gs)
// Full support for: Attendance Logs, Dual-Photo Stock Inward QC,
// Tasks QC, Auto Rejections Vault & Automated Staff Alert Engine
// ==============================================================================

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
      .setTitle('Yash Marketing - Master QC')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function GRANT_PERMISSIONS_RUN_ME_FIRST() {
  var root = DriveApp.getRootFolder();
  var resp = UrlFetchApp.fetch("https://www.google.com");
  Logger.log("Permissions OK! Root: " + root.getName() + ", HTTP: " + resp.getResponseCode());
  return "OK";
}

/**
 * Intelligent sheet retriever by checking aliases
 */
function getSheetSafely(ss, aliases) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (var a = 0; a < aliases.length; a++) {
    var search = aliases[a].trim().toLowerCase();
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
 * Server-side high-resolution photo resolver bypassing CORS & iframe restrictions
 */
function fetchSingleRowPhoto(sheetName, rowId, targetCol) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getTargetSheet(ss, sheetName);
    if (!sheet) return { ok: false, err: "Invalid sheet: " + sheetName };

    var col = targetCol || getPhotoColForSheet(sheet, sheetName);
    var range = sheet.getRange(rowId, col);
    
    var rawVal = String(range.getValue() || "").trim();
    var rtVal = range.getRichTextValue();
    var formulaVal = range.getFormula();

    var link = "";
    if (rtVal && rtVal.getLinkUrl()) {
      link = rtVal.getLinkUrl();
    } else if (rtVal) {
      var runs = rtVal.getRuns();
      for (var r = 0; r < runs.length; r++) {
        if (runs[r].getLinkUrl()) { link = runs[r].getLinkUrl(); break; }
      }
    }
    if (!link && formulaVal) {
      var m = String(formulaVal).match(/https?:\/\/[^"'\s)]+/i);
      if (m) link = m[0];
    }
    if (!link) link = rawVal;

    var fileId = "";
    var matchD = link.match(/\/d\/([a-zA-Z0-9_-]{15,})/);
    var matchId = link.match(/[?&]id=([a-zA-Z0-9_-]{15,})/);
    if (matchD && matchD[1]) fileId = matchD[1];
    else if (matchId && matchId[1]) fileId = matchId[1];

    if (fileId) {
      try {
        var file = DriveApp.getFileById(fileId);
        try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}
        var blob = file.getBlob();
        return {
          ok: true,
          source: "DriveApp",
          dataUrl: "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes())
        };
      } catch (driveErr) {
        try {
          var resp = UrlFetchApp.fetch("https://lh3.googleusercontent.com/d/" + fileId + "=w500", { muteHttpExceptions: true });
          if (resp.getResponseCode() === 200) {
            return {
              ok: true,
              source: "LH3 Fetch",
              dataUrl: "data:image/jpeg;base64," + Utilities.base64Encode(resp.getContent())
            };
          }
        } catch (fetchErr) {}
        return { ok: false, err: "Drive Error: " + driveErr.message };
      }
    }

    var jpgIdx = link.indexOf("/9j/");
    if (link.indexOf("data:image") === 0 || jpgIdx !== -1) {
      var b64 = jpgIdx !== -1 ? link.substring(jpgIdx) : link.split(",").pop();
      b64 = b64.replace(/[^A-Za-z0-9+/]/g, "");
      while (b64.length % 4 !== 0) b64 += "=";
      return {
        ok: true,
        source: "Base64",
        dataUrl: "data:image/jpeg;base64," + b64
      };
    }

    return { ok: false, err: "Cell Value: " + link.substring(0, 30) };
  } catch (err) {
    return { ok: false, err: err.message };
  }
}

function processPhotoField(rawVal) {
  if (!rawVal) return { type: "none", data: "", len: 0 };
  var str = String(rawVal).trim();

  // If multiple URLs separated by newline, take first
  if (str.indexOf("\n") !== -1) str = str.split("\n")[0].trim();
  else if (str.indexOf(" \n ") !== -1) str = str.split(" \n ")[0].trim();

  var jpgIndex = str.indexOf("/9j/");
  if (str.indexOf("data:image") === 0 || jpgIndex !== -1) {
    var b64Part = jpgIndex !== -1 ? str.substring(jpgIndex) : str.split(",").pop();
    var cleanB64 = b64Part.replace(/[^A-Za-z0-9+/]/g, "");
    while (cleanB64.length % 4 !== 0) cleanB64 += "=";
    return { type: "base64", len: str.length, data: "data:image/jpeg;base64," + cleanB64 };
  }

  var fileId = "";
  var matchD = str.match(/\/d\/([a-zA-Z0-9_-]{15,})/);
  var matchId = str.match(/[?&]id=([a-zA-Z0-9_-]{15,})/);
  if (matchD && matchD[1]) fileId = matchD[1];
  else if (matchId && matchId[1]) fileId = matchId[1];

  if (fileId) {
    return {
      type: "drive",
      fileId: fileId,
      len: str.length,
      data: "https://lh3.googleusercontent.com/d/" + fileId
    };
  }

  if (str.indexOf("http") === 0) {
    return { type: "url", data: str, len: str.length };
  }

  return { type: "invalid", data: str.substring(0, 30), len: str.length };
}

function formatTime(val) {
  if (!val) return "N/A";
  try {
    if (val instanceof Date) {
      return Utilities.formatDate(val, "Asia/Kolkata", "dd-MMM-yyyy hh:mm a");
    }
    var d = new Date(val);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, "Asia/Kolkata", "dd-MMM-yyyy hh:mm a");
    }
  } catch (e) {}
  return String(val).trim();
}

/**
 * Fetch Staff Directory from Staff Master or Attendance
 */
function getStaffDirectory(ss) {
  var dir = { byId: {}, byName: {}, list: [] };
  try {
    var staffSheet = getSheetSafely(ss, ["Staff Master", "Staff_Master", "Staff", "Attendance", "Employees"]);
    if (staffSheet) {
      var sData = staffSheet.getDataRange().getValues();
      var headers = sData[0];
      var nameCol = -1, phoneCol = -1, idCol = -1;

      for (var c = 0; c < headers.length; c++) {
        var h = String(headers[c]).toLowerCase().trim();
        if (nameCol === -1 && (h === "name" || h === "staff name" || h === "employee name")) nameCol = c;
        if (phoneCol === -1 && (h.indexOf("phone") !== -1 || h.indexOf("mobile") !== -1 || h.indexOf("whatsapp") !== -1)) phoneCol = c;
        if (idCol === -1 && (h.indexOf("employee id") !== -1 || h === "empid" || h === "emp id" || h === "id")) idCol = c;
      }
      if (nameCol === -1) nameCol = 1;
      if (phoneCol === -1) phoneCol = 3;
      if (idCol === -1) idCol = 0;

      for (var i = 1; i < sData.length; i++) {
        var empId = String(sData[i][idCol] || "").trim().toUpperCase();
        var name = String(sData[i][nameCol] || "").trim();
        var phone = String(sData[i][phoneCol] || "").replace(/\D/g, "");
        if (phone.length === 10) phone = "91" + phone;
        if (name) {
          var obj = { id: empId, name: name, phone: phone };
          dir.list.push(obj);
          dir.byName[name.toLowerCase()] = obj;
          if (empId) dir.byId[empId] = obj;
        }
      }
    }
  } catch (e) {
    Logger.log("getStaffDirectory error: " + e.message);
  }
  return dir;
}

function getInitialMeta() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dir = getStaffDirectory(ss);
  return { staffList: dir.list };
}

function getTargetSheet(ss, sheetName) {
  if (sheetName === "Attendance_Logs" || sheetName.indexOf("Attend") !== -1) {
    return getSheetSafely(ss, ["Attendance_Logs", "Attendance Logs", "Attendance", "Punch_Logs"]);
  } else if (sheetName === "Stock Inward" || sheetName.indexOf("Stock") !== -1 || sheetName.indexOf("Inward") !== -1) {
    return getSheetSafely(ss, ["Stock Inward", "Box Inward", "Stock_Inward", "Inward", "Inward Records"]);
  } else if (sheetName === "Task_Submissions" || sheetName.indexOf("Task") !== -1) {
    return getSheetSafely(ss, ["Task_Submissions", "Task Submissions", "Tasks", "Staff_Tasks"]);
  }
  return ss.getSheetByName(sheetName);
}

function getPhotoColForSheet(sheet, sheetName) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c]).toLowerCase();
    if (h.indexOf("photo") !== -1 || h.indexOf("selfie") !== -1 || h.indexOf("image") !== -1) {
      return c + 1;
    }
  }
  return (sheetName === "Attendance_Logs") ? 6 : 19;
}

/**
 * MAIN: GET PENDING QC ITEMS
 */
function getPendingQC(sheetName) {
  sheetName = sheetName || "Attendance_Logs";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dir = getStaffDirectory(ss);

  if (sheetName === "Rejected_List") {
    return getRejectedItemsAll(ss, dir);
  }

  var sheet = getTargetSheet(ss, sheetName);
  if (!sheet) {
    Logger.log("Sheet not found for: " + sheetName);
    return [];
  }

  var actualSheetName = sheet.getName();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol <= 0) return [];

  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = data[0];
  var colMap = mapSheetHeaders(headers, actualSheetName);

  var items = [];
  var limit = Math.max(1, lastRow - 120);

  // Traverse newest rows first
  for (var r = lastRow - 1; r >= limit; r--) {
    var row = data[r];
    var rowId = r + 1;

    // Check status
    var status = "";
    if (colMap.colStatus !== -1 && row[colMap.colStatus]) {
      status = String(row[colMap.colStatus]).trim().toLowerCase();
    }
    if (status === "approved" || status === "rejected") continue;

    var timeStr = formatTime(colMap.colTime !== -1 ? row[colMap.colTime] : new Date());
    var staffName = colMap.colStaff !== -1 && row[colMap.colStaff] ? String(row[colMap.colStaff]).trim() : "";
    var empId = colMap.colEmpId !== -1 && row[colMap.colEmpId] ? String(row[colMap.colEmpId]).trim() : "";
    
    // Resolve phone
    var staffPhone = "";
    if (empId && dir.byId[empId.toUpperCase()]) {
      staffPhone = dir.byId[empId.toUpperCase()].phone;
      if (!staffName) staffName = dir.byId[empId.toUpperCase()].name;
    } else if (staffName && dir.byName[staffName.toLowerCase()]) {
      staffPhone = dir.byName[staffName.toLowerCase()].phone;
    }

    // 1. STOCK INWARD WITH DUAL PHOTO & QUANTITIES
    if (actualSheetName.indexOf("Stock") !== -1 || actualSheetName.indexOf("Inward") !== -1) {
      var invNo = colMap.colInvoiceNo !== -1 && row[colMap.colInvoiceNo] ? String(row[colMap.colInvoiceNo]).trim() : "N/A";
      var invDate = colMap.colInvoiceDate !== -1 && row[colMap.colInvoiceDate] ? String(row[colMap.colInvoiceDate]).trim() : "";
      var brand = colMap.colBrand !== -1 && row[colMap.colBrand] ? String(row[colMap.colBrand]).trim() : "";
      var category = colMap.colCategory !== -1 && row[colMap.colCategory] ? String(row[colMap.colCategory]).trim() : "";
      var modelName = colMap.colModel !== -1 && row[colMap.colModel] ? String(row[colMap.colModel]).trim() : "Stock Item";
      
      var invQty = colMap.colInvQty !== -1 && !isNaN(parseFloat(row[colMap.colInvQty])) ? parseFloat(row[colMap.colInvQty]) : 0;
      var actQty = colMap.colActQty !== -1 && !isNaN(parseFloat(row[colMap.colActQty])) ? parseFloat(row[colMap.colActQty]) : 0;
      var shortQty = colMap.colShortQty !== -1 && !isNaN(parseFloat(row[colMap.colShortQty])) ? parseFloat(row[colMap.colShortQty]) : (invQty - actQty);

      var invPhotoInfo = colMap.colInvPhoto !== -1 ? processPhotoField(row[colMap.colInvPhoto]) : { data: "" };
      var prodPhotoInfo = colMap.colProdPhoto !== -1 ? processPhotoField(row[colMap.colProdPhoto]) : { data: "" };
      var boxPhotoInfo = colMap.colBoxPhoto !== -1 ? processPhotoField(row[colMap.colBoxPhoto]) : { data: "" };
      var pavtiPhotoInfo = colMap.colPavtiPhoto !== -1 ? processPhotoField(row[colMap.colPavtiPhoto]) : { data: "" };

      items.push({
        uid: "Stock_Inward_" + rowId,
        sheetName: actualSheetName,
        rowId: rowId,
        time: timeStr,
        staffName: staffName || category || brand || "Store Checker",
        staffPhone: staffPhone,
        empName: modelName,
        taskType: "Inv #" + invNo + (brand ? " (" + brand + ")" : ""),
        invoiceNo: invNo,
        invoiceDate: invDate,
        brand: brand,
        category: category,
        modelName: modelName,
        invoiceQty: invQty,
        actualQty: actQty,
        shortageQty: shortQty,
        invoicePhotoUrl: invPhotoInfo.data || "",
        productPhotoUrl: prodPhotoInfo.data || "",
        boxPhotoUrl: boxPhotoInfo.data || "",
        pavtiPhotoUrl: pavtiPhotoInfo.data || "",
        photoUrl: invPhotoInfo.data || prodPhotoInfo.data || "",
        status: "Pending"
      });
    }
    // 2. ATTENDANCE LOGS
    else if (actualSheetName.indexOf("Attend") !== -1) {
      var actionType = colMap.colAction !== -1 && row[colMap.colAction] ? String(row[colMap.colAction]).trim() : "Punch Log";
      var loc = colMap.colLoc !== -1 && row[colMap.colLoc] ? String(row[colMap.colLoc]).trim() : "";
      var selfieInfo = colMap.colPhoto !== -1 ? processPhotoField(row[colMap.colPhoto]) : { data: "" };

      items.push({
        uid: "Attendance_Logs_" + rowId,
        sheetName: actualSheetName,
        rowId: rowId,
        time: timeStr,
        staffName: staffName || "Staff Member",
        staffPhone: staffPhone,
        empName: actionType,
        taskType: (empId ? "Emp ID: " + empId : "") + (loc ? " | 📍 " + loc : ""),
        empId: empId,
        location: loc,
        photoUrl: selfieInfo.data || "",
        status: "Pending"
      });
    }
    // 3. TASK SUBMISSIONS
    else {
      var taskTitle = colMap.colTask !== -1 && row[colMap.colTask] ? String(row[colMap.colTask]).trim() : "Routine Task";
      var taskDetail = colMap.colDetail !== -1 && row[colMap.colDetail] ? String(row[colMap.colDetail]).trim() : "";
      var taskPhotoInfo = colMap.colPhoto !== -1 ? processPhotoField(row[colMap.colPhoto]) : { data: "" };

      items.push({
        uid: "Task_Submissions_" + rowId,
        sheetName: actualSheetName,
        rowId: rowId,
        time: timeStr,
        staffName: staffName || "Staff Member",
        staffPhone: staffPhone,
        empName: taskTitle,
        taskType: taskDetail || "Task Proof",
        photoUrl: taskPhotoInfo.data || "",
        status: "Pending"
      });
    }
  }

  return items;
}

/**
 * Dynamic Header Mapper ensuring resilience against column alterations
 */
function mapSheetHeaders(headers, sName) {
  var lowerName = sName.toLowerCase();
  var map = {
    colTime: -1, colStaff: -1, colEmpId: -1, colStatus: -1, colRemark: -1, colPhoto: -1,
    colInvoiceNo: -1, colInvoiceDate: -1, colBrand: -1, colCategory: -1, colModel: -1,
    colInvQty: -1, colActQty: -1, colShortQty: -1,
    colInvPhoto: -1, colProdPhoto: -1, colBoxPhoto: -1, colPavtiPhoto: -1,
    colAction: -1, colLoc: -1, colTask: -1, colDetail: -1
  };

  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c] || "").toLowerCase().trim();
    if (h.indexOf("qc status") !== -1 || h === "qc_status") map.colStatus = c;
    else if (map.colStatus === -1 && h === "status") map.colStatus = c;

    if (h.indexOf("qc remark") !== -1 || h.indexOf("reject reason") !== -1 || h === "qc_remark") map.colRemark = c;

    if (map.colTime === -1 && (h.indexOf("time") !== -1 || h.indexOf("date") !== -1 || h === "timestamp")) map.colTime = c;
    if (map.colStaff === -1 && (h.indexOf("staff") !== -1 || h === "name" || h === "checker" || h === "employee name")) map.colStaff = c;
    if (map.colEmpId === -1 && (h.indexOf("employee id") !== -1 || h === "empid" || h === "emp id" || h === "id")) map.colEmpId = c;

    // Stock Inward
    if (h.indexOf("invoice no") !== -1 || h === "invoiceno") map.colInvoiceNo = c;
    if (h.indexOf("invoice date") !== -1) map.colInvoiceDate = c;
    if (h.indexOf("brand") !== -1) map.colBrand = c;
    if (h.indexOf("category") !== -1) map.colCategory = c;
    if (h.indexOf("model") !== -1) map.colModel = c;
    if (h.indexOf("pieces per carton") !== -1 || h.indexOf("invoice qty") !== -1) map.colInvQty = c;
    if (h.indexOf("actual qty") !== -1 || h.indexOf("received qty") !== -1) map.colActQty = c;
    if (h.indexOf("shortage") !== -1) map.colShortQty = c;

    if (h.indexOf("invoice photo") !== -1) map.colInvPhoto = c;
    if (h.indexOf("product photo") !== -1 || h.indexOf("stock photo") !== -1) map.colProdPhoto = c;
    if (h.indexOf("box photo") !== -1) map.colBoxPhoto = c;
    if (h.indexOf("pavti") !== -1) map.colPavtiPhoto = c;

    // Attendance
    if (h.indexOf("action") !== -1 || h.indexOf("punch") !== -1 || h === "type") map.colAction = c;
    if (h.indexOf("location") !== -1 || h.indexOf("address") !== -1) map.colLoc = c;

    // Task
    if (h.indexOf("task") !== -1 && h.indexOf("status") === -1) map.colTask = c;
    if (h.indexOf("detail") !== -1 || h.indexOf("desc") !== -1) map.colDetail = c;

    if (map.colPhoto === -1 && (h.indexOf("photo") !== -1 || h.indexOf("selfie") !== -1 || h.indexOf("image") !== -1 || h.indexOf("url") !== -1)) {
      map.colPhoto = c;
    }
  }

  // Fallbacks for Attendance
  if (lowerName.indexOf("attend") !== -1) {
    if (map.colTime === -1) map.colTime = 0;
    if (map.colStaff === -1) map.colStaff = 1;
    if (map.colEmpId === -1) map.colEmpId = 2;
    if (map.colAction === -1) map.colAction = 3;
    if (map.colLoc === -1) map.colLoc = 4;
    if (map.colPhoto === -1) map.colPhoto = 5;
  }
  // Fallbacks for Stock Inward
  if (lowerName.indexOf("stock") !== -1 || lowerName.indexOf("inward") !== -1) {
    if (map.colTime === -1) map.colTime = 0;
    if (map.colInvoiceNo === -1) map.colInvoiceNo = 3;
    if (map.colStaff === -1) map.colStaff = 5;
    if (map.colCategory === -1) map.colCategory = 10;
    if (map.colModel === -1) map.colModel = 11;
    if (map.colInvQty === -1) map.colInvQty = 12;
    if (map.colActQty === -1) map.colActQty = 13;
    if (map.colShortQty === -1) map.colShortQty = 14;
    if (map.colInvPhoto === -1) map.colInvPhoto = 18;
    if (map.colBoxPhoto === -1) map.colBoxPhoto = 19;
    if (map.colPavtiPhoto === -1) map.colPavtiPhoto = 20;
    if (map.colProdPhoto === -1) map.colProdPhoto = 21;
  }

  return map;
}

/**
 * Ensure QC Columns exist in target sheet
 */
function ensureQCColumns(sheet) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var statusCol = -1, remarkCol = -1, timeCol = -1;

  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).toLowerCase().trim();
    if (h === "qc status" || h === "qc_status") statusCol = i + 1;
    if (h === "qc remark" || h === "reject reason" || h === "qc_remark") remarkCol = i + 1;
    if (h === "qc timestamp" || h === "qc verified at" || h === "qc_time") timeCol = i + 1;
  }

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
 * REJECTIONS VAULT: Centralized logging so all staff apps show live alert
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
    sheet.getRange(1, 1, 1, headers.length).setBackground("#991b1b").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setRowHeight(1, 35);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getRejectedItemsAll(ss, dir) {
  var list = [];
  try {
    var rejSheet = ss.getSheetByName("QC_Rejections") || ss.getSheetByName("QC Rejections");
    if (rejSheet && rejSheet.getLastRow() > 1) {
      var data = rejSheet.getDataRange().getValues();
      for (var r = data.length - 1; r >= 1; r--) {
        var row = data[r];
        list.push({
          uid: "Rej_" + r,
          rowId: row[3] || r,
          sheetName: row[2] || "QC_Rejections",
          time: formatTime(row[0]),
          staffName: row[4] || "Staff",
          staffPhone: row[5] || "",
          empName: row[6] || "Rejected Item",
          taskType: row[7] || "",
          remark: row[8] || "Rejected",
          photoUrl: "",
          status: "Rejected"
        });
      }
    }
  } catch (e) {
    Logger.log("getRejectedItemsAll error: " + e.message);
  }
  return list;
}

/**
 * SUBMIT QC VERDICT
 */
function submitQC(sheetName, rowId, status, remark) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getTargetSheet(ss, sheetName);
    if (!sheet) return { success: false, error: "Sheet not found" };

    rowId = parseInt(rowId, 10);
    if (isNaN(rowId) || rowId < 2) return { success: false, error: "Invalid row ID" };

    var cols = ensureQCColumns(sheet);
    var nowStr = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MM-yyyy hh:mm a");

    sheet.getRange(rowId, cols.statusCol).setValue(status);
    sheet.getRange(rowId, cols.remarkCol).setValue(remark || (status === "Approved" ? "OK" : "Rejected"));
    sheet.getRange(rowId, cols.timeCol).setValue(nowStr);

    try {
      var rowRange = sheet.getRange(rowId, 1, 1, Math.min(sheet.getLastColumn(), 15));
      if (status === "Approved") rowRange.setBackground("#f0fdf4");
      else if (status === "Rejected") rowRange.setBackground("#fef2f2");
    } catch (e) {}

    // Auto-log to QC_Rejections
    if (status === "Rejected") {
      try {
        var rejSheet = getOrCreateRejectionsSheet(ss);
        var rowVals = sheet.getRange(rowId, 1, 1, sheet.getLastColumn()).getValues()[0];
        var sName = sheet.getName();
        var staff = "", item = "", detail = "";

        if (sName.indexOf("Stock") !== -1 || sName.indexOf("Inward") !== -1) {
          staff = rowVals[5] || "Store Checker";
          item = rowVals[11] || "Stock Item";
          detail = "Inv #" + (rowVals[3] || "N/A") + " | Qty: " + (rowVals[12] || "0");
        } else if (sName.indexOf("Attend") !== -1) {
          staff = rowVals[1] || "Staff";
          item = rowVals[3] || "Attendance Punch";
          detail = "ID: " + (rowVals[2] || "");
        } else {
          staff = rowVals[1] || "Staff";
          item = rowVals[4] || "Task";
          detail = rowVals[5] || "";
        }

        rejSheet.appendRow([
          new Date(),
          Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MM-yyyy"),
          sName,
          rowId,
          staff,
          "",
          item,
          detail,
          remark,
          "Active (Action Required)",
          "WhatsApp Alert Sent"
        ]);
      } catch (logErr) {}
    }

    return { success: true, status: status, rowId: rowId };
  } catch (err) {
    return { success: false, error: err.message };
  }
}
