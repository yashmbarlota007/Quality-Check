function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Yash Marketing - Master QC')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function GRANT_PERMISSIONS_RUN_ME_FIRST() {
  var root = DriveApp.getRootFolder();
  var resp = UrlFetchApp.fetch("https://www.google.com");
  Logger.log("Permissions OK! Root: " + root.getName() + ", HTTP: " + resp.getResponseCode());
}

// Kisi ek row ki photo ko Server se nikalne ka function
function fetchSingleRowPhoto(sheetName, rowId) {
  try {
    var config = QC_SETTINGS[sheetName];
    if (!config) return { ok: false, err: "Invalid sheet" };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName) || 
              ss.getSheetByName(sheetName.replace(/_/g, " ")) || 
              ss.getSheetByName(sheetName.replace(/\s+/g, "_"));
    var range = sheet.getRange(rowId, config.colPhoto);
    
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
        return { ok: false, err: "Drive Permission Error: " + driveErr.message };
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
        isTruncated: (link.length >= 49900),
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

  var jpgIndex = str.indexOf("/9j/");
  if (str.indexOf("data:image") === 0 || jpgIndex !== -1) {
    var b64Part = jpgIndex !== -1 ? str.substring(jpgIndex) : str.split(",").pop();
    var cleanB64 = b64Part.replace(/[^A-Za-z0-9+/]/g, "");
    while (cleanB64.length % 4 !== 0) cleanB64 += "=";

    return {
      type: "base64",
      len: str.length,
      data: "data:image/jpeg;base64," + cleanB64
    };
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
  if (!val) return "Time N/A";
  try {
    if (val instanceof Date) {
      return Utilities.formatDate(val, "Asia/Kolkata", "dd-MMM-yyyy hh:mm a");
    }
  } catch (e) {}
  return String(val).trim();
}

// "Attendance" sheet se Staff Directory (ID, Name, Phone) nikalna
function getStaffDirectory(ss) {
  var dir = { byId: {}, byName: {}, list: [] };
  try {
    var staffSheet = ss.getSheetByName("Attendance");
    if (staffSheet) {
      var sData = staffSheet.getDataRange().getValues();
      for (var i = 1; i < sData.length; i++) {
        var empId = String(sData[i][0] || "").trim().toUpperCase();
        var name = String(sData[i][1] || "").trim();
        var phone = String(sData[i][3] || "").replace(/\D/g, "");
        if (phone.length === 10) phone = "91" + phone;
        if (name) {
          var obj = { id: empId, name: name, phone: phone };
          dir.list.push(obj);
          dir.byName[name.toLowerCase()] = obj;
          if (empId) dir.byId[empId] = obj;
        }
      }
    }
  } catch (e) {}
  return dir;
}

// Kisi bhi row ke shuruati columns (A, B, C, D) se Staff ka Naam pehchanna
function detectStaffFromRow(row, sheetName, config, dir) {
  // 1. Pehle check karo agar row me kisi staff ka exact Name ya Employee ID (YM001...) likha hai
  var checkLimit = Math.min(row.length, 6);
  for (var c = 0; c < checkLimit; c++) {
    var val = String(row[c] || "").trim();
    if (!val || val.indexOf("data:image") === 0 || val.indexOf("http") === 0) continue;

    var upper = val.toUpperCase();
    var lower = val.toLowerCase();
    if (dir.byId[upper]) return dir.byId[upper];
    if (dir.byName[lower]) return dir.byName[lower];

    for (var s = 0; s < dir.list.length; s++) {
      var st = dir.list[s];
      var firstName = st.name.split(" ")[0].toLowerCase();
      if (firstName.length > 2 && lower === firstName) {
        return st;
      }
    }
  }

  // 2. Agar Attendance_Logs hai toh Column B (index 1) me direct naam hota hai
  if (sheetName === "Attendance_Logs") {
    return { name: String(row[1] || "N/A").trim(), phone: "" };
  }

  // 3. Task_Submissions me Column B, C, ya A me jo bhi staff name/role likha ho use uthao
  if (sheetName === "Task_Submissions") {
    var possibleCols = [1, 2, 0, 3]; // Column B, C, A, D
    for (var p = 0; p < possibleCols.length; p++) {
      var idx = possibleCols[p];
      var cell = row[idx];
      if (cell && !(cell instanceof Date)) {
        var strCell = String(cell).trim();
        // Skip numbers, dates, or frequency words like "Daily", "Weekly"
        if (strCell && !/^\d+$/.test(strCell) && !/^(daily|weekly|monthly|one time|urgent)$/i.test(strCell) && strCell.indexOf("2026") === -1) {
          return { name: strCell, phone: "" };
        }
      }
    }
    // Agar A-D me frequency/type likha hai (jaise Column B ya C), toh wahi dikha do
    var colB = String(row[1] || "").trim();
    var colC = String(row[2] || "").trim();
    var fallback = colB || colC || "Staff";
    return { name: fallback, phone: "" };
  }

  // 4. Stock Inward me agar staff column ho
  return { name: String(row[config.colEmp - 1] || "N/A").trim(), phone: "" };
}

var QC_SETTINGS = {
  "Stock Inward": {
    label: "Stock Inward",
    colTime: 1,     // Column A (Date/Time)
    colEmp: 11,     // Column K (Category)
    colTask: 10,    // Column J (Model Name)
    colSub: 3,      // Column C (Supplier / Party / Invoice Info)
    colPhoto: 19,   // Column S (Pavti Photo)
    colStatus: 21,  // Column U (QC Status)
    colRemark: 22,  // Column V (QC Remark)
    colQCTime: 23,  // Column W (QC Time)
    useRichText: false
  },
  "Task_Submissions": { 
    label: "Task Submission",
    colTime: 8,     // Column H (Submit Time)
    colEmp: 2,      // Column B (Staff Name)
    colTask: 5,     // Column E (Task Description - Task Kya Hai)
    colSub: 6,      // Column F (Status / Report Name - Konsa Task Hai)
    colDone: 9,     // Column I (Task Completed Status)
    colPhoto: 10,   // Column J (View Photo Button)
    colStatus: 12,  // Column L (QC Status)
    colRemark: 13,  // Column M (QC Remark)
    colQCTime: 14,  // Column N (QC Time)
    useRichText: true 
  },
  "Attendance_Logs": { 
    label: "Attendance Selfie",
    colTime: 1,     // Column A (Timestamp)
    colEmp: 2,      // Column B (Employee Name)
    colTask: 4,     // Column D (Punch In / Punch Out)
    colSub: 3,      // Column C (Employee ID - YM005)
    colPhoto: 6,    // Column F (Selfie Photo)
    colStatus: 7,   // Column G (QC Status)
    colRemark: 8,   // Column H (QC Remark)
    colQCTime: 9,   // Column I (QC Time)
    useRichText: true
  }
};

function getInitialMeta() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dir = getStaffDirectory(ss);
  return { staffList: dir.list };
}

function getPendingQC(sheetName) {
  sheetName = sheetName || "Attendance_Logs";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dir = getStaffDirectory(ss);

  if (sheetName === "Rejected_List") {
    var rejectedAll = [];
    for (var sName in QC_SETTINGS) {
      rejectedAll = rejectedAll.concat(getSheetItemsByStatus(ss, sName, "rejected", 15, dir));
    }
    return rejectedAll;
  }

  if (!QC_SETTINGS[sheetName]) throw new Error("Invalid Selection: " + sheetName);
  return getSheetItemsByStatus(ss, sheetName, "pending", 25, dir);
}

function getSheetItemsByStatus(ss, sheetName, targetMode, maxItems, dir) {
  var config = QC_SETTINGS[sheetName];
  var sheet = ss.getSheetByName(sheetName) || 
              ss.getSheetByName(sheetName.replace(/_/g, " ")) || 
              ss.getSheetByName(sheetName.replace(/\s+/g, "_")); 
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow === 0 || lastCol === 0) return [];

  var maxColNeeded = Math.max(config.colTime, config.colEmp, config.colTask, config.colSub || 1, config.colPhoto, config.colStatus, config.colRemark);
  var readCols = Math.min(lastCol, maxColNeeded);

  var data = sheet.getRange(1, 1, lastRow, readCols).getValues();
  
  var richTextLinks = null;
  var formulaLinks = null;
  if (config.useRichText && lastCol >= config.colPhoto) {
    try {
      richTextLinks = sheet.getRange(1, config.colPhoto, lastRow, 1).getRichTextValues();
      formulaLinks = sheet.getRange(1, config.colPhoto, lastRow, 1).getFormulas();
    } catch (e) {}
  }
  
  var items = [];

  for (var i = data.length - 1; i >= 0; i--) {
    if (!data[i] || data[i].length < 4) continue;

    var status = "";
    if (data[i].length >= config.colStatus) {
      status = String(data[i][config.colStatus - 1] || "").trim().toLowerCase();
    }
    
    var isMatch = (targetMode === "rejected") ? (status === "rejected") : (status === "pending" || status === "");
    if (!isMatch) continue;

    var photoLink = "";
    if (richTextLinks && richTextLinks[i] && richTextLinks[i][0]) {
      var rt = richTextLinks[i][0];
      photoLink = rt.getLinkUrl() || "";
      if (!photoLink) {
        var runs = rt.getRuns();
        for (var r = 0; r < runs.length; r++) {
          if (runs[r].getLinkUrl()) { photoLink = runs[r].getLinkUrl(); break; }
        }
      }
    }
    if (!photoLink && formulaLinks && formulaLinks[i] && formulaLinks[i][0]) {
      var fStr = String(formulaLinks[i][0]);
      var match = fStr.match(/https?:\/\/[^"'\s)]+/i);
      if (match) photoLink = match[0];
    }
    if (!photoLink) {
      photoLink = String(data[i][config.colPhoto - 1] || "").trim();
    }

    var photoInfo = processPhotoField(photoLink);

    if (photoInfo.type !== "none" && photoLink !== "" && photoLink.toLowerCase().indexOf("no photo") === -1 && i > 0) {
      var entryTime = formatTime(data[i][config.colTime - 1]);
      var staffInfo = detectStaffFromRow(data[i], sheetName, config, dir);

      var staffDisplay = staffInfo.name || "N/A";
      var mainTask = "";
      var subTask = "";

      if (sheetName === "Task_Submissions") {
        // Task Kya Hai (Column E) aur Konsa Task/Report Hai (Column F)
        mainTask = String(data[i][config.colTask - 1] || "N/A").trim();
        var colF = String(data[i][config.colSub - 1] || "").trim();
        var colI = config.colDone ? String(data[i][config.colDone - 1] || "").trim() : "";
        subTask = colF ? (colF + (colI ? " (" + colI + ")" : "")) : colI;
      } else if (sheetName === "Stock Inward" || sheetName.indexOf("Stock") !== -1 || sheetName.indexOf("Inward") !== -1) {
        staffDisplay = String(data[i][10] || data[i][4] || "Stock").trim();
        mainTask = String(data[i][11] || data[i][9] || "Model").trim();
        var invNo = String(data[i][3] || "").trim();
        subTask = "Inv #" + invNo + (data[i][9] ? " (" + data[i][9] + ")" : "");
        
        var invQty = parseFloat(data[i][12]) || 0;
        var actQty = parseFloat(data[i][13]) || 0;
        var shortQty = parseFloat(data[i][14]) || (invQty - actQty);

        var invRaw = data[i].length >= 19 ? data[i][18] : "";
        var invPhotoUrl = processPhotoField(invRaw).data || "";
        
        var boxRaw = data[i].length >= 20 ? data[i][19] : "";
        var boxPhotoUrl = processPhotoField(boxRaw).data || "";

        var pavRaw = data[i].length >= 21 ? data[i][20] : "";
        var pavtiPhotoUrl = processPhotoField(pavRaw).data || "";

        var prodRaw = data[i].length >= 22 ? data[i][21] : "";
        var prodPhotoUrl = processPhotoField(prodRaw).data || "";
      } else {
        // Attendance_Logs: Staff Name + Punch In/Out + Emp ID
        mainTask = String(data[i][config.colTask - 1] || "N/A").trim();
        subTask = config.colSub ? String(data[i][config.colSub - 1] || "").trim() : "";
      }

      var remark = data[i].length >= config.colRemark ? String(data[i][config.colRemark - 1] || "") : "";

      items.push({
        uid: sheetName.replace(/\s+/g, "_") + "_" + (i + 1),
        sheetName: sheetName,
        categoryLabel: config.label,
        rowId: i + 1, 
        time: entryTime,
        staffName: staffDisplay,
        staffPhone: staffInfo.phone || "",
        empName: mainTask,
        taskType: subTask,
        photoType: photoInfo.type,
        rawLen: photoInfo.len,
        fileId: photoInfo.fileId || "",
        photoUrl: (typeof invPhotoUrl !== "undefined" && invPhotoUrl) ? invPhotoUrl : (typeof prodPhotoUrl !== "undefined" && prodPhotoUrl ? prodPhotoUrl : photoInfo.data),
        invoicePhotoUrl: (typeof invPhotoUrl !== "undefined") ? invPhotoUrl : "",
        productPhotoUrl: (typeof prodPhotoUrl !== "undefined") ? prodPhotoUrl : "",
        boxPhotoUrl: (typeof boxPhotoUrl !== "undefined") ? boxPhotoUrl : "",
        pavtiPhotoUrl: (typeof pavtiPhotoUrl !== "undefined") ? pavtiPhotoUrl : "",
        invoiceQty: (typeof invQty !== "undefined") ? invQty : 0,
        actualQty: (typeof actQty !== "undefined") ? actQty : 0,
        shortageQty: (typeof shortQty !== "undefined") ? shortQty : 0,
        remark: remark
      });

      if (maxItems && items.length >= maxItems) break;
    }
  }
  return items;
}

function submitQC(sheetName, rowId, status, remark) {
  var config = QC_SETTINGS[sheetName];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName) || 
              ss.getSheetByName(sheetName.replace(/_/g, " ")) || 
              ss.getSheetByName(sheetName.replace(/\s+/g, "_"));
  
  if (sheet) {
    var qcTimestamp = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MMM-yyyy hh:mm:ss a");
    sheet.getRange(rowId, config.colStatus, 1, 3).setValues([[String(status), String(remark), qcTimestamp]]);
    return "Success";
  } else {
    throw new Error("Sheet nahi mili.");
  }
}