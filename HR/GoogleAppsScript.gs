/**
 * 祥麟紙器 考勤打卡系統 - Google Apps Script (GAS) 雙向自動同步腳本
 * 
 * 包含優化之日期與時間格式化處理：
 * 1. 日期一律格式化為 YYYY-MM-DD
 * 2. 上下班時間一律格式化為 HH:mm:ss (去除 1899 年、時區與文字)
 */

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // === 處理 1: 更新 / 新增員工資料 (action === 'saveEmployee') ===
    if (data.action === 'saveEmployee') {
      var empSheet = getOrInitEmpSheet(ss);
      var emp = data.employee;
      var lastRow = empSheet.getLastRow();
      var foundRow = -1;

      if (lastRow > 1) {
        var ids = empSheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (var i = 0; i < ids.length; i++) {
          if (ids[i][0] === emp.id) {
            foundRow = i + 2;
            break;
          }
        }
      }

      var rowData = [
        emp.id, emp.username, emp.pass, emp.name, emp.role, emp.dept, 
        emp.isAdmin ? "主管" : "員工", emp.avatar || "👨‍💼", 
        emp.shift || "09:00 - 18:00", emp.annualLeave || 10, 
        emp.sickLeave || 30, emp.compLeave || 0
      ];

      if (foundRow > -1) {
        empSheet.getRange(foundRow, 1, 1, 12).setValues([rowData]);
      } else {
        empSheet.appendRow(rowData);
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "員工資料已同步至 Google 試算表！"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // === 處理 2: 考勤打卡同步 ===
    var sheet = ss.getSheetByName("考勤紀錄");
    if (!sheet) {
      sheet = ss.insertSheet("考勤紀錄");
      sheet.appendRow([
        "打卡ID", "員工編號", "員工姓名", "部門", "日期", 
        "上班時間", "下班時間", "累計工時(小時)", 
        "考勤狀態", "遲到(分鐘)", "打卡備註", "同步時間"
      ]);
      var headerRange = sheet.getRange(1, 1, 1, 12);
      headerRange.setBackground("#1e3a8a");
      headerRange.setFontColor("#ffffff");
      headerRange.setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    var logId = data.id || "";
    var empId = data.empId || "";
    var empName = data.empName || "";
    var dept = data.dept || "";
    var date = cleanDateStr(data.date);
    var checkIn = cleanTimeStr(data.checkIn);
    var checkOut = cleanTimeStr(data.checkOut);
    var hours = data.hours || 0;
    var status = data.status === 'normal' ? '正常出勤' : '遲到/異常';
    var lateMinutes = data.lateMinutes || 0;
    var note = data.note || "";
    var syncTime = new Date().toLocaleString("zh-TW", { timeZone: "Asia/Taipei" });

    var lastRow = sheet.getLastRow();
    var foundRow = -1;

    if (lastRow > 1) {
      var values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      for (var i = 0; i < values.length; i++) {
        if (values[i][1] === empId && values[i][4] === date) {
          foundRow = i + 2;
          break;
        }
      }
    }

    if (foundRow > -1) {
      sheet.getRange(foundRow, 6).setValue(checkIn);
      sheet.getRange(foundRow, 7).setValue(checkOut);
      sheet.getRange(foundRow, 8).setValue(hours);
      sheet.getRange(foundRow, 9).setValue(status);
      sheet.getRange(foundRow, 10).setValue(lateMinutes);
      sheet.getRange(foundRow, 11).setValue(note);
      sheet.getRange(foundRow, 12).setValue(syncTime);
    } else {
      sheet.appendRow([
        logId, empId, empName, dept, date, checkIn, checkOut, hours, status, lateMinutes, note, syncTime
      ]);
    }

    updateMonthlyReportSheet(ss, empId, empName, dept, date, hours, status, lateMinutes);
    updatePersonalMonthlySheet(ss, empId, empName, dept, date, checkIn, checkOut, hours, status, lateMinutes, note, syncTime);

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "打卡資料已同步寫入 Google 試算表！",
      syncTime: syncTime
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 輔助函式：清理日期為 YYYY-MM-DD
function cleanDateStr(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Taipei", "yyyy-MM-dd");
  }
  var str = String(val);
  var match = str.match(/(\d{4}-\d{2}-\d{2})/);
  if (match) return match[1];
  try {
    var d = new Date(str);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, "Asia/Taipei", "yyyy-MM-dd");
    }
  } catch (e) {}
  return str.split('T')[0];
}

// 輔助函式：清理時間為 HH:mm:ss (去除 1899 年與時區)
function cleanTimeStr(val) {
  if (!val || val === '--:--:--') return '--:--:--';
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Taipei", "HH:mm:ss");
  }
  var str = String(val);
  var timeMatch = str.match(/(\d{2}:\d{2}:\d{2})/);
  if (timeMatch) return timeMatch[1];
  
  var shortTimeMatch = str.match(/(\d{2}:\d{2})/);
  if (shortTimeMatch) return shortTimeMatch[1] + ":00";

  try {
    var d = new Date(str);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, "Asia/Taipei", "HH:mm:ss");
    }
  } catch (e) {}

  return str;
}

// 輔助函式 1：自動更新「全公司每月出勤」
function updateMonthlyReportSheet(ss, empId, empName, dept, dateStr, hours, status, lateMinutes) {
  var monthKey = dateStr.substring(0, 7);
  var sheetName = "全公司每月出勤_" + monthKey;
  var mSheet = ss.getSheetByName(sheetName);

  if (!mSheet) {
    mSheet = ss.insertSheet(sheetName);
    mSheet.appendRow(["員工編號", "員工姓名", "部門", "月份", "正常出勤天數", "遲到天數/次數", "當月總遲到(分鐘)", "當月累計工時(小時)"]);
    var headerRange = mSheet.getRange(1, 1, 1, 8);
    headerRange.setBackground("#2563eb");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");
    mSheet.setFrozenRows(1);
  }

  var lastRow = mSheet.getLastRow();
  var foundRow = -1;
  if (lastRow > 1) {
    var ids = mSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (ids[i][0] === empId) {
        foundRow = i + 2;
        break;
      }
    }
  }

  if (foundRow > -1) {
    var curNormal = Number(mSheet.getRange(foundRow, 5).getValue()) || 0;
    var curLate = Number(mSheet.getRange(foundRow, 6).getValue()) || 0;
    var curLateMins = Number(mSheet.getRange(foundRow, 7).getValue()) || 0;
    var curHours = Number(mSheet.getRange(foundRow, 8).getValue()) || 0;

    if (status === '正常出勤') curNormal += 1;
    if (status === '遲到/異常') {
      curLate += 1;
      curLateMins += Number(lateMinutes) || 0;
    }
    curHours += Number(hours) || 0;

    mSheet.getRange(foundRow, 5).setValue(curNormal);
    mSheet.getRange(foundRow, 6).setValue(curLate);
    mSheet.getRange(foundRow, 7).setValue(curLateMins);
    mSheet.getRange(foundRow, 8).setValue(curHours.toFixed(1));
  } else {
    mSheet.appendRow([
      empId, empName, dept, monthKey, 
      status === '正常出勤' ? 1 : 0, 
      status === '遲到/異常' ? 1 : 0, 
      status === '遲到/異常' ? (Number(lateMinutes) || 0) : 0,
      hours
    ]);
  }
}

// 輔助函式 2：自動更新「個人單月出勤明細」
function updatePersonalMonthlySheet(ss, empId, empName, dept, dateStr, checkIn, checkOut, hours, status, lateMinutes, note, syncTime) {
  var monthKey = dateStr.substring(0, 7);
  var sheetName = "個人單月出勤明細";
  var pSheet = ss.getSheetByName(sheetName);

  if (!pSheet) {
    pSheet = ss.insertSheet(sheetName);
    pSheet.appendRow([
      "員工編號", "員工姓名", "部門", "考勤月份", 
      "最新打卡日期", "上班時間", "下班時間", "當日工時(小時)", 
      "當日出勤評定", "遲到(分鐘)", "打卡備註", "同步時間"
    ]);
    var headerRange = pSheet.getRange(1, 1, 1, 12);
    headerRange.setBackground("#4c1d95");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");
    pSheet.setFrozenRows(1);
  }

  pSheet.appendRow([
    empId, empName, dept, monthKey, 
    dateStr, checkIn, checkOut, hours, 
    status, lateMinutes, note, syncTime
  ]);
}

// GET 請求：雙向讀取「員工名冊」與格式化後的「考勤紀錄」
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. 讀取員工名冊
    var empSheet = getOrInitEmpSheet(ss);
    var empLastRow = empSheet.getLastRow();
    var employees = [];

    if (empLastRow > 1) {
      var data = empSheet.getRange(2, 1, empLastRow - 1, 12).getValues();
      for (var i = 0; i < data.length; i++) {
        var row = data[i];
        if (row[0]) {
          employees.push({
            id: String(row[0]),
            username: String(row[1]),
            pass: String(row[2]),
            name: String(row[3]),
            role: String(row[4]),
            dept: String(row[5]),
            isAdmin: row[6] === "主管",
            avatar: row[7] || "👨‍💼",
            shift: row[8] || "09:00 - 18:00",
            annualLeave: Number(row[9]) || 10,
            sickLeave: Number(row[10]) || 30,
            compLeave: Number(row[11]) || 0
          });
        }
      }
    }

    // 2. 讀取考勤紀錄 (格式化 Date 與 Time)
    var logSheet = ss.getSheetByName("考勤紀錄");
    var logs = [];

    if (logSheet) {
      var logLastRow = logSheet.getLastRow();
      if (logLastRow > 1) {
        var logData = logSheet.getRange(2, 1, logLastRow - 1, 12).getValues();
        for (var j = 0; j < logData.length; j++) {
          var lRow = logData[j];
          if (lRow[1] && lRow[4]) {
            var dateStr = cleanDateStr(lRow[4]);
            var checkInStr = cleanTimeStr(lRow[5]);
            var checkOutStr = cleanTimeStr(lRow[6]);
            var statusStr = String(lRow[8] || '');
            var isNormal = statusStr === '正常出勤' || statusStr === 'normal';

            logs.push({
              id: String(lRow[0] || ('LOG' + j)),
              empId: String(lRow[1]),
              empName: String(lRow[2] || ''),
              dept: String(lRow[3] || ''),
              date: dateStr,
              checkIn: checkInStr,
              checkOut: checkOutStr,
              hours: Number(lRow[7]) || 0,
              status: isNormal ? 'normal' : 'late',
              lateMinutes: Number(lRow[9]) || 0,
              note: String(lRow[10] || '')
            });
          }
        }
      }
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      employees: employees,
      logs: logs
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString(),
      employees: [],
      logs: []
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrInitEmpSheet(ss) {
  var empSheet = ss.getSheetByName("員工名冊與假額度");
  if (!empSheet) {
    empSheet = ss.insertSheet("員工名冊與假額度");
    empSheet.appendRow([
      "員工編號", "登入帳號", "登入密碼", "員工姓名", 
      "職稱/職位", "部門", "角色權限(主管/員工)", "頭像符號", 
      "工作班次", "特休天數", "病假天數", "補休小時"
    ]);
    var headerRange = empSheet.getRange(1, 1, 1, 12);
    headerRange.setBackground("#047857");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");
    empSheet.setFrozenRows(1);

    empSheet.appendRow(["EMP001", "emp01", "123456", "張小明", "前端開發工程師", "技術研發部", "員工", "👨‍💻", "09:00 - 18:00", 10, 30, 0]);
    empSheet.appendRow(["EMP002", "admin", "admin888", "李梅", "HR 人資主管", "人力資源部", "主管", "👩‍💼", "09:00 - 18:00", 14, 30, 8]);
    empSheet.appendRow(["EMP003", "emp03", "123456", "王大同", "資深產品經理", "產品設計部", "員工", "👨‍💼", "09:30 - 18:30", 12, 30, 12.5]);
  }
  return empSheet;
}
