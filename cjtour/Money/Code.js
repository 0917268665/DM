/**
 * 旅行社專屬帳務系統 - Web App API 後端 (Code.js)
 * 
 * 本程式碼將 Google Apps Script 部署為 API 服務端點，供 Cloudflare 獨立前端網頁進行呼叫。
 * 1. doPost(e) - 統一入口，解析來自 Cloudflare 的 JSON 請求並執行對應的後端邏輯。
 * 2. 系統初始化 (setupSheets) - 一鍵自動建立工作表與範例資料。
 * 3. 訂單管理 (createOrder) - 新增訂單並自動計算毛利與純利。
 * 4. 網頁版互動日曆數據 (getCalendarData) - 提供指定月份的出團資料。
 * 5. PDF 報表生成引擎 (generateSalesReportPDF / generateFinancialReportPDF) - 產出 PDF 並傳回雲端下載連結。
 * 
 * 開發專家：Antigravity
 */

// 系統全域變數定義
const SHEET_NAMES = {
  SALES: "業務清單",
  ITEMS: "項目清單",
  ORDERS: "訂單",
  TRANSACTIONS: "收支紀錄"
};

/**
 * 當試算表開啟時，自動執行此函式以建立自訂選單。
 * 即使網頁部署在 Cloudflare，管理者仍可在試算表中點選此選單進行初始化。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("💼 旅行社帳務系統")
    .addItem("初始化系統工作表", "setupSheets")
    .addToUi();
}

/**
 * 核心 API 進入點 (POST 請求)
 * 接收 Cloudflare 前端發送的 JSON 請求，並分流至各個功能函式
 */
function doPost(e) {
  try {
    // 解析前端傳來的 JSON 字串 (前端會以 text/plain 發送以避免 CORS preflight 限制)
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;
    let result = {};

    if (action === "getSales") {
      result = { success: true, data: getSalesList() };
    } 
    else if (action === "getCalendar") {
      result = { success: true, data: getCalendarData(postData.year, postData.month) };
    } 
    else if (action === "createOrder") {
      result = createOrder(postData.orderData);
    } 
    else if (action === "genSalesReport") {
      result = generateSalesReportPDF(postData.salesName, postData.year, postData.month);
    } 
    else if (action === "genFinReport") {
      result = generateFinancialReportPDF(postData.year);
    } 
    else {
      throw new Error(`未定義的 API Action: ${action}`);
    }

    // 回傳 JSON 物件，Google Web App 會自動處理 CORS 重定向
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log("API 執行失敗: " + err.message);
    const errResult = { success: false, error: err.message };
    return ContentService.createTextOutput(JSON.stringify(errResult))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 取得所有在職業務姓名清單
 */
function getSalesList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SALES);
  if (!sheet) throw new Error("找不到業務清單工作表。");
  
  const data = sheet.getDataRange().getValues();
  const salesList = [];
  
  // 標頭：業務編號, 業務姓名, 聯絡電話, 電子郵件, 在職狀態
  for (let i = 1; i < data.length; i++) {
    const name = data[i][1];
    const status = data[i][4];
    if (name && status === "在職") {
      salesList.push(name);
    }
  }
  return salesList;
}

/**
 * 取得指定年月出團資料
 */
function getCalendarData(year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
  if (!orderSheet) return [];

  const orders = orderSheet.getDataRange().getValues();
  const calendarData = [];

  const targetYear = parseInt(year, 10);
  const targetMonth = parseInt(month, 10) - 1; // 0-11

  // 出發日期: index 2, 產品項目: index 7, 人數: index 8, 應收: index 9, 訂單狀態: index 18
  for (let i = 1; i < orders.length; i++) {
    const depDateVal = orders[i][2];
    if (!depDateVal) continue;
    
    const depDate = new Date(depDateVal);
    const status = orders[i][18];
    
    if (status !== "已取消" && depDate.getFullYear() === targetYear && depDate.getMonth() === targetMonth) {
      calendarData.push({
        orderNo: orders[i][0],
        date: Utilities.formatDate(depDate, "GMT+8", "yyyy-MM-dd"),
        sales: orders[i][3],
        customer: orders[i][5],
        item: orders[i][7],
        pax: parseInt(orders[i][8], 10) || 0,
        revenue: parseFloat(orders[i][9]) || 0
      });
    }
  }
  return calendarData;
}

/**
 * 新增一筆訂單
 */
function createOrder(orderData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
    if (!sheet) throw new Error("找不到訂單工作表。");

    // 生成訂單編號 (ORD-YYYYMMDD-三位流水號)
    const todayStr = Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd");
    const lastRow = sheet.getLastRow();
    let serial = 1;
    if (lastRow >= 2) {
      const lastOrderNo = sheet.getRange(lastRow, 1).getValue();
      if (lastOrderNo.startsWith("ORD-" + todayStr)) {
        const parts = lastOrderNo.split("-");
        serial = parseInt(parts[2], 10) + 1;
      }
    }
    const orderNo = `ORD-${todayStr}-${String(serial).padStart(3, "0")}`;

    const newRow = [
      orderNo,
      orderData.orderDate || Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd"),
      orderData.departureDate,
      orderData.salesName,
      orderData.guideName || "無",
      orderData.customerName,
      orderData.customerPhone,
      orderData.itemName,
      parseInt(orderData.pax, 10) || 0,
      parseFloat(orderData.revenue) || 0,
      parseFloat(orderData.cost) || 0,
      parseFloat(orderData.guidePaid) || 0,
      parseFloat(orderData.commission) || 0,
      parseFloat(orderData.taxRefund) || 0,
      "", // 毛利公式
      "", // 淨毛利公式
      0,  // 實收
      0,  // 實付
      orderData.status || "已報名"
    ];

    sheet.appendRow(newRow);
    
    // 套用公式
    const newRowIdx = sheet.getLastRow();
    sheet.getRange(newRowIdx, 15).setFormula(`=J${newRowIdx}-K${newRowIdx}-L${newRowIdx}+N${newRowIdx}`);
    sheet.getRange(newRowIdx, 16).setFormula(`=O${newRowIdx}-M${newRowIdx}`);
    sheet.getRange(newRowIdx, 10, 1, 9).setNumberFormat("$#,##0");
    
    return { success: true, orderNo: orderNo };
  } catch (e) {
    Logger.log("新增訂單失敗: " + e.message);
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 產生 PDF 業務業績報表
 */
function generateSalesReportPDF(salesName, year, month) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
    if (!orderSheet) throw new Error("找不到訂單工作表。");

    const orders = orderSheet.getDataRange().getValues();
    const filteredOrders = [];
    
    let totalRevenue = 0, totalCost = 0, totalGuidePaid = 0, totalCommission = 0, totalProfit = 0, totalNetProfit = 0, totalPax = 0, orderCount = 0;

    const targetYear = parseInt(year, 10);
    const targetMonth = month ? parseInt(month, 10) : null;

    for (let i = 1; i < orders.length; i++) {
      const orderDate = new Date(orders[i][1]); // 依訂購日歸屬業績
      const oSales = orders[i][3];
      const status = orders[i][18];

      if (oSales === salesName && status !== "已取消") {
        const orderYear = orderDate.getFullYear();
        const orderMonth = orderDate.getMonth() + 1;

        if (orderYear === targetYear && (!targetMonth || orderMonth === targetMonth)) {
          filteredOrders.push(orders[i]);
          
          totalRevenue += parseFloat(orders[i][9]) || 0;
          totalCost += parseFloat(orders[i][10]) || 0;
          totalGuidePaid += parseFloat(orders[i][11]) || 0;
          totalCommission += parseFloat(orders[i][12]) || 0;
          
          const profit = (parseFloat(orders[i][9]) || 0) - (parseFloat(orders[i][10]) || 0) - (parseFloat(orders[i][11]) || 0) + (parseFloat(orders[i][13]) || 0);
          totalProfit += profit;
          totalNetProfit += (profit - (parseFloat(orders[i][12]) || 0));
          totalPax += parseInt(orders[i][8], 10) || 0;
          orderCount++;
        }
      }
    }

    const timeLabel = month ? `${year}年${String(month).padStart(2, '0')}月` : `${year}年度`;
    const docName = `業務績效報表_${salesName}_${timeLabel}`;

    const doc = DocumentApp.create(docName);
    const body = doc.getBody();
    body.setMarginTop(36).setMarginBottom(36).setMarginLeft(36).setMarginRight(36);

    const title = body.appendParagraph(`💼 業務個人績效報表 - ${salesName}`);
    title.setFontSize(22).setFontColor("#1e293b").setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    const subtitle = body.appendParagraph(`報表區間：${timeLabel}   |   產生時間：${Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm")}`);
    subtitle.setFontSize(10).setFontColor("#64748b").setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(24);

    const kpiTableCells = [
      ["總訂單數", "營業總額 (TWD)", "服務總人數"],
      [`${orderCount} 筆`, `$${formatNumber(totalRevenue)}`, `${totalPax} 人`],
      ["領隊代墊總計", "業務佣金總計", "業務淨毛利總計"],
      [`$${formatNumber(totalGuidePaid)}`, `$${formatNumber(totalCommission)}`, `$${formatNumber(totalNetProfit)}`]
    ];
    const kpiTable = body.appendTable(kpiTableCells);
    formatReportTable(kpiTable, true);

    body.appendParagraph("").setSpacingBefore(12);

    const detailTitle = body.appendParagraph("📄 訂單交易明細");
    detailTitle.setFontSize(14).setFontColor("#1e293b").setBold(true).setSpacingAfter(10);

    const detailHeaders = ["訂單編號", "出發日期", "客戶名稱", "人數", "營業額", "代墊/佣金", "訂單毛利"];
    const detailRows = [detailHeaders];

    filteredOrders.forEach(ord => {
      detailRows.push([
        ord[0],
        Utilities.formatDate(new Date(ord[2]), "GMT+8", "yyyy-MM-dd"),
        ord[5],
        `${ord[8]}人`,
        `$${formatNumber(ord[9])}`,
        `$${formatNumber(ord[11])} / $${formatNumber(ord[12])}`,
        `$${formatNumber(ord[14])}`
      ]);
    });

    if (filteredOrders.length === 0) {
      detailRows.push(["-", "-", "此期間尚無相關出團訂單", "-", "-", "-", "-"]);
    }

    const detailTable = body.appendTable(detailRows);
    formatReportTable(detailTable, false);

    doc.saveAndClose();
    const pdfBlob = doc.getAs('application/pdf');

    let folder;
    const folders = DriveApp.getFoldersByName("旅行社帳務報表");
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder("旅行社帳務報表");
    }

    const pdfFile = folder.createFile(pdfBlob);
    pdfFile.setName(`${docName}.pdf`);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    DriveApp.getFileById(doc.getId()).setTrashed(true);

    return { success: true, pdfUrl: pdfFile.getUrl(), pdfName: pdfFile.getName() };
  } catch (e) {
    Logger.log("產生 PDF 業務報表失敗: " + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 產生 PDF 年度財務年報
 */
function generateFinancialReportPDF(year) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
    const txnSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
    
    if (!orderSheet || !txnSheet) throw new Error("找不到訂單或收支工作表。");

    const monthlyStats = Array.from({ length: 12 }, (_, i) => ({
      month: i + 1,
      revenue: 0, cost: 0, guidePaid: 0, commission: 0, taxRefund: 0, adminExpense: 0, received: 0, paid: 0
    }));

    // 1. 統計訂單
    const orders = orderSheet.getDataRange().getValues();
    for (let i = 1; i < orders.length; i++) {
      const orderDate = new Date(orders[i][1]);
      const status = orders[i][18];
      if (orderDate.getFullYear() === parseInt(year, 10) && status !== "已取消") {
        const mIdx = orderDate.getMonth();
        monthlyStats[mIdx].revenue += parseFloat(orders[i][9]) || 0;
        monthlyStats[mIdx].cost += parseFloat(orders[i][10]) || 0;
        monthlyStats[mIdx].guidePaid += parseFloat(orders[i][11]) || 0;
        monthlyStats[mIdx].commission += parseFloat(orders[i][12]) || 0;
        monthlyStats[mIdx].taxRefund += parseFloat(orders[i][13]) || 0;
      }
    }

    // 2. 統計實際收支與行政雜支
    const txns = txnSheet.getDataRange().getValues();
    for (let i = 1; i < txns.length; i++) {
      const txnDate = new Date(txns[i][1]);
      if (txnDate.getFullYear() === parseInt(year, 10)) {
        const mIdx = txnDate.getMonth();
        const type = txns[i][3];
        const category = txns[i][4];
        const orderNo = txns[i][2];
        const amount = parseFloat(txns[i][5]) || 0;

        if (type === "收入") {
          monthlyStats[mIdx].received += amount;
        } else if (type === "支出") {
          if (!orderNo || category === "行政雜支") {
            monthlyStats[mIdx].adminExpense += amount;
          } else {
            monthlyStats[mIdx].paid += amount;
          }
        }
      }
    }

    const docName = `公司年度財務報表_${year}年`;
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();
    body.setMarginTop(36).setMarginBottom(36).setMarginLeft(36).setMarginRight(36);

    const title = body.appendParagraph(`📊 公司年度財務收支報表 (${year}年)`);
    title.setFontSize(22).setFontColor("#0f172a").setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    const subtitle = body.appendParagraph(`產生時間：${Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm")}`);
    subtitle.setFontSize(10).setFontColor("#64748b").setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(20);

    const headers = ["月份", "營業總額", "業務成本", "領隊代墊", "業務佣金", "行政雜支", "公司純利"];
    const rows = [headers];

    let grandRevenue = 0, grandCost = 0, grandGuide = 0, grandComm = 0, grandAdmin = 0, grandNet = 0;

    monthlyStats.forEach(stat => {
      const netProfit = stat.revenue - stat.cost - stat.guidePaid - stat.commission - stat.adminExpense + stat.taxRefund;
      grandRevenue += stat.revenue;
      grandCost += stat.cost;
      grandGuide += stat.guidePaid;
      grandComm += stat.commission;
      grandAdmin += stat.adminExpense;
      grandNet += netProfit;

      rows.push([
        `${stat.month}月`,
        `$${formatNumber(stat.revenue)}`,
        `$${formatNumber(stat.cost)}`,
        `$${formatNumber(stat.guidePaid)}`,
        `$${formatNumber(stat.commission)}`,
        `$${formatNumber(stat.adminExpense)}`,
        `$${formatNumber(netProfit)}`
      ]);
    });

    rows.push([
      "總計",
      `$${formatNumber(grandRevenue)}`,
      `$${formatNumber(grandCost)}`,
      `$${formatNumber(grandGuide)}`,
      `$${formatNumber(grandComm)}`,
      `$${formatNumber(grandAdmin)}`,
      `$${formatNumber(grandNet)}`
    ]);

    const table = body.appendTable(rows);
    formatReportTable(table, false);

    const totalRowIdx = table.getNumberOfRows() - 1;
    const totalRow = table.getRow(totalRowIdx);
    for (let c = 0; c < totalRow.getNumCells(); c++) {
      totalRow.getCell(c).setBold(true).setBackgroundColor("#f1f5f9");
    }

    doc.saveAndClose();
    const pdfBlob = doc.getAs('application/pdf');

    let folder;
    const folders = DriveApp.getFoldersByName("旅行社帳務報表");
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder("旅行社帳務報表");
    }

    const pdfFile = folder.createFile(pdfBlob);
    pdfFile.setName(`${docName}.pdf`);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    DriveApp.getFileById(doc.getId()).setTrashed(true);

    return { success: true, pdfUrl: pdfFile.getUrl(), pdfName: pdfFile.getName() };
  } catch (e) {
    Logger.log("產生 PDF 財務年報失敗: " + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 輔助函數：初始化系統工作表
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const sheetDefinitions = [
    {
      name: SHEET_NAMES.SALES,
      headers: ["業務編號", "業務姓名", "聯絡電話", "電子郵件", "在職狀態"],
      initData: [
        ["S001", "陳小明", "0912-345678", "xiaoming@example.com", "在職"],
        ["S002", "林美玲", "0923-456789", "meiling@example.com", "在職"],
        ["S003", "張大華", "0934-567890", "dahua@example.com", "離職"]
      ]
    },
    {
      name: SHEET_NAMES.ITEMS,
      headers: ["項目編號", "項目名稱", "項目大類", "參考底價", "參考售價"],
      initData: [
        ["I001", "日本東京五日遊", "團體遊", 25000, 29900],
        ["I002", "韓國首爾四日遊", "團體遊", 18000, 21900],
        ["I003", "泰國曼谷自由行", "自由行", 12000, 14500],
        ["I004", "全球訂房服務", "代訂服務", 0, 0]
      ]
    },
    {
      name: SHEET_NAMES.ORDERS,
      headers: [
        "訂單編號", "訂購日期", "出發日期", "業務姓名", "領隊姓名", 
        "客戶名稱", "客戶電話", "產品項目", "人數 (人頭)", "應收總額 (營業額)", 
        "應付成本", "領隊代墊款", "業務佣金", "退稅金額", "訂單毛利", "業務淨毛利", 
        "實收金額", "實付金額", "訂單狀態"
      ],
      initData: [
        ["ORD-20260601-001", "2026-06-01", "2026-06-15", "陳小明", "張大山", "王大同", "0955-111222", "日本東京五日遊", 4, 119600, 100000, 2000, 5000, 1500, "", "", 119600, 102000, "已成行"],
        ["ORD-20260601-002", "2026-06-01", "2026-06-20", "林美玲", "無", "李莉莉", "0966-333444", "泰國曼谷自由行", 2, 29000, 24000, 0, 1500, 0, "", "", 10000, 0, "已報名"]
      ]
    },
    {
      name: SHEET_NAMES.TRANSACTIONS,
      headers: [
        "交易編號", "交易日期", "關聯訂單號", "收支類型", "交易大類", 
        "金額", "支付方式", "交易對象", "經辦人", "備註說明"
      ],
      initData: [
        ["TXN-20260601-001", "2026-06-01", "ORD-20260601-001", "收入", "團費收入", 119600, "匯款", "王大同", "會計阿花", "全額付清"],
        ["TXN-20260601-002", "2026-06-01", "ORD-20260601-001", "支出", "代收轉付", 100000, "匯款", "供應商A", "會計阿花", "供應商團費成本"],
        ["TXN-20260601-003", "2026-06-01", "ORD-20260601-001", "支出", "領隊代墊報銷", 2000, "現金", "張大山", "會計阿花", "報銷門票雜支"],
        ["TXN-20260601-004", "2026-06-01", "ORD-20260601-002", "收入", "團費收入", 10000, "刷卡", "李莉莉", "會計阿花", "訂金"]
      ]
    }
  ];

  sheetDefinitions.forEach(def => {
    let sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
    } else {
      sheet.clear();
    }
    
    if (def.headers.length > 0) {
      sheet.appendRow(def.headers);
      const headerRange = sheet.getRange(1, 1, 1, def.headers.length);
      headerRange.setBackground("#1e293b")
                 .setFontColor("#ffffff")
                 .setFontWeight("bold")
                 .setHorizontalAlignment("center");
    }
    
    if (def.initData.length > 0) {
      sheet.getRange(2, 1, def.initData.length, def.initData[0].length).setValues(def.initData);
    }
    
    if (def.headers.length > 0) {
      sheet.autoResizeColumns(1, def.headers.length);
    }
  });

  setupOrderSheetFormulas();
  ui.alert("系統初始化成功！", "所有包含領隊代墊、業務佣金與退稅的資料表結構已建立完成。", ui.ButtonSet.OK);
}

function setupOrderSheetFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  for (let r = 2; r <= lastRow; r++) {
    sheet.getRange(r, 15).setFormula(`=J${r}-K${r}-L${r}+N${r}`);
    sheet.getRange(r, 16).setFormula(`=O${r}-M${r}`);
  }
  sheet.getRange(2, 10, lastRow - 1, 9).setNumberFormat("$#,##0");
}

function formatNumber(num) {
  if (isNaN(num)) return "0";
  return Math.round(num).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function formatReportTable(table, isKpi) {
  const headerBgColor = isKpi ? "#f8fafc" : "#1e293b";
  const headerTextColor = isKpi ? "#334155" : "#ffffff";
  const rowCount = table.getNumberOfRows();

  for (let r = 0; r < rowCount; r++) {
    const row = table.getRow(r);
    const colCount = row.getNumCells();
    const cellMargin = isKpi ? 8 : 5;

    for (let c = 0; c < colCount; c++) {
      const cell = row.getCell(c);
      const p = cell.getChild(0).asParagraph();
      p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      
      cell.setPaddingTop(cellMargin).setPaddingBottom(cellMargin).setPaddingLeft(6).setPaddingRight(6);

      if (isKpi) {
        if (r % 2 === 0) {
          cell.setBackgroundColor(headerBgColor);
          p.setFontSize(10).setFontColor(headerTextColor).setBold(true);
        } else {
          cell.setBackgroundColor("#ffffff");
          p.setFontSize(14).setFontColor("#6366f1").setBold(true);
        }
      } else {
        if (r === 0) {
          cell.setBackgroundColor(headerBgColor);
          p.setFontSize(10).setFontColor(headerTextColor).setBold(true);
        } else {
          cell.setBackgroundColor(r % 2 === 0 ? "#f8fafc" : "#ffffff");
          p.setFontSize(9).setFontColor("#334155");
        }
      }
      cell.setBorderColor("#cbd5e1");
    }
  }
}
