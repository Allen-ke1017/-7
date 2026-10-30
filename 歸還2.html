function doGet(e) {
  // 取得前端傳來的 id 參數
  const id = e.parameter.id;
  
  if (!id) {
    return ContentService.createTextOutput("未接收到設備編號！");
  }

  // 開啟綁定的 Google 試算表與指定的 工作表 (例如: "設備清單")
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("設備清單");
  const data = sheet.getDataRange().getValues(); // 取得所有資料
  
  // 假設試算表結構如下：
  // 欄 A (index 0): 設備編號 (id)
  // 欄 B (index 1): 設備名稱
  // 欄 C (index 2): 目前狀態 ("可借用" 或 "已借出")
  // 欄 D (index 3): 最後異動時間
  
  let foundRow = -1;
  let currentStatus = "";
  let equipmentName = "";

  // 尋找對應 ID 的資料列（從第 2 列開始，避開標題列）
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() == id) {
      foundRow = i + 1; // 試算表列數從 1 開始
      equipmentName = data[i][1];
      currentStatus = data[i][2];
      break;
    }
  }

  // 若找不到該設備編號
  if (foundRow === -1) {
    return ContentService.createTextOutput("❌ 找不到此設備編號：" + id);
  }

  const now = new Date();
  const timeString = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd HH:mm:ss");

  // 核心邏輯：自動切換狀態
  if (currentStatus === "已借出") {
    // 進行歸還流程
    sheet.getRange(foundRow, 3).setValue("可借用"); // 更新狀態欄位為 "可借用"
    sheet.getRange(foundRow, 4).setValue(timeString); // 更新時間欄位
    
    return ContentService.createTextOutput(
      `成功歸還！\n設備：${equipmentName}\n目前狀態：可借用\n時間：${timeString}`
    );
    
  } else {
    // 進行借用流程
    sheet.getRange(foundRow, 3).setValue("已借出"); // 更新狀態欄位為 "已借出"
    sheet.getRange(foundRow, 4).setValue(timeString); // 更新時間欄位
    
    return ContentService.createTextOutput(
      `借用成功！\n設備：${equipmentName}\n目前狀態：已借出\n時間：${timeString}`
    );
  }
}
