/**
 * NIIMBOT 商城後端 v7.0 - 穩定版
 */

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("產品清單");
  if (!sheet) return makeJson([]);

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return makeJson([]);

  const headers = data[0];
  const rows = data.slice(1);
  
  // 建立欄位索引字典
  const idx = {
    id: headers.indexOf("編號"),
    name: headers.indexOf("品項名稱"),
    mainCat: headers.indexOf("大分類"),
    subCat1: headers.indexOf("二分類"),
    subCat2: headers.indexOf("三分類"),
    color: headers.indexOf("顏色"),
    price: headers.indexOf("特價"),
    originalPrice: headers.indexOf("原價"),
    stock: headers.indexOf("庫存"),
    img: headers.indexOf("圖片連結")
  };

  // 轉換資料，過濾掉品項名稱或編號為空的列
  const result = rows
    .filter(r => r[idx.id] && r[idx.name])
    .map(r => ({
      id: String(r[idx.id]),
      name: String(r[idx.name]),
      mainCat: String(r[idx.mainCat]),
      subCat1: String(r[idx.subCat1]),
      subCat2: String(r[idx.subCat2]),
      color: String(r[idx.color] || "預設"),
      price: Number(r[idx.price]) || 0,
      originalPrice: Number(r[idx.originalPrice]) || 0,
      stock: Number(r[idx.stock]) || 0,
      img: String(r[idx.img] || ""),
      type: r[idx.mainCat] === "標籤機專區" ? "printer" : "sticker"
    }));

  return makeJson(result);
}

function makeJson(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = JSON.parse(e.postData.contents);
  const sheet = ss.getSheetByName("全訂單紀錄總表") || ss.insertSheet("全訂單紀錄總表");
  
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["訂單編號", "時間", "姓名", "電話", "Email", "產品", "規格", "數量", "小計", "總計", "支付", "備註"]);
  }

  data.itemsList.forEach(item => {
    sheet.appendRow([
      data.orderId, data.timestamp, data.name, data.phone, data.email, 
      item.name, item.color, item.quantity, item.subtotal, data.totalAmount, data.payment, data.address
    ]);
  });
  return ContentService.createTextOutput("Success");
}

function setupDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("產品清單") || ss.insertSheet("產品清單");
  
  // 初始化標題
  const headers = [["編號", "品項名稱", "大分類", "二分類", "三分類", "顏色", "特價", "原價", "庫存", "圖片連結"]];
  sheet.getRange(1, 1, 1, 10).setValues(headers).setFontWeight("bold").setBackground("#f8fafc");

  // 大分類
  sheet.getRange("C2:C2000").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["標籤機專區", "貼紙專區"]).build());
  // 二分類
  const s1 = ["D系列手持小機器", "B系列商用家用大機器", "N1/M2/M3熱轉印標籤機", "K2/K3大量列印商用機", "D系列機器通用貼紙", "B系列機器通用貼紙", "Pro機器專用貼紙", "B31/B3S_P專用大貼紙", "B4專用大貼紙", "N1專用碳帶及貼紙", "M2/M3專用碳帶及貼紙", "K2/K3專用貼紙"];
  sheet.getRange("D2:D2000").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(s1).build());
  // 三分類
  const s2 = ["D110標籤機", "D110M標籤機", "D11S標籤機", "D11S三麗鷗聯名款", "D101標籤機", "H1S標籤機", "B1標籤機", "B1_PRO標籤機", "B21S標籤機", "B21PRO標籤機", "B3S_P標籤機", "B31標籤機", "B4標籤機", "N1標籤機", "M2標籤機", "M3標籤機", "K2標籤機", "K3標籤機", "全部通用"];
  sheet.getRange("E2:E2000").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(s2).build());
  
  SpreadsheetApp.getUi().alert("✅ 試算表結構設定完成！");
}