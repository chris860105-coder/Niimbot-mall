/**
 * 精臣商城後端系統 v23.0 - 智慧模糊匹配與欄位分離防護版
 * 適用於：全訂單紀錄總表 (14 欄位 A~N 完美對應)
 */

function doGet(e) {
  // 安全性防護：若在編輯器直接執行 doGet()，攔截 parameter 錯誤並給予友善提示
  if (!e || !e.parameter) {
    return ContentService.createTextOutput("🎉 後端服務運作正常中！\n請使用部署後的網址進行對接，不要直接在 Google Apps Script 編輯器內點擊執行 doGet 按鈕。")
                         .setMimeType(ContentService.MimeType.TEXT);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const action = e.parameter.action;
  
  // 處理「整筆取消並回補庫存」
  if (action === "cancelEntireOrder") {
    return handleCancelEntireOrder(e.parameter.orderId);
  }

  // 處理「查詢紀錄」
  if (e.parameter.phone || e.parameter.orderId) {
    return lookupOrder(e.parameter);
  }

  // 預設：產品清單
  return getProductList(ss);
}

function doPost(e) {
  try {
    // 安全性防護：避免在編輯器直接點執行 doPost 發生錯誤
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput("Error: 無法讀取 POST 資料內容，此接口僅供前端網頁傳送訂單。")
                           .setMimeType(ContentService.MimeType.TEXT);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = JSON.parse(e.postData.contents);
    const orderSheet = ss.getSheets()[0]; // 預設第一個工作表為訂單總表
    
    // 總表標題列：獨立分出 14 個欄位 (完美對應 A 到 N 欄)
    if (orderSheet.getLastRow() === 0 || orderSheet.getRange("A1").getValue() === "") {
      orderSheet.getRange("A1:N1").setValues([["訂單編號", "時間", "姓名", "電話", "Email", "產品", "規格", "數量", "小計", "總計", "支付/狀態", "地址", "備註", "取貨方式"]]);
    }

    data.itemsList.forEach(item => {
      // 確保品項名稱寫入時，前端的商品狀態（[現貨] 或 [預購]）能完美呈現在「產品類型 / 產品名稱」欄位
      const displayItemName = `[${item.status || '現貨'}] ${item.name}`;

      orderSheet.appendRow([
        data.orderId, 
        data.timestamp, 
        data.name, 
        "'" + data.phone,           // 加上單引號防 0 遺失
        data.email || "", 
        displayItemName,            // 寫入帶有現貨/預購標籤的品項名稱 (Col F)
        item.color,                 // 規格 (Col G)
        item.quantity,              // 數量 (Col H)
        (item.price * item.quantity), // 小計 (Col I)
        data.totalAmount,           // 總計 (Col J)
        data.payment,               // 支付方式 (Col K)
        data.address || "",         // 獨立地址 (L欄 / Col L)
        data.memo || "",            // 獨立備註 (M欄 / Col M)
        data.deliveryOption || ""   // 獨立出貨方式 (N欄 / Col N)
      ]);
      // 下單扣庫存 (傳入原始商品名稱，不要傳帶有狀態前綴的名稱)
      changeProductStock(item.name, item.color, -Math.abs(item.quantity));
    });
    
    return ContentService.createTextOutput("Success").setMimeType(ContentService.MimeType.TEXT);
  } catch (error) {
    return ContentService.createTextOutput("Error: " + error.toString()).setMimeType(ContentService.MimeType.TEXT);
  }
}

/**
 * 取消訂單邏輯：將狀態改為「已取消」，金額歸零，並把數量加回庫存
 */
function handleCancelEntireOrder(orderId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheets()[0]; 
  const orderData = orderSheet.getDataRange().getValues();
  const restoredItems = [];
  let found = false;
  
  for (let i = 1; i < orderData.length; i++) {
    if (String(orderData[i][0]).trim() === String(orderId).trim() && orderData[i][10] !== "已取消") {
      const rawItemName = String(orderData[i][5]);
      // 去除前綴 [現貨] 或 [預購]，還原成原始品項名稱以利回補庫存
      const cleanItemName = rawItemName.replace(/^\[.*?\]\s*/, '').trim(); 
      const itemColor = String(orderData[i][6]);
      const qty = Number(orderData[i][7]);
      const price = Number(orderData[i][8]) / (qty || 1);

      restoredItems.push({ name: cleanItemName, color: itemColor, quantity: qty, price: price });

      orderSheet.getRange(i + 1, 11).setValue("已取消");
      orderSheet.getRange(i + 1, 9).setValue(0);
      
      // 同步回補產品清單的庫存
      changeProductStock(cleanItemName, itemColor, qty);
      found = true;
    }
  }

  SpreadsheetApp.flush();

  if (found) {
    return makeJson({success: true, restoredItems: restoredItems});
  } else {
    return makeJson({success: false, message: "訂單不存在或已是取消狀態"});
  }
}

// 智慧模糊比對標題索引輔助函數
function findHeaderIdx(headers, keywords) {
  return headers.findIndex(h => keywords.some(k => String(h).trim().toLowerCase().includes(k.toLowerCase())));
}

function changeProductStock(itemName, itemColor, delta) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("產品清單");
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameIdx = findHeaderIdx(headers, ["品項名稱", "品項", "名稱", "商品名稱", "產品名稱"]);
  const colorIdx = findHeaderIdx(headers, ["顏色", "規格", "樣式"]);
  const stockIdx = findHeaderIdx(headers, ["庫存", "數量"]);

  if (nameIdx === -1 || stockIdx === -1) return;

  for (let i = 1; i < data.length; i++) {
    const rowName = String(data[i][nameIdx]).trim();
    const rowColor = String(data[i][colorIdx] || "預設").trim();
    
    if (rowName === String(itemName).trim() && rowColor === String(itemColor).trim()) {
      const currentStock = Number(data[i][stockIdx]) || 0;
      sheet.getRange(i + 1, stockIdx + 1).setValue(currentStock + delta);
      break;
    }
  }
}

function getProductList(ss) {
  const sheet = ss.getSheetByName("產品清單");
  if (!sheet) return makeJson([]);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // 使用智慧模糊比對，徹底防止因為標題多一兩個字而讀取失敗
  const idx = { 
    id: findHeaderIdx(headers, ["編號", "id", "商品編號"]), 
    name: findHeaderIdx(headers, ["品項名稱", "品項", "名稱", "商品名稱", "產品名稱"]), 
    main: findHeaderIdx(headers, ["大分類", "主分類"]), 
    sub1: findHeaderIdx(headers, ["二分類", "子分類"]), 
    status: findHeaderIdx(headers, ["狀態", "現貨", "預購", "商品狀態"]),     
    color: findHeaderIdx(headers, ["顏色", "規格", "樣式"]), 
    price: findHeaderIdx(headers, ["特價", "價格", "售價", "金額"]), 
    original: findHeaderIdx(headers, ["原價", "原價價格"]), 
    stock: findHeaderIdx(headers, ["庫存", "數量"]), 
    img: findHeaderIdx(headers, ["圖片連結", "圖片", "商品圖片"]) 
  };
  
  const products = data.slice(1).filter(r => idx.name > -1 && idx.id > -1 && r[idx.name] && r[idx.id]).map(r => ({
    id: String(r[idx.id]), 
    name: String(r[idx.name]), 
    mainCat: idx.main > -1 ? String(r[idx.main]) : "", 
    subCat1: idx.sub1 > -1 ? String(r[idx.sub1]) : "", 
    status: idx.status > -1 ? String(r[idx.status] || "現貨").trim() : "現貨", 
    color: idx.color > -1 ? String(r[idx.color] || "預設") : "預設",
    price: idx.price > -1 ? (Number(r[idx.price]) || 0) : 0, 
    originalPrice: idx.original > -1 ? (Number(r[idx.original]) || 0) : 0, 
    stock: idx.stock > -1 ? (Number(r[idx.stock]) || 0) : 0, 
    img: idx.img > -1 ? String(r[idx.img] || "") : "",
    type: (idx.main > -1 && String(r[idx.main]).indexOf("標籤機") > -1) ? "printer" : "sticker"
  }));
  return makeJson(products);
}

function lookupOrder(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("全訂單紀錄總表");
  if (!sheet) return makeJson([]);
  const data = sheet.getDataRange().getValues();
  const tempMap = new Map();
  
  const targetPhone = params.phone ? String(params.phone).replace(/^0+/, '').trim() : null;
  const targetId = params.orderId ? String(params.orderId).trim() : null;

  for (let i = data.length - 1; i >= 1; i--) {
    const oid = String(data[i][0]).trim();
    const rowPhone = String(data[i][3]).replace(/^0+/, '').trim();
    
    if ((targetPhone && rowPhone === targetPhone) || (targetId && oid === targetId)) {
      if (!tempMap.has(oid)) {
        // 地址對應 L 欄 (index 11)、備註對應 M 欄 (index 12)、出貨方式對應 N 欄 (index 13)
        tempMap.set(oid, { 
          orderId: oid, 
          timestamp: data[i][1], 
          name: data[i][2], 
          phone: String(data[i][3]), 
          email: data[i][4] || "", 
          itemsList: [], 
          totalAmount: 0, 
          payment: data[i][10], 
          address: data[i][11] || "",
          memo: data[i][12] || "",
          deliveryOption: data[i][13] || ""
        });
      }
      const order = tempMap.get(oid);
      if (data[i][10] !== "已取消") {
        const qty = Math.max(1, Number(data[i][7]));
        const rawName = String(data[i][5]);
        const cleanName = rawName.replace(/^\[.*?\]\s*/, '').trim(); 
        order.itemsList.push({ name: cleanName, color: data[i][6], quantity: qty, price: (Number(data[i][8])/qty) });
        order.totalAmount += Number(data[i][8]);
      }
    }
  }
  return makeJson(Array.from(tempMap.values()));
}

function makeJson(d) { return ContentService.createTextOutput(JSON.stringify(d)).setMimeType(ContentService.MimeType.JSON); }
