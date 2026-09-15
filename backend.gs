/**
 * 精臣商城後端系統 v28.0 - 效能、併發防護與安全性極致優化版 (STEP 1)
 * 部署網址：請替換為您的 Web App URL
 */

function doGet(e) {
  if (!e || !e.parameter) {
    return ContentService.createTextOutput("🎉 後端服務運作正常中！\n請使用部署後的網址進行對接，不要直接在編輯器內點擊執行 doGet。")
                         .setMimeType(ContentService.MimeType.TEXT);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const action = e.parameter.action;
  
  if (action === "cancelEntireOrder") {
    return handleCancelEntireOrder(e.parameter.orderId);
  }

  if (e.parameter.phone || e.parameter.orderId) {
    return lookupOrder(e.parameter);
  }

  return getProductList(ss);
}

function doPost(e) {
  // 1. 啟用 LockService 防止 Race Condition (等待時間上限 10 秒)
  const lock = LockService.getScriptLock();
  
  try {
    if (!lock.tryLock(10000)) {
      return makeJson({success: false, message: "系統繁忙，許多人正在結帳，請稍後再試！"});
    }

    if (!e || !e.postData || !e.postData.contents) {
      return makeJson({success: false, message: "無效的請求"});
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = JSON.parse(e.postData.contents);
    const orderSheet = ss.getSheets()[0]; 
    const sheetProduct = ss.getSheetByName("產品清單");

    if (!sheetProduct) {
      return makeJson({success: false, message: "找不到產品清單"});
    }

    if (!Array.isArray(data.itemsList) || data.itemsList.length === 0) {
      return makeJson({success: false, message: "商品數量格式錯誤"});
    }

    for (let item of data.itemsList) {
      const qty = Number(item.quantity);
      if (!Number.isFinite(qty) || qty <= 0 || !Number.isInteger(qty)) {
        return makeJson({success: false, message: "商品數量格式錯誤"});
      }
      item.quantity = qty; 
    }

    // ==========================================
    // 🛡️ 結帳前最後防線：嚴格檢查庫存是否足夠 (統一讀取一次)
    // ==========================================
    const prodData = sheetProduct.getDataRange().getValues();
    const headers = prodData[0];
    const nameIdx = findHeaderIdx(headers, ["品項名稱", "品項", "名稱", "商品名稱"]);
    const colorIdx = findHeaderIdx(headers, ["顏色", "規格"]);
    const stockIdx = findHeaderIdx(headers, ["庫存", "數量"]);

    if (nameIdx === -1 || stockIdx === -1) {
      return makeJson({success: false, message: "產品清單格式錯誤"});
    }

    let stockMap = {}; 
    // 建立當下最真實的庫存對照表 (In-Memory)
    for (let i = 1; i < prodData.length; i++) {
      let pName = String(prodData[i][nameIdx]).trim();
      let pColor = String(prodData[i][colorIdx] || "預設").trim();
      let pStock = Number(prodData[i][stockIdx]) || 0;
      
      // 記錄真實列號 (Row index 1-based)，用於稍後精準單獨寫回
      stockMap[pName + "-" + pColor] = { rowIndex: i + 1, currentStock: pStock };
    }

    // ==========================================
    // 🔍 修正同商品重複項目的防超賣漏洞 (Requested Qty Map)
    // ==========================================
    let requestedQtyMap = {};
    for (let item of data.itemsList) {
      let key = String(item.name).trim() + "-" + String(item.color || "預設").trim();
      if (!requestedQtyMap[key]) {
        requestedQtyMap[key] = 0;
      }
      requestedQtyMap[key] += item.quantity;
    }

    let outOfStockItems = [];
    // 核對聚合後的購物車數量
    for (let key in requestedQtyMap) {
      let reqQty = requestedQtyMap[key];
      let targetStock = stockMap[key] ? stockMap[key].currentStock : 0;
      
      if (reqQty > targetStock) {
        // 為了友善提示，把 key 切割回原本名稱
        let itemName = key.split("-")[0];
        outOfStockItems.push(`${itemName} (需求: ${reqQty}, 僅剩: ${targetStock})`);
      }
    }

    // 如果有任何一項商品超賣，立刻無條件拒絕整筆訂單
    if (outOfStockItems.length > 0) {
      return makeJson({
        success: false,
        message: `抱歉！部分商品剛剛被搶空：${outOfStockItems.join('、')}。請調整數量！`
      });
    }

    // ==========================================
    // 📝 準備批次寫入訂單與「單獨」更新庫存
    // ==========================================
    
    // 初始化標題
    if (orderSheet.getLastRow() === 0 || orderSheet.getRange("A1").getValue() === "") {
      orderSheet.getRange("A1:N1").setValues([["訂單編號", "時間", "姓名", "電話", "Email", "產品", "規格", "數量", "小計", "總計", "支付/狀態", "地址", "備註", "取貨方式"]]);
    }

    let orderRowsToInsert = [];

    data.itemsList.forEach(item => {
      const displayItemName = `[${item.status || '現貨'}] ${item.name}`;

      // 準備訂單資料 row
      orderRowsToInsert.push([
        data.orderId, 
        data.timestamp, 
        data.name, 
        "'" + data.phone,           
        data.email || "", 
        displayItemName,            
        item.color,                 
        item.quantity,              
        (item.price * item.quantity), 
        data.totalAmount,           
        data.payment,               
        data.address || "",         
        data.memo || "",            
        data.deliveryOption || ""   
      ]);
    });

    // 🚀 執行批次寫入 (Batch Write) 訂單
    if (orderRowsToInsert.length > 0) {
      // 一次性寫入所有訂單項目
      orderSheet.getRange(orderSheet.getLastRow() + 1, 1, orderRowsToInsert.length, orderRowsToInsert[0].length).setValues(orderRowsToInsert);
    }
    
    // 🚀 執行精準寫入 (Selective Write) 庫存
    // 避免覆寫整欄，只針對本次交易有牽涉到的商品單獨 setValue
    for (let key in requestedQtyMap) {
      if (stockMap[key]) {
        let newStock = stockMap[key].currentStock - requestedQtyMap[key];
        sheetProduct.getRange(stockMap[key].rowIndex, stockIdx + 1).setValue(newStock);
      }
    }
    
    SpreadsheetApp.flush(); // 強制應用所有更改
    return makeJson({success: true, message: "訂單建立成功"});

  } catch (error) {
    return makeJson({success: false, message: "系統發生錯誤：" + error.toString()});
  } finally {
    // 安全釋放 Lock
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

// ⚠️ 退貨邏輯同步套用「精準回補庫存」
function handleCancelEntireOrder(orderId) {
  const lock = LockService.getScriptLock();
  
  try {
    if (!lock.tryLock(10000)) {
      return makeJson({success: false, message: "系統繁忙，請稍後再試！"});
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = ss.getSheets()[0]; 
    const sheetProduct = ss.getSheetByName("產品清單");
    
    if (!sheetProduct) return makeJson({success: false, message: "找不到產品清單"});

    const orderData = orderSheet.getDataRange().getValues();
    const restoredItems = [];
    let found = false;

    // 準備讀取產品庫存
    const prodData = sheetProduct.getDataRange().getValues();
    const headers = prodData[0];
    const nameIdx = findHeaderIdx(headers, ["品項名稱", "品項", "名稱", "商品名稱"]);
    const colorIdx = findHeaderIdx(headers, ["顏色", "規格"]);
    const stockIdx = findHeaderIdx(headers, ["庫存", "數量"]);
    
    if (nameIdx === -1 || stockIdx === -1) return makeJson({success: false, message: "產品清單格式錯誤"});

    let stockMap = {}; 
    for (let j = 1; j < prodData.length; j++) {
      let pName = String(prodData[j][nameIdx]).trim();
      let pColor = String(prodData[j][colorIdx] || "預設").trim();
      stockMap[pName + "-" + pColor] = { rowIndex: j + 1, currentStock: Number(prodData[j][stockIdx]) || 0 };
    }
    
    // 建立要回補的數量對照表 (防重複商品疊加寫入衝突)
    let restockMap = {};

    // 處理訂單狀態與累計要退回的數量
    for (let i = 1; i < orderData.length; i++) {
      if (String(orderData[i][0]).trim() === String(orderId).trim() && orderData[i][10] !== "已取消") {
        const rawItemName = String(orderData[i][5]);
        const cleanItemName = rawItemName.replace(/^\[.*?\]\s*/, '').trim(); 
        const itemColor = String(orderData[i][6]);
        const qty = Number(orderData[i][7]);
        const price = Number(orderData[i][8]) / (qty || 1);

        restoredItems.push({ name: cleanItemName, color: itemColor, quantity: qty, price: price });

        // 取消狀態寫入 (因為退貨的 row 可能不連續，保留單筆 setValue，但這發生機率低且筆數少，影響不大)
        orderSheet.getRange(i + 1, 11).setValue("已取消");
        orderSheet.getRange(i + 1, 9).setValue(0);
        
        // 記憶體中累加要回補的庫存
        let key = cleanItemName + "-" + (itemColor || "預設").trim();
        if (!restockMap[key]) restockMap[key] = 0;
        restockMap[key] += qty;

        found = true;
      }
    }

    // 精準寫回退貨後的庫存
    for (let key in restockMap) {
      if (stockMap[key]) {
        let restoredStock = stockMap[key].currentStock + restockMap[key];
        sheetProduct.getRange(stockMap[key].rowIndex, stockIdx + 1).setValue(restoredStock);
      }
    }

    SpreadsheetApp.flush();
    return makeJson(found ? {success: true, restoredItems: restoredItems} : {success: false, message: "訂單不存在或已是取消狀態"});

  } catch (error) {
    return makeJson({success: false, message: "系統發生錯誤：" + error.toString()});
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

function findHeaderIdx(headers, keywords) {
  return headers.findIndex(h => keywords.some(k => String(h).trim().toLowerCase().includes(k.toLowerCase())));
}

// 保留原本的行為不修改
function getProductList(ss) {
  const sheet = ss.getSheetByName("產品清單");
  if (!sheet) return makeJson([]);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const idx = { 
    id: findHeaderIdx(headers, ["編號", "id", "商品編號"]), 
    name: findHeaderIdx(headers, ["品項名稱", "品項", "名稱", "商品名稱"]), 
    main: findHeaderIdx(headers, ["大分類", "主分類"]), 
    sub1: findHeaderIdx(headers, ["二分類", "子分類", "第二分類"]), 
    sub2: findHeaderIdx(headers, ["三分類", "第三分類", "子子分類"]), 
    status: findHeaderIdx(headers, ["狀態", "現貨", "預購", "商品狀態"]),     
    color: findHeaderIdx(headers, ["顏色", "規格", "樣式"]), 
    price: findHeaderIdx(headers, ["特價", "價格", "售價"]), 
    original: findHeaderIdx(headers, ["原價"]), 
    stock: findHeaderIdx(headers, ["庫存", "數量"]), 
    img: findHeaderIdx(headers, ["圖片連結", "圖片"]) 
  };
  
  const products = data.slice(1).filter(r => idx.name > -1 && idx.id > -1 && r[idx.name] && r[idx.id]).map(r => ({
    id: String(r[idx.id]), 
    name: String(r[idx.name]), 
    mainCat: idx.main > -1 ? String(r[idx.main]) : "", 
    subCat1: idx.sub1 > -1 ? String(r[idx.sub1]) : "", 
    subCat2: idx.sub2 > -1 ? String(r[idx.sub2]).trim() : "", 
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

// 恢復原本查詢總表的設定
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
