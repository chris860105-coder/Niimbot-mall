/**
 * 精臣商城後端系統 v28.0 - Sheets API 庫存批次更新進階版
 * 必須開啟 Google Sheets Advanced Service 才可運行
 */

function doGet(e) {
  if (!e || !e.parameter) {
    return ContentService.createTextOutput("🎉 後端服務運作正常中！\n請使用部署後的網址進行對接，不要直接在編輯器內點擊執行 doGet。").setMimeType(ContentService.MimeType.TEXT);
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
  const tTotal = Date.now();
  let tLast = tTotal;

  // === 鎖外作業：無共享狀態的資料驗證提早做，不佔用 Lock 時間 ===
  if (!e || !e.postData || !e.postData.contents) {
    return makeJson({success: false, message: "無效的請求"});
  }

  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return makeJson({success: false, message: "JSON 格式錯誤"});
  }

  if (!Array.isArray(data.itemsList) || data.itemsList.length === 0) {
    return makeJson({success: false, message: "購物車不能為空"});
  }

  // 安全性驗證與正規化 qty，同時合併重複商品 (防超賣漏洞)
  let requestedQtyMap = {};
  let cleanCart = [];
  let totalAmount = 0;

  for (let i = 0; i < data.itemsList.length; i++) {
    let item = data.itemsList[i];
    let qty = Number(item.quantity);
    
    if (!Number.isFinite(qty) || qty <= 0 || !Number.isInteger(qty)) {
      return makeJson({success: false, message: "商品數量格式錯誤"});
    }

    item.quantity = qty;
    let cleanName = String(item.name || "未知名稱").trim();
    let cleanColor = String(item.color || "預設").trim();
    let key = cleanName + "-" + cleanColor;
    
    requestedQtyMap[key] = (requestedQtyMap[key] || 0) + qty;
    cleanCart.push(item);
    totalAmount += (Number(item.price) || 0) * qty;
  }
  
  data.itemsList = cleanCart;
  data.totalAmount = totalAmount;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName("全訂單紀錄總表");
  if (!orderSheet) {
    return makeJson({success: false, message: "找不到訂單總表"});
  }

  const sheetProduct = ss.getSheetByName("產品清單");
  if (!sheetProduct) {
    return makeJson({success: false, message: "找不到產品清單"});
  }

  console.log("[PERF][doPost] request validation: " + (Date.now() - tLast) + "ms");
  tLast = Date.now();

  // === 進入危險區 (Critical Section) ===
  const lock = LockService.getScriptLock();
  
  try {
    if (!lock.tryLock(10000)) {
      return makeJson({success: false, message: "系統繁忙，許多人正在結帳，請稍後再試！"});
    }

    console.log("[PERF][doPost] lock wait: " + (Date.now() - tLast) + "ms");
    tLast = Date.now();

    // 初始化標題 (移入 Critical Section)
    if (orderSheet.getLastRow() === 0 || orderSheet.getRange("A1").getValue() === "") {
      orderSheet.getRange("A1:N1").setValues([["訂單編號", "時間", "姓名", "電話", "Email", "產品", "規格", "數量", "小計", "總計", "支付/狀態", "地址", "備註", "取貨方式"]]);
    }

    // 1. 讀取最新產品庫存
    const prodData = sheetProduct.getDataRange().getValues();
    const headers = prodData[0];
    const nameIdx = findHeaderIdx(headers, ["品項名稱", "品項", "名稱", "商品名稱"]);
    const colorIdx = findHeaderIdx(headers, ["顏色", "規格"]);
    const stockIdx = findHeaderIdx(headers, ["庫存", "數量"]);

    if (nameIdx === -1 || stockIdx === -1) {
      return makeJson({success: false, message: "產品清單格式錯誤"});
    }

    console.log("[PERF][doPost] product read: " + (Date.now() - tLast) + "ms");
    tLast = Date.now();

    let stockMap = {}; 
    for (let i = 1; i < prodData.length; i++) {
      let pName = String(prodData[i][nameIdx]).trim();
      let pColor = String(prodData[i][colorIdx] || "預設").trim();
      let pStock = Number(prodData[i][stockIdx]) || 0;
      stockMap[pName + "-" + pColor] = { rowIndex: i + 1, currentStock: pStock };
    }

    // 2. 驗證庫存 (防超賣)
    let outOfStockItems = [];
    for (let key in requestedQtyMap) {
      let targetStock = stockMap[key] ? stockMap[key].currentStock : 0;
      if (requestedQtyMap[key] > targetStock) {
        let itemName = key.split("-")[0];
        outOfStockItems.push(`${itemName} (僅剩 ${targetStock} 件)`);
      }
    }

    if (outOfStockItems.length > 0) {
      return makeJson({ success: false, message: `抱歉！部分商品剛剛被搶空：${outOfStockItems.join('、')}。請調整數量！` });
    }

    console.log("[PERF][doPost] stock map & validation: " + (Date.now() - tLast) + "ms");
    tLast = Date.now();

    // 3. 準備訂單批次寫入
    let orderRowsToInsert = [];
    data.itemsList.forEach(item => {
      const displayItemName = `[${item.status || '現貨'}] ${item.name}`;
      orderRowsToInsert.push([
        data.orderId, 
        data.timestamp, 
        data.name, 
        "'" + data.phone,           
        data.email || "", 
        displayItemName,            
        item.color,                 
        item.quantity,              
        (Number(item.price) * item.quantity), 
        data.totalAmount,           
        data.payment,               
        data.address || "",         
        data.memo || "",            
        data.deliveryOption || ""   
      ]);
    });

    if (orderRowsToInsert.length > 0) {
      orderSheet.getRange(orderSheet.getLastRow() + 1, 1, orderRowsToInsert.length, 14).setValues(orderRowsToInsert);
    }
    
    console.log("[PERF][doPost] order write: " + (Date.now() - tLast) + "ms");
    tLast = Date.now();

    // 4. 批次更新庫存 (使用 Google Sheets Advanced Service)
    const stockColLetter = columnIndexToLetter(stockIdx + 1);
    const stockUpdates = [];

    for (let key in requestedQtyMap) {
      if (stockMap[key]) {
        let newStock = stockMap[key].currentStock - requestedQtyMap[key];
        stockUpdates.push({
          range: "'產品清單'!" + stockColLetter + stockMap[key].rowIndex,
          values: [[newStock]]
        });
      }
    }
    
    if (stockUpdates.length > 0) {
      Sheets.Spreadsheets.Values.batchUpdate({
        valueInputOption: "RAW",
        data: stockUpdates
      }, ss.getId());
    }
    
    console.log("[PERF][doPost] stock write: " + (Date.now() - tLast) + "ms");
    tLast = Date.now();

    // 確保持久化寫入，避免下一個取得 Lock 的請求讀到舊資料
    SpreadsheetApp.flush();
    
    console.log("[PERF][doPost] flush: " + (Date.now() - tLast) + "ms");
    console.log("[PERF][doPost] TOTAL: " + (Date.now() - tTotal) + "ms");

    return makeJson({success: true, message: "訂單建立成功"});

  } catch (error) {
    return makeJson({success: false, message: error.toString()});
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

function handleCancelEntireOrder(orderId) {
  const tTotal = Date.now();
  let tLast = tTotal;

  if (!orderId) return makeJson({success: false, message: "無效的訂單編號"});

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName("全訂單紀錄總表");
  if (!orderSheet) return makeJson({success: false, message: "找不到訂單總表"});

  const sheetProduct = ss.getSheetByName("產品清單");
  if (!sheetProduct) return makeJson({success: false, message: "找不到產品清單"});

  // 1. TextFinder 鎖定目標訂單 (鎖外作業)
  const lastRow = orderSheet.getLastRow();
  if (lastRow < 2) return makeJson({success: false, message: "訂單不存在"});

  const foundCells = orderSheet
    .getRange(2, 1, lastRow - 1, 1)
    .createTextFinder(String(orderId).trim())
    .matchEntireCell(true)
    .findAll();

  if (foundCells.length === 0) {
    return makeJson({success: false, message: "訂單不存在"});
  }

  // 取出符合條件的所有行號，並排序
  const rowNumbers = foundCells.map(c => c.getRow()).sort((a, b) => a - b);

  console.log("[PERF][handleCancelEntireOrder] TextFinder: " + (Date.now() - tLast) + "ms");
  tLast = Date.now();

  // === 進入危險區 ===
  const lock = LockService.getScriptLock();
  
  try {
    if (!lock.tryLock(10000)) {
      return makeJson({success: false, message: "系統繁忙，請稍後再試！"});
    }

    console.log("[PERF][handleCancelEntireOrder] lock wait: " + (Date.now() - tLast) + "ms");

    let orderReadTotal = 0;
    let orderWriteTotal = 0;

    const restoredItems = [];
    let restockMap = {};
    let found = false;

    // 判斷是否為連續列
    let isContinuous = true;
    for (let i = 1; i < rowNumbers.length; i++) {
      if (rowNumbers[i] !== rowNumbers[i - 1] + 1) {
        isContinuous = false;
        break;
      }
    }

    if (isContinuous && rowNumbers.length > 0) {
      // === 連續列優化處理 ===
      const startRow = rowNumbers[0];
      const rowCount = rowNumbers.length;
      
      let tr0 = Date.now();
      // 1次 Read
      const rowDataArray = orderSheet.getRange(startRow, 1, rowCount, 14).getValues();
      orderReadTotal += (Date.now() - tr0);

      const updateDataArray = [];
      let updatesNeeded = false;

      for (let i = 0; i < rowDataArray.length; i++) {
        let rowData = rowDataArray[i];
        if (String(rowData[0]).trim() === String(orderId).trim() && rowData[10] !== "已取消") {
          found = true;
          updatesNeeded = true;
          const rawItemName = String(rowData[5]);
          const cleanItemName = rawItemName.replace(/^\[.*?\]\s*/, '').trim(); 
          const itemColor = String(rowData[6]);
          const qty = Number(rowData[7]) || 1;
          const price = Number(rowData[8]) / qty;

          restoredItems.push({ name: cleanItemName, color: itemColor, quantity: qty, price: price });
          
          let key = cleanItemName + "-" + (itemColor || "預設").trim();
          restockMap[key] = (restockMap[key] || 0) + qty;

          // [小計=0, 總計不變, 狀態="已取消"]
          updateDataArray.push([0, rowData[9], "已取消"]);
        } else {
          updateDataArray.push([rowData[8], rowData[9], rowData[10]]); // 保持原樣
        }
      }

      if (updatesNeeded) {
        let tw0 = Date.now();
        // 1次 Write
        orderSheet.getRange(startRow, 9, rowCount, 3).setValues(updateDataArray);
        orderWriteTotal += (Date.now() - tw0);
      }

    } else {
      // === 非連續列 Fallback 處理 ===
      for (let r of rowNumbers) {
        let tr0 = Date.now();
        const rowData = orderSheet.getRange(r, 1, 1, 14).getValues()[0];
        orderReadTotal += (Date.now() - tr0);
        
        if (String(rowData[0]).trim() === String(orderId).trim() && rowData[10] !== "已取消") {
          found = true;
          const rawItemName = String(rowData[5]);
          const cleanItemName = rawItemName.replace(/^\[.*?\]\s*/, '').trim(); 
          const itemColor = String(rowData[6]);
          const qty = Number(rowData[7]) || 1;
          const price = Number(rowData[8]) / qty;

          restoredItems.push({ name: cleanItemName, color: itemColor, quantity: qty, price: price });

          let tw0 = Date.now();
          // 1次 Write (單列)
          orderSheet.getRange(r, 9, 1, 3).setValues([[0, rowData[9], "已取消"]]);
          orderWriteTotal += (Date.now() - tw0);
          
          let key = cleanItemName + "-" + (itemColor || "預設").trim();
          restockMap[key] = (restockMap[key] || 0) + qty;
        }
      }
    }

    console.log("[PERF][handleCancelEntireOrder] order read: " + orderReadTotal + "ms");
    console.log("[PERF][handleCancelEntireOrder] order write: " + orderWriteTotal + "ms");

    if (!found) {
      return makeJson({success: false, message: "訂單已是取消狀態或不存在"});
    }

    let tp0 = Date.now();
    // 2. 準備讀取產品庫存以進行回補
    const prodData = sheetProduct.getDataRange().getValues();
    const headers = prodData[0];
    const nameIdx = findHeaderIdx(headers, ["品項名稱", "品項", "名稱", "商品名稱"]);
    const colorIdx = findHeaderIdx(headers, ["顏色", "規格"]);
    const stockIdx = findHeaderIdx(headers, ["庫存", "數量"]);

    if (nameIdx === -1 || stockIdx === -1) {
       return makeJson({success: false, message: "產品清單格式錯誤"});
    }

    let stockMap = {};
    for (let j = 1; j < prodData.length; j++) {
      let pName = String(prodData[j][nameIdx]).trim();
      let pColor = String(prodData[j][colorIdx] || "預設").trim();
      stockMap[pName + "-" + pColor] = { rowIndex: j + 1, currentStock: Number(prodData[j][stockIdx]) || 0 };
    }

    console.log("[PERF][handleCancelEntireOrder] product read: " + (Date.now() - tp0) + "ms");

    let ts0 = Date.now();
    // 3. 精準寫回退貨後的庫存 (使用 Google Sheets Advanced Service)
    const stockColLetter = columnIndexToLetter(stockIdx + 1);
    const stockUpdates = [];

    for (let key in restockMap) {
      if (stockMap[key]) {
        let restoredStock = stockMap[key].currentStock + restockMap[key];
        stockUpdates.push({
          range: "'產品清單'!" + stockColLetter + stockMap[key].rowIndex,
          values: [[restoredStock]]
        });
      }
    }

    if (stockUpdates.length > 0) {
      Sheets.Spreadsheets.Values.batchUpdate({
        valueInputOption: "RAW",
        data: stockUpdates
      }, ss.getId());
    }

    console.log("[PERF][handleCancelEntireOrder] stock write: " + (Date.now() - ts0) + "ms");

    let tf0 = Date.now();
    // 確保持久化寫入
    SpreadsheetApp.flush();
    console.log("[PERF][handleCancelEntireOrder] flush: " + (Date.now() - tf0) + "ms");
    console.log("[PERF][handleCancelEntireOrder] TOTAL: " + (Date.now() - tTotal) + "ms");

    return makeJson({success: true, restoredItems: restoredItems});

  } catch (error) {
    return makeJson({success: false, message: error.toString()});
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

// 將欄位索引 (1-based) 轉換為英文欄名 A, B, C... Z, AA...
function columnIndexToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function findHeaderIdx(headers, keywords) {
  return headers.findIndex(h => keywords.some(k => String(h).trim().toLowerCase().includes(k.toLowerCase())));
}

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
