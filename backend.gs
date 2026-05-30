/**
 * 精臣商城後端系統 v26.0 - 嚴格防超賣機制升級版
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
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return makeJson({success: false, message: "無效的請求"});
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = JSON.parse(e.postData.contents);
    const orderSheet = ss.getSheets()[0]; 
    
    // ==========================================
    // 🛡️ 結帳前最後防線：嚴格檢查庫存是否足夠
    // ==========================================
    const sheetProduct = ss.getSheetByName("產品清單");
    if (sheetProduct) {
      const prodData = sheetProduct.getDataRange().getValues();
      const headers = prodData[0];
      const nameIdx = findHeaderIdx(headers, ["品項名稱", "品項", "名稱", "商品名稱"]);
      const colorIdx = findHeaderIdx(headers, ["顏色", "規格"]);
      const stockIdx = findHeaderIdx(headers, ["庫存", "數量"]);

      if (nameIdx > -1 && stockIdx > -1) {
        let stockMap = {}; 
        // 建立當下最真實的庫存對照表
        for (let i = 1; i < prodData.length; i++) {
          let pName = String(prodData[i][nameIdx]).trim();
          let pColor = String(prodData[i][colorIdx] || "預設").trim();
          let pStock = Number(prodData[i][stockIdx]) || 0;
          stockMap[pName + "-" + pColor] = { rowIndex: i + 1, currentStock: pStock };
        }

        let outOfStockItems = [];
        // 核對購物車內的每一項商品
        for (let item of data.itemsList) {
          let key = String(item.name).trim() + "-" + String(item.color || "預設").trim();
          let targetStock = stockMap[key] ? stockMap[key].currentStock : 0;
          
          if (item.quantity > targetStock) {
            outOfStockItems.push(`${item.name} (僅剩 ${targetStock} 件)`);
          }
        }

        // 如果有任何一項商品超賣，立刻無條件拒絕整筆訂單
        if (outOfStockItems.length > 0) {
          return makeJson({
            success: false,
            message: `抱歉！部分商品剛剛被搶空：${outOfStockItems.join('、')}。請調整數量！`
          });
        }
      }
    }
    // ==========================================

    if (orderSheet.getLastRow() === 0 || orderSheet.getRange("A1").getValue() === "") {
      orderSheet.getRange("A1:N1").setValues([["訂單編號", "時間", "姓名", "電話", "Email", "產品", "規格", "數量", "小計", "總計", "支付/狀態", "地址", "備註", "取貨方式"]]);
    }

    data.itemsList.forEach(item => {
      const displayItemName = `[${item.status || '現貨'}] ${item.name}`;

      orderSheet.appendRow([
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
      changeProductStock(item.name, item.color, -Math.abs(item.quantity));
    });
    
    // 回傳成功格式給前端
    return makeJson({success: true, message: "訂單建立成功"});
  } catch (error) {
    return makeJson({success: false, message: error.toString()});
  }
}

function handleCancelEntireOrder(orderId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheets()[0]; 
  const orderData = orderSheet.getDataRange().getValues();
  const restoredItems = [];
  let found = false;
  
  for (let i = 1; i < orderData.length; i++) {
    if (String(orderData[i][0]).trim() === String(orderId).trim() && orderData[i][10] !== "已取消") {
      const rawItemName = String(orderData[i][5]);
      const cleanItemName = rawItemName.replace(/^\[.*?\]\s*/, '').trim(); 
      const itemColor = String(orderData[i][6]);
      const qty = Number(orderData[i][7]);
      const price = Number(orderData[i][8]) / (qty || 1);

      restoredItems.push({ name: cleanItemName, color: itemColor, quantity: qty, price: price });

      orderSheet.getRange(i + 1, 11).setValue("已取消");
      orderSheet.getRange(i + 1, 9).setValue(0);
      
      changeProductStock(cleanItemName, itemColor, qty);
      found = true;
    }
  }

  SpreadsheetApp.flush();
  return makeJson(found ? {success: true, restoredItems: restoredItems} : {success: false, message: "訂單不存在或已是取消狀態"});
}

function findHeaderIdx(headers, keywords) {
  return headers.findIndex(h => keywords.some(k => String(h).trim().toLowerCase().includes(k.toLowerCase())));
}

function changeProductStock(itemName, itemColor, delta) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("產品清單");
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameIdx = findHeaderIdx(headers, ["品項名稱", "品項", "名稱", "商品名稱"]);
  const colorIdx = findHeaderIdx(headers, ["顏色", "規格"]);
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
