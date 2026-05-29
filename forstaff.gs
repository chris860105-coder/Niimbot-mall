// 替換成你儲存訂單的那份 Google Sheet 的 ID (在網址列 d/ 跟 /edit 之間那段)
const SHEET_ID = '12uJmKGYTPzJ6lA3mm-VVYzWS48RLZkte8RbxJ7arWvQ'; 

// 🔥 已經幫你更新為你實際的分頁名稱
const CUSTOMER_SHEET_NAME = '全訂單紀錄總表'; // 客人原本寫入的分頁
const STAFF_SHEET_NAME = '工作人員狀態'; // 🔥 請確認你有在試算表底下新增這個分頁喔！

// ==========================================
// 1. 讀取 API：將兩張表用「訂單編號」合併後傳給手機
// ==========================================
function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    
    // --- 讀取客人的原始資料 ---
    const customerSheet = ss.getSheetByName(CUSTOMER_SHEET_NAME);
    if (!customerSheet) throw new Error("找不到分頁：" + CUSTOMER_SHEET_NAME);
    
    const customerData = customerSheet.getDataRange().getDisplayValues();
    const customerRows = customerData.slice(1); // 略過表頭
    
    // --- 讀取工作人員的狀態資料 ---
    let staffData = [];
    const staffSheet = ss.getSheetByName(STAFF_SHEET_NAME);
    if (staffSheet) {
      staffData = staffSheet.getDataRange().getDisplayValues();
    }
    const staffRows = staffData.length > 1 ? staffData.slice(1) : [];
    
    // 建立 Staff 狀態字典，用訂單編號當 Key
    const staffStatusMap = {};
    staffRows.forEach(row => {
      const orderId = row[0];
      if (orderId) {
        staffStatusMap[orderId] = {
          isPaid: row[2] === '已付',
          deliveryType: row[3] || '',
          pickupTime: row[4] || '',
          createdBy: row[5] || '',
          lastModifiedBy: row[6] || '',
          lastModifiedTime: row[7] || '',
          isCompleted: row[8] === '已結案'
        };
      }
    });

    // 用來「合併」相同訂單編號的物件
    const ordersMap = {};

    customerRows.forEach(row => {
      const orderId = row[0]; // A欄：訂單編號
      if (!orderId) return; 

      // 初始化訂單並合併 Staff 狀態
      if (!ordersMap[orderId]) {
        const status = staffStatusMap[orderId] || {}; 
        const paymentMethod = row[10]; // K欄：支付方式
        
        ordersMap[orderId] = {
          id: orderId,
          orderTime: row[1],
          customerName: row[2],
          phone: row[3],
          email: row[4],
          totalAmount: row[9],
          paymentMethod: paymentMethod,
          address: row[11],
          note: row[12],
          customerDelivery: row[13],
          
          isCancelled: paymentMethod === '取消',
          
          isPaid: status.isPaid || false,
          deliveryType: status.deliveryType || '',
          pickupTime: status.pickupTime || '',
          createdBy: status.createdBy || '',
          lastModifiedBy: status.lastModifiedBy || '',
          lastModifiedTime: status.lastModifiedTime || '',
          isCompleted: status.isCompleted || false,
          
          items: [] 
        };
      }
      
      // 把商品塞進 items 陣列 (F, G, H, I 欄)
      ordersMap[orderId].items.push({
        type: row[5],
        name: row[6],
        qty: row[7],
        subtotal: row[8]
      });
    });

    const result = Object.values(ordersMap);
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    // 捕捉錯誤並回傳，避免前端 CORS 阻擋
    return ContentService.createTextOutput(JSON.stringify({ error: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 2. 更新 API：員工滑動狀態時，只更新或寫入「工作人員狀態」分頁
// ==========================================
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const targetOrderId = payload.id;
    
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const staffSheet = ss.getSheetByName(STAFF_SHEET_NAME);
    
    if (!staffSheet) {
       throw new Error(`找不到分頁「${STAFF_SHEET_NAME}」，請先在試算表中建立！`);
    }

    const staffData = staffSheet.getDataRange().getValues();
    let foundRowIndex = -1;
    
    for (let i = 1; i < staffData.length; i++) {
      if (staffData[i][0] === targetOrderId) {
        foundRowIndex = i + 1; 
        break;
      }
    }
    
    const isPaidStr = payload.isPaid ? '已付' : '';
    const isCompletedStr = payload.isCompleted ? '已結案' : '';
    
    if (foundRowIndex !== -1) {
      staffSheet.getRange(foundRowIndex, 3).setValue(isPaidStr);
      staffSheet.getRange(foundRowIndex, 4).setValue(payload.deliveryType || '');
      staffSheet.getRange(foundRowIndex, 5).setValue(payload.pickupTime || '');
      
      if (!staffData[foundRowIndex - 1][5]) {
        staffSheet.getRange(foundRowIndex, 6).setValue(payload.lastModifiedBy);
      }
      
      staffSheet.getRange(foundRowIndex, 7).setValue(payload.lastModifiedBy);
      staffSheet.getRange(foundRowIndex, 8).setValue(payload.lastModifiedTime);
      staffSheet.getRange(foundRowIndex, 9).setValue(isCompletedStr);
      
    } else {
      staffSheet.appendRow([
        targetOrderId,               
        payload.customerName || '',  
        isPaidStr,                   
        payload.deliveryType || '',  
        payload.pickupTime || '',    
        payload.lastModifiedBy,      
        payload.lastModifiedBy,      
        payload.lastModifiedTime,    
        isCompletedStr               
      ]);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: 'success', message: '狀態更新成功！' }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
