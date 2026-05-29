// ==========================================
// 1. 全局變數設定
// ==========================================
const SHEET_ID = '12uJmKGYTPzJ6lA3mm-VVYzWS48RLZkte8RbxJ7arWvQ'; // 🔥 記得換回您的 ID

const CUSTOMER_SHEET_NAME = '全訂單紀錄總表'; 
const STAFF_SHEET_NAME = '工作人員狀態'; 

// ==========================================
// 2. 讀取 API：將兩張表合併後傳給手機
// ==========================================
function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    
    const customerSheet = ss.getSheetByName(CUSTOMER_SHEET_NAME);
    if (!customerSheet) throw new Error("找不到分頁：" + CUSTOMER_SHEET_NAME);
    
    const customerData = customerSheet.getDataRange().getDisplayValues();
    const customerRows = customerData.slice(1); 
    
    let staffData = [];
    const staffSheet = ss.getSheetByName(STAFF_SHEET_NAME);
    if (staffSheet) {
      staffData = staffSheet.getDataRange().getDisplayValues();
    }
    const staffRows = staffData.length > 1 ? staffData.slice(1) : [];
    
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
          isCompleted: row[8] === '已結案',
          staffNote: row[9] || '' // 🔥 讀取 J 欄備註
        };
      }
    });

    const ordersMap = {};

    customerRows.forEach(row => {
      const orderId = row[0]; 
      if (!orderId) return; 

      if (!ordersMap[orderId]) {
        const status = staffStatusMap[orderId] || {}; 
        const paymentMethod = row[10]; 
        
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
          staffNote: status.staffNote || '', // 🔥 傳給前端備註
          
          items: [] 
        };
      }
      
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
    return ContentService.createTextOutput(JSON.stringify({ error: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 3. 更新 API：員工滑動狀態時寫入分頁
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
      if (payload.createdBy) {
        staffSheet.getRange(foundRowIndex, 6).setValue(payload.createdBy);
      }
      staffSheet.getRange(foundRowIndex, 7).setValue(payload.lastModifiedBy);
      staffSheet.getRange(foundRowIndex, 8).setValue(payload.lastModifiedTime);
      staffSheet.getRange(foundRowIndex, 9).setValue(isCompletedStr);
      staffSheet.getRange(foundRowIndex, 10).setValue(payload.staffNote || ''); // 🔥 寫入 J 欄備註
      
    } else {
      staffSheet.appendRow([
        targetOrderId,               
        payload.customerName || '',  
        isPaidStr,                   
        payload.deliveryType || '',  
        payload.pickupTime || '',    
        payload.createdBy || payload.lastModifiedBy,     
        payload.lastModifiedBy,      
        payload.lastModifiedTime,    
        isCompletedStr,
        payload.staffNote || '' // 🔥 寫入 J 欄備註
      ]);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: 'success', message: '狀態更新成功！' }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
