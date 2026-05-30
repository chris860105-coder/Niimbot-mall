// ==========================================
// 1. 全局變數設定
// ==========================================
const SHEET_ID = '12uJmKGYTPzJ6lA3mm-VVYzWS48RLZkte8RbxJ7arWvQ'; // 🔥 記得換回您的 ID

const CUSTOMER_SHEET_NAME = '全訂單紀錄總表'; 
const STAFF_SHEET_NAME = '工作人員狀態'; 
const ACCOUNT_SHEET_NAME = '員工帳號管理'; // 🔥 新增
const LOG_SHEET_NAME = '操作紀錄'; // 🔥 新增

function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const customerSheet = ss.getSheetByName(CUSTOMER_SHEET_NAME);
    if (!customerSheet) throw new Error("找不到分頁：" + CUSTOMER_SHEET_NAME);
    
    const customerData = customerSheet.getDataRange().getDisplayValues();
    const customerRows = customerData.slice(1); 
    
    let staffData = [];
    const staffSheet = ss.getSheetByName(STAFF_SHEET_NAME);
    if (staffSheet) staffData = staffSheet.getDataRange().getDisplayValues();
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
          staffNote: row[9] || ''
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
          staffNote: status.staffNote || '',
          items: [] 
        };
      }
      
      ordersMap[orderId].items.push({
        type: row[5], name: row[6], qty: row[7], subtotal: row[8]
      });
    });

    return ContentService.createTextOutput(JSON.stringify(Object.values(ordersMap))).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);

    // ========================================
    // 🔥 功能 A：處理「登入驗證」
    // ========================================
    if (payload.action === 'login') {
      const accSheet = ss.getSheetByName(ACCOUNT_SHEET_NAME);
      if (!accSheet) throw new Error(`找不到分頁「${ACCOUNT_SHEET_NAME}」`);
      
      const accData = accSheet.getDataRange().getValues();
      for (let i = 1; i < accData.length; i++) {
        // A欄:帳號, B欄:密碼, C欄:姓名, D欄:狀態
        if (String(accData[i][0]).trim() === String(payload.account).trim() && 
            String(accData[i][1]).trim() === String(payload.password).trim()) {
          
          if (accData[i][3] === '停用') {
            return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: '此帳號已被停用，請聯絡管理員' })).setMimeType(ContentService.MimeType.JSON);
          }
          return ContentService.createTextOutput(JSON.stringify({ status: 'success', staffName: accData[i][2] })).setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: '帳號或密碼錯誤' })).setMimeType(ContentService.MimeType.JSON);
    }

    // ========================================
    // 🔥 功能 B：處理「訂單狀態更新與紀錄」
    // ========================================
    if (payload.action === 'updateOrder') {
      const staffSheet = ss.getSheetByName(STAFF_SHEET_NAME);
      const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
      if (!staffSheet) throw new Error(`找不到分頁「${STAFF_SHEET_NAME}」`);

      const targetOrderId = payload.id;
      const staffData = staffSheet.getDataRange().getValues();
      let foundRowIndex = -1;
      
      for (let i = 1; i < staffData.length; i++) {
        if (staffData[i][0] === targetOrderId) { foundRowIndex = i + 1; break; }
      }
      
      const isPaidStr = payload.isPaid ? '已付' : '';
      const isCompletedStr = payload.isCompleted ? '已結案' : '';
      
      // -- 寫入 Log 的小工具 --
      const addLog = (field, oldVal, newVal) => {
        if (oldVal !== newVal && logSheet) {
          const now = new Date();
          const timeStr = `${now.getFullYear()}/${now.getMonth()+1}/${now.getDate()} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}:${String(now.getSeconds()).padStart(2, '0')}`;
          logSheet.appendRow([timeStr, payload.lastModifiedBy, targetOrderId, field, oldVal, newVal]);
        }
      };

      if (foundRowIndex !== -1) {
        // 比對差異並寫入 Log
        addLog('付款狀態', staffData[foundRowIndex - 1][2] || '', isPaidStr);
        addLog('交付狀態', staffData[foundRowIndex - 1][3] || '', payload.deliveryType || '');
        addLog('自取時間', staffData[foundRowIndex - 1][4] || '', payload.pickupTime || '');
        addLog('結案狀態', staffData[foundRowIndex - 1][8] || '', isCompletedStr);
        addLog('人員備註', staffData[foundRowIndex - 1][9] || '', payload.staffNote || '');

        // 更新儲存格
        staffSheet.getRange(foundRowIndex, 3).setValue(isPaidStr);
        staffSheet.getRange(foundRowIndex, 4).setValue(payload.deliveryType || '');
        staffSheet.getRange(foundRowIndex, 5).setValue(payload.pickupTime || '');
        if (payload.createdBy && !staffData[foundRowIndex - 1][5]) staffSheet.getRange(foundRowIndex, 6).setValue(payload.createdBy);
        staffSheet.getRange(foundRowIndex, 7).setValue(payload.lastModifiedBy);
        staffSheet.getRange(foundRowIndex, 8).setValue(payload.lastModifiedTime);
        staffSheet.getRange(foundRowIndex, 9).setValue(isCompletedStr);
        staffSheet.getRange(foundRowIndex, 10).setValue(payload.staffNote || ''); 
        
      } else {
        addLog('建立訂單狀態', '', '首次操作');
        staffSheet.appendRow([
          targetOrderId, payload.customerName || '', isPaidStr, payload.deliveryType || '', payload.pickupTime || '',    
          payload.createdBy || payload.lastModifiedBy, payload.lastModifiedBy, payload.lastModifiedTime, isCompletedStr, payload.staffNote || ''
        ]);
      }
      
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', message: '狀態更新成功！' })).setMimeType(ContentService.MimeType.JSON);
    }
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}
