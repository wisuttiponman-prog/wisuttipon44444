const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'getProducts';
    const ss = SpreadsheetApp.openById(SHEET_ID);
    
    if (action === 'getProducts') {
      const sheet = ss.getSheetByName('Products');
      if (!sheet) return responseJSON({ status: 'error', message: 'ไม่พบแท็บชื่อ Products ใน Google Sheets' });
      
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return responseJSON({ status: 'success', data: [] });
      
      const headers = data.shift().map(h => String(h).trim());
      const products = data
        .filter(row => row[0] !== '')
        .map(row => {
          let obj = {};
          headers.forEach((h, i) => obj[h] = row[i]);
          return obj;
        });
      return responseJSON({ status: 'success', data: products });
    }
    
    if (action === 'getOrders') {
      const sheet = ss.getSheetByName('Orders');
      if (!sheet) return responseJSON({ status: 'error', message: 'ไม่พบแท็บชื่อ Orders ใน Google Sheets' });
      
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return responseJSON({ status: 'success', data: [] });
      
      const headers = data.shift().map(h => String(h).trim());
      const orders = data
        .filter(row => row[0] !== '')
        .map(row => {
          let obj = {};
          headers.forEach((h, i) => obj[h] = row[i]);
          return obj;
        }).reverse();
      return responseJSON({ status: 'success', data: orders });
    }
    
    return responseJSON({ status: 'error', message: 'Invalid action' });
  } catch (err) {
    return responseJSON({ status: 'error', message: err.toString() });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    
    if (payload.action === 'createOrder') {
      const orderSheet = ss.getSheetByName('Orders');
      const prodSheet = ss.getSheetByName('Products');
      const prodData = prodSheet.getDataRange().getValues();
      
      const orderId = 'ORD-' + Math.floor(100000 + Math.random() * 900000);
      const timestamp = new Date().toLocaleString('th-TH');
      
      payload.items.forEach(item => {
        for (let i = 1; i < prodData.length; i++) {
          if (String(prodData[i][0]).trim() === String(item.sku).trim()) {
            let currentStock = Number(prodData[i][4]);
            prodSheet.getRange(i + 1, 5).setValue(Math.max(0, currentStock - item.qty));
          }
        }
      });

      const itemsSummary = payload.items.map(i => `${i.sku} (${i.name}) x${i.qty} [${i.location}]`).join(' | ');
      
      orderSheet.appendRow([
        orderId,
        timestamp,
        payload.customerName,
        payload.phone,
        payload.address,
        itemsSummary,
        payload.total,
        'PENDING_PICK',
        ''
      ]);
      
      return responseJSON({ status: 'success', orderId: orderId });
    }
    
    if (payload.action === 'updateStatus') {
      const orderSheet = ss.getSheetByName('Orders');
      const data = orderSheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === String(payload.orderId).trim()) {
          orderSheet.getRange(i + 1, 8).setValue(payload.newStatus);
          if (payload.trackingNo) {
            orderSheet.getRange(i + 1, 9).setValue(payload.trackingNo);
          }
          break;
        }
      }
      return responseJSON({ status: 'success' });
    }
  } catch (err) {
    return responseJSON({ status: 'error', message: err.toString() });
  }
}

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
