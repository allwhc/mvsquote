/**
 * ============================================
 * MV SOLUTIONS - QUOTATION SYSTEM
 * Google Apps Script Backend v4 (Simplified)
 * ============================================
 * 
 * SHEETS STRUCTURE:
 * - Items: Components master list
 * - Products: Product master list
 * - BOM: Product → Component mapping
 * - Sub_Products: Product → Product mapping (with Qty)
 * - Quotations: Quotation headers
 * - Quotation_Items: Quotation line items
 * 
 * ============================================
 */

const SHEETS = {
  ITEMS: 'Items',
  PRODUCTS: 'Products',
  BOM: 'BOM',
  SUB_PRODUCTS: 'Sub_Products',
  QUOTATIONS: 'Quotations',
  QUOTATION_ITEMS: 'Quotation_Items'
};

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  
  if (!sheet) {
    sheet = ss.insertSheet(name);
    initializeSheetHeaders(sheet, name);
  }
  
  return sheet;
}

function initializeSheetHeaders(sheet, name) {
  const headers = {
    [SHEETS.ITEMS]: ['Code', 'Name', 'Price', 'Category', 'Vendor', 'Vendor Contact'],
    [SHEETS.PRODUCTS]: ['SKU', 'Name', 'Category', 'Selling Price', 'MRP', 'D2C Rate', 'Amazon Rate', 'Flipkart Rate', 'Dealer Rate'],
    [SHEETS.BOM]: ['Product SKU', 'Component Code', 'Qty'],
    [SHEETS.SUB_PRODUCTS]: ['Product SKU', 'Sub Product SKU', 'Qty'],
    [SHEETS.QUOTATIONS]: ['ID', 'Number', 'Date', 'Customer Name', 'Customer Phone', 'Customer Address', 'Subtotal', 'GST', 'Grand Total', 'Notes', 'Status'],
    [SHEETS.QUOTATION_ITEMS]: ['Quotation ID', 'Product SKU', 'Qty', 'MRP', 'Discount %', 'Net Price', 'Total']
  };
  
  if (headers[name]) {
    sheet.appendRow(headers[name]);
    sheet.getRange(1, 1, 1, headers[name].length)
      .setFontWeight('bold')
      .setBackground('#1E3A8A')
      .setFontColor('white');
    sheet.setFrozenRows(1);
  }
}

// ============ ITEMS FUNCTIONS ============

function getItems() {
  const sheet = getSheet(SHEETS.ITEMS);
  const data = sheet.getDataRange().getValues();
  const items = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      items.push({
        code: data[i][0],
        name: data[i][1],
        price: Number(data[i][2]) || 0,
        category: data[i][3] || 'Electronics',
        vendor: data[i][4] || '',
        vendorAddress: data[i][5] || ''
      });
    }
  }
  
  return items;
}

function syncItems(items) {
  const sheet = getSheet(SHEETS.ITEMS);
  clearSheetData(sheet);

  if (items && items.length > 0) {
    // Log first item to debug
    Logger.log('First item received: ' + JSON.stringify(items[0]));

    const data = items.map(item => [
      item.code || '',
      item.name || '',
      item.price || 0,
      item.category || 'Electronics',
      item.vendor || '',
      item.vendorAddress || ''
    ]);

    // Log data being written
    Logger.log('Writing data rows: ' + data.length + ', columns: 6');
    Logger.log('First row: ' + JSON.stringify(data[0]));

    sheet.getRange(2, 1, data.length, 6).setValues(data);
  }

  return { success: true, count: items ? items.length : 0 };
}

// ============ PRODUCTS FUNCTIONS ============

function getProducts() {
  const sheet = getSheet(SHEETS.PRODUCTS);
  const data = sheet.getDataRange().getValues();
  const products = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      products.push({
        sku: data[i][0],
        name: data[i][1],
        category: data[i][2],
        sellingPrice: data[i][3] ? Number(data[i][3]) : null,
        mrp: data[i][4] ? Number(data[i][4]) : null,
        d2cRate: data[i][5] !== '' && data[i][5] !== undefined ? Number(data[i][5]) : 3,
        amazonRate: data[i][6] !== '' && data[i][6] !== undefined ? Number(data[i][6]) : 12,
        flipkartRate: data[i][7] !== '' && data[i][7] !== undefined ? Number(data[i][7]) : 11,
        dealerRate: data[i][8] !== '' && data[i][8] !== undefined ? Number(data[i][8]) : 500
      });
    }
  }

  return products;
}

function syncProducts(products) {
  const sheet = getSheet(SHEETS.PRODUCTS);
  clearSheetData(sheet);

  if (products && products.length > 0) {
    const data = products.map(p => [
      p.sku,
      p.name,
      p.category,
      p.sellingPrice || '',
      p.mrp || '',
      p.d2cRate !== undefined ? p.d2cRate : 3,
      p.amazonRate !== undefined ? p.amazonRate : 12,
      p.flipkartRate !== undefined ? p.flipkartRate : 11,
      p.dealerRate !== undefined ? p.dealerRate : 500
    ]);
    sheet.getRange(2, 1, data.length, 9).setValues(data);
  }

  return { success: true, count: products ? products.length : 0 };
}

// ============ BOM FUNCTIONS ============

function getBOM() {
  const sheet = getSheet(SHEETS.BOM);
  const data = sheet.getDataRange().getValues();
  const bom = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][1]) {
      bom.push({
        productSku: data[i][0],
        componentCode: data[i][1],
        qty: Number(data[i][2]) || 1
      });
    }
  }
  
  return bom;
}

function syncBOM(bom) {
  const sheet = getSheet(SHEETS.BOM);
  clearSheetData(sheet);
  
  if (bom && bom.length > 0) {
    const data = bom.map(b => [b.productSku, b.componentCode, b.qty]);
    sheet.getRange(2, 1, data.length, 3).setValues(data);
  }
  
  return { success: true, count: bom ? bom.length : 0 };
}

// ============ SUB_PRODUCTS FUNCTIONS ============

function getSubProducts() {
  const sheet = getSheet(SHEETS.SUB_PRODUCTS);
  const data = sheet.getDataRange().getValues();
  const subProducts = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][1]) {
      subProducts.push({
        productSku: data[i][0],
        subProductSku: data[i][1],
        qty: Number(data[i][2]) || 1
      });
    }
  }
  
  return subProducts;
}

function syncSubProducts(subProducts) {
  const sheet = getSheet(SHEETS.SUB_PRODUCTS);
  clearSheetData(sheet);
  
  if (subProducts && subProducts.length > 0) {
    const data = subProducts.map(sp => [sp.productSku, sp.subProductSku, sp.qty]);
    sheet.getRange(2, 1, data.length, 3).setValues(data);
  }
  
  return { success: true, count: subProducts ? subProducts.length : 0 };
}

// ============ QUOTATIONS FUNCTIONS ============

function getQuotations() {
  const quotSheet = getSheet(SHEETS.QUOTATIONS);
  const itemsSheet = getSheet(SHEETS.QUOTATION_ITEMS);
  
  const quotData = quotSheet.getDataRange().getValues();
  const itemsData = itemsSheet.getDataRange().getValues();
  
  const quotations = [];
  
  for (let i = 1; i < quotData.length; i++) {
    if (quotData[i][0]) {
      const quotId = quotData[i][0];
      
      // Get items for this quotation
      const items = [];
      for (let j = 1; j < itemsData.length; j++) {
        if (itemsData[j][0] == quotId) {
          items.push({
            sku: itemsData[j][1],
            qty: Number(itemsData[j][2]) || 1,
            mrp: Number(itemsData[j][3]) || 0,
            discount: Number(itemsData[j][4]) || 0,
            netPrice: Number(itemsData[j][5]) || 0,
            total: Number(itemsData[j][6]) || 0
          });
        }
      }
      
      quotations.push({
        id: quotId,
        number: quotData[i][1],
        date: quotData[i][2],
        customer: {
          name: quotData[i][3],
          phone: quotData[i][4],
          address: quotData[i][5]
        },
        subtotal: Number(quotData[i][6]) || 0,
        gst: Number(quotData[i][7]) || 0,
        grandTotal: Number(quotData[i][8]) || 0,
        notes: quotData[i][9] || '',
        status: quotData[i][10] || 'draft',
        items: items
      });
    }
  }
  
  return quotations;
}

function syncQuotations(quotations) {
  const quotSheet = getSheet(SHEETS.QUOTATIONS);
  const itemsSheet = getSheet(SHEETS.QUOTATION_ITEMS);
  
  clearSheetData(quotSheet);
  clearSheetData(itemsSheet);
  
  if (quotations && quotations.length > 0) {
    // Sync quotation headers
    const quotData = quotations.map(q => [
      q.id,
      q.number,
      q.date,
      q.customer?.name || '',
      q.customer?.phone || '',
      q.customer?.address || '',
      q.subtotal,
      q.gst,
      q.grandTotal,
      q.notes || '',
      q.status || 'draft'
    ]);
    quotSheet.getRange(2, 1, quotData.length, 11).setValues(quotData);
    
    // Sync quotation items
    const allItems = [];
    quotations.forEach(q => {
      if (q.items && q.items.length > 0) {
        q.items.forEach(item => {
          allItems.push([
            q.id,
            item.sku,
            item.qty,
            item.mrp,
            item.discount,
            item.netPrice,
            item.total
          ]);
        });
      }
    });
    
    if (allItems.length > 0) {
      itemsSheet.getRange(2, 1, allItems.length, 7).setValues(allItems);
    }
  }
  
  return { success: true, count: quotations ? quotations.length : 0 };
}

// ============ HELPER FUNCTIONS ============

function clearSheetData(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
}

// ============ GET ALL DATA ============

function getAllData() {
  return {
    items: getItems(),
    products: getProducts(),
    bom: getBOM(),
    subProducts: getSubProducts(),
    quotations: getQuotations()
  };
}

// ============ WEB APP HANDLERS ============

function doGet(e) {
  const action = e.parameter.action || 'getAllData';
  let result;
  
  try {
    switch(action) {
      case 'getItems':
        result = { success: true, data: getItems() };
        break;
      case 'getProducts':
        result = { success: true, data: getProducts() };
        break;
      case 'getBOM':
        result = { success: true, data: getBOM() };
        break;
      case 'getSubProducts':
        result = { success: true, data: getSubProducts() };
        break;
      case 'getQuotations':
        result = { success: true, data: getQuotations() };
        break;
      case 'getAllData':
      default:
        result = { success: true, data: getAllData() };
        break;
    }
  } catch (error) {
    result = { success: false, error: error.message };
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result;
  
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    switch(action) {
      case 'syncItems':
        result = syncItems(data.items);
        break;
      case 'syncProducts':
        result = syncProducts(data.products);
        break;
      case 'syncBOM':
        result = syncBOM(data.bom);
        break;
      case 'syncSubProducts':
        result = syncSubProducts(data.subProducts);
        break;
      case 'syncQuotations':
        result = syncQuotations(data.quotations);
        break;
      case 'syncAll':
        const results = {
          items: syncItems(data.items),
          products: syncProducts(data.products),
          bom: syncBOM(data.bom),
          subProducts: syncSubProducts(data.subProducts),
          quotations: syncQuotations(data.quotations)
        };
        result = { success: true, results: results };
        break;
      default:
        result = { success: false, error: 'Unknown action: ' + action };
    }
  } catch (error) {
    result = { success: false, error: error.message };
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============ MENU & INITIALIZATION ============

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('MV Solutions')
    .addItem('Initialize All Sheets', 'initializeAllSheets')
    .addItem('View Data Summary', 'showDataSummary')
    .addToUi();
}

function initializeAllSheets() {
  Object.values(SHEETS).forEach(sheetName => {
    getSheet(sheetName);
  });
  
  SpreadsheetApp.getUi().alert(
    'Sheets initialized!\n\n' +
    'Created sheets:\n' +
    '- Items\n' +
    '- Products\n' +
    '- BOM\n' +
    '- Sub_Products\n' +
    '- Quotations\n' +
    '- Quotation_Items'
  );
}

function showDataSummary() {
  const data = getAllData();
  SpreadsheetApp.getUi().alert(
    'Data Summary\n\n' +
    'Items: ' + data.items.length + '\n' +
    'Products: ' + data.products.length + '\n' +
    'BOM Entries: ' + data.bom.length + '\n' +
    'Sub-Products: ' + data.subProducts.length + '\n' +
    'Quotations: ' + data.quotations.length
  );
}