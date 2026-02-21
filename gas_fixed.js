// ============================================================
// BUYMA 在庫管理システム - Apps Script バックエンド v1.0
// ============================================================

var SS = SpreadsheetApp.getActiveSpreadsheet();

var SH = {
  ORDER:  '注文管理',
  STOCK:  '在庫明細',
  SALES:  '販売記録',
  CONFIG: '設定'
};

var HEADERS = {
  ORDER: ['入庫番号','注文日','ブランド','商品名','型番','仕入れ先',
          '配送業者','追跡番号','事前連絡済','関税払済','入荷日','許可書依頼済',
          '概算現地価格','通貨','概算レート','概算仕入れ額','概算関税消費税','概算原価',
          '確定現地価格','確定レート','確定仕入れ額','確定関税消費税','確定原価',
          'BUYMA単価','国内送料','BUYMA手数料','BUYMA利益','備考','登録日時'],
  STOCK: ['QR番号','入庫番号','ブランド','商品名','サイズ','カラー','数量','現在庫数','状態','登録日時'],
  SALES: ['出荷日','QR番号','入庫番号','ブランド','商品名','サイズ','カラー','販売数量','BUYMA注文番号','備考']
};

function doGet(e) {
  const SECRET_KEY = "123";

  if (!e.parameter.key || e.parameter.key !== SECRET_KEY) {
    return ContentService
      .createTextOutput("Unauthorized")
      .setMimeType(ContentService.MimeType.TEXT);
  }

  var p = e.parameter;
  var action = p.action || '';
  var result;
  try {
    switch(action) {
      case 'init':        result = initSheets(); break;
      case 'addOrder':    result = addOrder(JSON.parse(p.data)); break;
      case 'updateOrder': result = updateOrder(p.code, JSON.parse(p.data)); break;
      case 'addStock':    result = addStock(JSON.parse(p.data)); break;
      case 'getOrder':    result = getOrder(p.code); break;
      case 'getQR':       result = getQR(p.qr); break;
      case 'ship':        result = ship(p.qr, JSON.parse(p.data)); break;
      case 'getList':     result = getList(p); break;
      case 'getSummary':  result = getSummary(); break;
      default: result = {status:'error', message:'不明なアクション: '+action};
    }
  } catch(err) {
    result = {status:'error', message:err.message};
  }
  return out(result);
}

function out(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function initSheets() {
  Object.keys(SH).forEach(function(k) {
    var name = SH[k];
    var sh = SS.getSheetByName(name);
    if (!sh) {
      sh = SS.insertSheet(name);
      if (HEADERS[k]) {
        sh.getRange(1, 1, 1, HEADERS[k].length).setValues([HEADERS[k]]);
        sh.getRange(1, 1, 1, HEADERS[k].length)
          .setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
      }
    }
  });
  var cfg = SS.getSheetByName(SH.CONFIG);
  if (cfg.getRange('A1').getValue() === '') {
    cfg.getRange('A1').setValue('次の入庫番号');
    cfg.getRange('B1').setValue(405);
    cfg.getRange('A2').setValue('次のQR番号');
    cfg.getRange('B2').setValue(1);
  }
  return {status:'ok', message:'シート初期化完了'};
}

function nextOrderCode() {
  var cfg = SS.getSheetByName(SH.CONFIG);
  var n = cfg.getRange('B1').getValue() || 405;
  cfg.getRange('B1').setValue(n + 1);
  return String(n);
}

function nextQRCode(orderCode) {
  var cfg = SS.getSheetByName(SH.CONFIG);
  var n = cfg.getRange('B2').getValue() || 1;
  cfg.getRange('B2').setValue(n + 1);
  return orderCode + '-' + String(n).padStart(3, '0');
}

function addOrder(data) {
  var sh = SS.getSheetByName(SH.ORDER);
  if (!sh) initSheets();
  sh = SS.getSheetByName(SH.ORDER);

  var code = nextOrderCode();
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');

  var row = [
    code, data.orderDate || today, data.brand || '', data.name || '',
    data.itemCode || '', data.supplier || '', data.carrier || '',
    data.trackingNo || '', false, false, '', false,
    data.estPrice || 0, data.currency || 'EUR', data.estRate || 0,
    data.estPriceJPY || 0, data.estTax || 0, data.estCost || 0,
    '', '', '', '', '',
    data.buymaPrice || 0, data.domesticShip || 0, 0, 0,
    data.memo || '', now
  ];

  sh.appendRow(row);

  if (data.variants && data.variants.length > 0) {
    data.variants.forEach(function(v) {
      addStockRow(code, data.brand, data.name, v.size, v.color, v.qty);
    });
  }

  var qrNos = [];
  if (data.variants && data.variants.length > 0) {
    var stSh = SS.getSheetByName(SH.STOCK);
    var stRows = stSh.getDataRange().getValues();
    for (var i = stRows.length - 1; i >= 1; i--) {
      if (String(stRows[i][1]) === String(code)) {
        qrNos.unshift(stRows[i][0]);
        if (qrNos.length >= data.variants.length) break;
      }
    }
  }
  return {status:'ok', code:code, qrNos:qrNos, message:'入庫番号 '+code+' で登録しました'};
}

function updateOrder(code, data) {
  var sh = SS.getSheetByName(SH.ORDER);
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(code)) {
      Object.keys(data).forEach(function(key) {
        var colMap = {
          orderDate:1, brand:2, name:3, itemCode:4, supplier:5,
          carrier:6, trackingNo:7, contacted:8, taxPaid:9,
          arrivalDate:10, permitRequested:11,
          estPrice:12, currency:13, estRate:14, estPriceJPY:15, estTax:16, estCost:17,
          fixedPrice:18, fixedRate:19, fixedPriceJPY:20, fixedTax:21, fixedCost:22,
          buymaPrice:23, domesticShip:24, memo:27
        };
        if (colMap[key] !== undefined) {
          sh.getRange(i+1, colMap[key]+1).setValue(data[key]);
        }
      });
      var buymaPrice = sh.getRange(i+1, 24).getValue();
      var fee = buymaPrice * 0.077;
      var cost = sh.getRange(i+1, 23).getValue() || sh.getRange(i+1, 18).getValue();
      var domShip = sh.getRange(i+1, 25).getValue() || 0;
      sh.getRange(i+1, 26).setValue(Math.round(fee));
      sh.getRange(i+1, 27).setValue(Math.round(buymaPrice - fee - cost - domShip));
      return {status:'ok', message:'更新しました'};
    }
  }
  return {status:'error', message:'入庫番号 '+code+' が見つかりません'};
}

function addStockRow(orderCode, brand, name, size, color, qty) {
  var sh = SS.getSheetByName(SH.STOCK);
  if (!sh) initSheets();
  sh = SS.getSheetByName(SH.STOCK);
  var qr = nextQRCode(orderCode);
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  sh.appendRow([qr, orderCode, brand, name, size||'', color||'', qty||1, qty||1, '在庫中', now]);
  return qr;
}

function addStock(data) {
  var qr = addStockRow(data.orderCode, data.brand, data.name, data.size, data.color, data.qty);
  return {status:'ok', qr:qr, message:'QR番号 '+qr+' で登録しました'};
}

function getOrder(code) {
  var sh = SS.getSheetByName(SH.ORDER);
  var rows = sh.getDataRange().getValues();
  var headers = rows[0];
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(code)) {
      var obj = {};
      headers.forEach(function(h,j){ obj[h] = rows[i][j]; });
      obj.variants = getVariants(code);
      return {status:'ok', data:obj};
    }
  }
  return {status:'error', message:'見つかりません'};
}

function getVariants(orderCode) {
  var sh = SS.getSheetByName(SH.STOCK);
  if (!sh) return [];
  var rows = sh.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) === String(orderCode)) {
      result.push({
        qr: rows[i][0], orderCode: rows[i][1],
        brand: rows[i][2], name: rows[i][3],
        size: rows[i][4], color: rows[i][5],
        qty: rows[i][6], currentQty: rows[i][7],
        status: rows[i][8]
      });
    }
  }
  return result;
}

function getQR(qr) {
  var sh = SS.getSheetByName(SH.STOCK);
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(qr)) {
      var item = {
        qr: rows[i][0], orderCode: rows[i][1],
        brand: rows[i][2], name: rows[i][3],
        size: rows[i][4], color: rows[i][5],
        qty: rows[i][6], currentQty: rows[i][7],
        status: rows[i][8]
      };
      var order = getOrder(rows[i][1]);
      if (order.status === 'ok') item.order = order.data;
      return {status:'ok', data:item};
    }
  }
  return {status:'error', message:'QR番号 '+qr+' が見つかりません'};
}

function ship(qr, data) {
  var sh = SS.getSheetByName(SH.STOCK);
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(qr)) {
      var currentQty = rows[i][7];
      var shipQty = data.qty || 1;
      if (currentQty < shipQty) {
        return {status:'error', message:'在庫数が不足しています（現在庫：'+currentQty+'）'};
      }
      var newQty = currentQty - shipQty;
      sh.getRange(i+1, 8).setValue(newQty);
      sh.getRange(i+1, 9).setValue(newQty <= 0 ? '出荷済' : '在庫中');
      var salesSh = SS.getSheetByName(SH.SALES);
      var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
      salesSh.appendRow([
        data.shipDate || today, qr, rows[i][1], rows[i][2], rows[i][3],
        rows[i][4], rows[i][5], shipQty, data.buymaOrderNo || '', data.memo || ''
      ]);
      return {status:'ok', message:'出荷処理完了', remaining:newQty};
    }
  }
  return {status:'error', message:'QR番号が見つかりません'};
}

function getList(p) {
  var sh = SS.getSheetByName(SH.STOCK);
  if (!sh) return {status:'ok', data:[]};
  var rows = sh.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (p.all === 'true' || rows[i][8] === '在庫中') {
      if (rows[i][7] > 0 || p.all === 'true') {
        result.push({
          qr: rows[i][0], orderCode: rows[i][1],
          brand: rows[i][2], name: rows[i][3],
          size: rows[i][4], color: rows[i][5],
          qty: rows[i][6], currentQty: rows[i][7],
          status: rows[i][8]
        });
      }
    }
  }
  if (p.grouped === 'true') {
    var grouped = {};
    result.forEach(function(r) {
      var key = r.orderCode + '_' + r.brand + '_' + r.name;
      if (!grouped[key]) grouped[key] = {orderCode:r.orderCode, brand:r.brand, name:r.name, variants:[]};
      grouped[key].variants.push(r);
    });
    return {status:'ok', data:Object.values(grouped)};
  }
  return {status:'ok', data:result};
}

function getSummary() {
  var orderSh = SS.getSheetByName(SH.ORDER);
  var stockSh = SS.getSheetByName(SH.STOCK);
  if (!orderSh || !stockSh) return {status:'ok', data:{}};
  var orders = orderSh.getDataRange().getValues();
  var stocks = stockSh.getDataRange().getValues();
  var totalStock = 0, totalCost = 0, inTransit = 0;
  for (var i = 1; i < stocks.length; i++) {
    if (stocks[i][8] === '在庫中') totalStock += stocks[i][7];
  }
  for (var j = 1; j < orders.length; j++) {
    var cost = orders[j][22] || orders[j][17] || 0;
    totalCost += Number(cost);
    if (!orders[j][10]) inTransit++;
  }
  return {status:'ok', data:{
    totalStock: totalStock,
    totalCost: Math.round(totalCost),
    inTransit: inTransit
  }};
}
