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
          '現地価格','通貨','概算レート','概算仕入れ額','概算関税消費税','概算原価',
          '確定レート','確定仕入れ額','確定関税消費税','確定原価',
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
      case 'getTransit':  result = getTransit(); break;
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
    data.price || 0, data.currency || 'EUR', data.estRate || 0,
    data.priceJPY || 0, data.estTax || 0, data.estCost || 0,
    '', '', '', '',
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
          price:12, currency:13, estRate:14, priceJPY:15, estTax:16, estCost:17,
          fixedRate:18, fixedPriceJPY:19, fixedTax:20, fixedCost:21,
          buymaPrice:22, domesticShip:23, memo:26
        };
        if (colMap[key] !== undefined) {
          sh.getRange(i+1, colMap[key]+1).setValue(data[key]);
        }
      });
      var buymaPrice = sh.getRange(i+1, 23).getValue();
      var fee = buymaPrice * 0.077;
      var fixedCost = Number(sh.getRange(i+1, 22).getValue()) || 0;
      var estCost = Number(sh.getRange(i+1, 18).getValue()) || 0;
      var cost = fixedCost || estCost;
      var domShip = sh.getRange(i+1, 24).getValue() || 0;
      sh.getRange(i+1, 25).setValue(Math.round(fee));
      sh.getRange(i+1, 26).setValue(Math.round(buymaPrice - fee - cost - domShip));
      // 入荷日が入力された場合、在庫明細の状態を「在庫中」に更新
      if (data.arrivalDate) {
        var stockSh = SS.getSheetByName(SH.STOCK);
        if (stockSh) {
          var stockRows = stockSh.getDataRange().getValues();
          for (var s = 1; s < stockRows.length; s++) {
            if (String(stockRows[s][1]) === String(code)) {
              stockSh.getRange(s+1, 9).setValue('在庫中');
            }
          }
        }
      }
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
      headers.forEach(function(h,j){
        var v = rows[i][j];
        if (v instanceof Date) {
          obj[h] = Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy/MM/dd');
        } else if (typeof v === 'string' && v.includes('T') && v.includes('Z')) {
          obj[h] = v.substring(0, 10).replace(/-/g, '/');
        } else {
          obj[h] = v;
        }
      });
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
      sh.getRange(i+1, 9).setValue(newQty <= 0 ? '出荷済み' : '在庫中');
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

  // 注文管理から入荷日をマップ化
  var orderSh = SS.getSheetByName(SH.ORDER);
  var arrivalMap = {};
  if (orderSh) {
    var orderRows = orderSh.getDataRange().getValues();
    for (var j = 1; j < orderRows.length; j++) {
      var code = String(orderRows[j][0]);
      var arrival = orderRows[j][10]; // 入荷日は11列目（index10）
      arrivalMap[code] = arrival;
    }
  }

  var rows = sh.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (p.all === 'true' || rows[i][8] === '在庫中') {
      if (rows[i][7] > 0 || p.all === 'true') {
        var orderCode = String(rows[i][1]);
        result.push({
          qr: rows[i][0], orderCode: orderCode,
          brand: rows[i][2], name: rows[i][3],
          size: rows[i][4], color: rows[i][5],
          qty: rows[i][6], currentQty: rows[i][7],
          status: rows[i][8],
          arrivalDate: arrivalMap[orderCode] || ''
        });
      }
    }
  }
  if (p.grouped === 'true') {
    var grouped = {};
    result.forEach(function(r) {
      var key = r.orderCode + '_' + r.brand + '_' + r.name;
      if (!grouped[key]) grouped[key] = {
        orderCode: r.orderCode, brand: r.brand, name: r.name,
        arrivalDate: r.arrivalDate, variants: []
      };
      grouped[key].variants.push(r);
    });
    return {status:'ok', data:Object.values(grouped)};
  }
  return {status:'ok', data:result};
}

function getTransit() {
  var sh = SS.getSheetByName(SH.ORDER);
  if (!sh) return {status:'ok', data:[]};
  var rows = sh.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0] || String(rows[i][0]).trim() === '') continue;
    var arrival = rows[i][10];
    if (!arrival || String(arrival).trim() === '') {
      result.push({
        orderCode: String(rows[i][0]),
        brand:     rows[i][2],
        name:      rows[i][3],
        supplier:  rows[i][5],
        carrier:   rows[i][6],
        trackingNo: rows[i][7]
      });
    }
  }
  return {status:'ok', data:result};
}

function getSummary() {
  var orderSh = SS.getSheetByName(SH.ORDER);
  var stockSh = SS.getSheetByName(SH.STOCK);
  if (!orderSh || !stockSh) return {status:'ok', data:{}};
  var orders = orderSh.getDataRange().getValues();
  var stocks = stockSh.getDataRange().getValues();

  // 注文管理から原価マップを作成（確定原価優先、なければ概算原価）
  var costMap = {};
  for (var j = 1; j < orders.length; j++) {
    if (!orders[j][0] || String(orders[j][0]).trim() === '') continue;
    var code = String(orders[j][0]);
    var fixedCost = Number(orders[j][21]) || 0;
    var estCost   = Number(orders[j][17]) || 0;
    costMap[code] = fixedCost || estCost;
  }

  var totalStock = 0, inTransit = 0, stockValue = 0;
  for (var i = 1; i < stocks.length; i++) {
    if (stocks[i][8] === '在庫中') {
      var qty = Number(stocks[i][7]) || 0;
      var oCode = String(stocks[i][1]);
      totalStock += qty;
      stockValue += qty * (costMap[oCode] || 0);
    }
  }
  for (var k = 1; k < orders.length; k++) {
    if (!orders[k][0] || String(orders[k][0]).trim() === '') continue;
    if (!orders[k][10]) inTransit++;
  }

  // 本日出荷数
  var salesSh = SS.getSheetByName(SH.SALES);
  var todayShip = 0;
  if (salesSh) {
    var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
    var sales = salesSh.getDataRange().getValues();
    for (var s = 1; s < sales.length; s++) {
      var shipDate = sales[s][0] instanceof Date
        ? Utilities.formatDate(sales[s][0], 'Asia/Tokyo', 'yyyy/MM/dd')
        : String(sales[s][0]).substring(0,10);
      if (shipDate === today) todayShip += Number(sales[s][7]) || 1;
    }
  }

  // 輸送中リスト（ホーム画面用）
  var transitItems = [];
  for (var t = 1; t < orders.length; t++) {
    if (!orders[t][0] || String(orders[t][0]).trim() === '') continue;
    if (!orders[t][10]) {
      transitItems.push({
        orderCode: String(orders[t][0]),
        brand: orders[t][2],
        name: orders[t][3],
        supplier: orders[t][5]
      });
    }
  }

  return {status:'ok', data:{
    totalStock: totalStock,
    stockValue: Math.round(stockValue),
    inTransit: inTransit,
    todayShip: todayShip,
    transitItems: transitItems
  }};
}
