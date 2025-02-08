function fetchBitcoinPrice() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var url = 'https://api.bitflyer.com/v1/ticker';
  
  // bitFlyer APIからビットコインの現在価格を取得
  var response = UrlFetchApp.fetch(url);
  var data = JSON.parse(response.getContentText());
  var bitcoinPrice = data.ltp; // 最新取引価格（LTP: Last Traded Price）

  // 現在の日時を取得
  var now = new Date();

  // bitFlyer APIからアカウントの残高を取得
  var balance = fetchBalance();
  var btcBalance = balance.btc;
  var jpyBalance = balance.jpy;
  var totalBalance = btcBalance * bitcoinPrice + jpyBalance;
  
  // 注文の実行と記録
  var orderType = '';
  var orderQuantity = 0.001;
  
  if (bitcoinPrice <= 10500000 && jpyBalance >= orderQuantity * bitcoinPrice) {
    orderType = 'Buy';
    executeOrder('BUY', orderQuantity);
  } else if (bitcoinPrice >= 11000000 && btcBalance >= orderQuantity) {
    orderType = 'Sell';
    executeOrder('SELL', orderQuantity);
  }

  // スプレッドシートにデータを書き込む
  sheet.appendRow([now, bitcoinPrice, btcBalance, jpyBalance, totalBalance, orderType, orderQuantity]);
}

function fetchBalance() {
  var apiKey = 'XXXXXXXXXXXXXXXXXXXXXX';
  var apiSecret = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX';
  
  var timestamp = Math.floor(new Date().getTime() / 1000);
  var method = 'GET';
  var path = '/v1/me/getbalance';
  var text = timestamp + method + path;
  var signature = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, text, apiSecret);
  var headers = {
    'ACCESS-KEY': apiKey,
    'ACCESS-TIMESTAMP': timestamp,
    'ACCESS-SIGN': Utilities.base64Encode(signature)
  };

  var options = {
    'method': method,
    'headers': headers,
    'muteHttpExceptions': true
  };

  var response = UrlFetchApp.fetch('https://api.bitflyer.com' + path, options);
  var data = JSON.parse(response.getContentText());

  var btcBalance = 0;
  var jpyBalance = 0;
  for (var i = 0; i < data.length; i++) {
    if (data[i].currency_code === 'BTC') {
      btcBalance = data[i].amount;
    }
    if (data[i].currency_code === 'JPY') {
      jpyBalance = data[i].amount;
    }
  }

  return {
    btc: btcBalance,
    jpy: jpyBalance
  };
}

function executeOrder(side, size) {
  var apiKey = '5YjvMN84ajPEyY6uNTC5H4';
  var apiSecret = 'V2FJ1v5mbBCcE3DsgWxPEy7K/77/PZ8tCVF3h3em24U=';
  
  var timestamp = Math.floor(new Date().getTime() / 1000);
  var method = 'POST';
  var path = '/v1/me/sendchildorder';
  var body = {
    'product_code': 'BTC_JPY',
    'child_order_type': 'MARKET',
    'side': side,
    'size': size
  };
  var text = timestamp + method + path + JSON.stringify(body);
  var signature = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, text, apiSecret);
  var headers = {
    'ACCESS-KEY': apiKey,
    'ACCESS-TIMESTAMP': timestamp,
    'ACCESS-SIGN': Utilities.base64Encode(signature),
    'Content-Type': 'application/json'
  };

  var options = {
    'method': method,
    'headers': headers,
    'payload': JSON.stringify(body),
    'muteHttpExceptions': true
  };

  var response = UrlFetchApp.fetch('https://api.bitflyer.com' + path, options);
  var data = JSON.parse(response.getContentText());
  
  return data;
}

function createTimeDrivenTriggers() {
  // 既存のトリガーを削除
  deleteTriggers();
  
  // 1分ごとのトリガーを設定
  ScriptApp.newTrigger('fetchBitcoinPrice')
           .timeBased()
           .everyMinutes(1)
           .create();
}

function deleteTriggers() {
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
}
