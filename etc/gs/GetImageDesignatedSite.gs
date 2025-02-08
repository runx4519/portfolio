function downloadImages() {
  // スプレッドシートを開く
  const spreadsheet = SpreadsheetApp.openById('1spy0eTE7GI8cB5G6oWKSl8o2UfxBkyZ4i9q59eeiSxc');
  const sheet = spreadsheet.getActiveSheet();
  
  // データ範囲を取得
  const lastRow = sheet.getLastRow();
  const imageUrls = sheet.getRange('A2:A' + lastRow).getValues();
  
  // Google Driveにimageフォルダがない場合は作成
  const parentFolder = DriveApp.getFileById(spreadsheet.getId()).getParents().next();
  let imageFolder;
  const folders = parentFolder.getFoldersByName('image');
  
  if (folders.hasNext()) {
    imageFolder = folders.next();
  } else {
    imageFolder = parentFolder.createFolder('image');
  }
  
  // 各URLに対して処理を実行
  imageUrls.forEach((url, index) => {
    if (url[0]) {
      try {
        // 画像をダウンロード
        const response = UrlFetchApp.fetch(url[0]);
        const imageBlob = response.getBlob();
        
        // ファイル名を生成（年月日時分秒）
        const now = new Date();
        const fileName = Utilities.formatDate(now, 'JST', 'yyyyMMddHHmmss') + '_' + index;
        
        // 画像の拡張子を取得
        const contentType = response.getHeaders()['Content-Type'];
        let extension = '.jpg'; // デフォルト拡張子
        
        if (contentType.includes('png')) {
          extension = '.png';
        } else if (contentType.includes('gif')) {
          extension = '.gif';
        } else if (contentType.includes('jpeg') || contentType.includes('jpg')) {
          extension = '.jpg';
        }
        
        // 画像を保存
        const file = imageFolder.createFile(imageBlob.setName(fileName + extension));
        
        // スプレッドシートのB列に画像IDを出力
        sheet.getRange(index + 2, 2).setValue(file.getId());
        
      } catch (error) {
        // エラーが発生した場合はログに出力
        Logger.log('Error processing URL: ' + url[0]);
        Logger.log('Error: ' + error.toString());
        // B列にエラーメッセージを出力
        sheet.getRange(index + 2, 2).setValue('Error: ' + error.toString());
      }
    }
  });
}

function onOpen() {
  // スプレッドシートにカスタムメニューを追加
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('画像ダウンロード')
    .addItem('画像を保存', 'downloadImages')
    .addToUi();
}