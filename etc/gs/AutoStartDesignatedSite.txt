function onOpen() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const searchTerm = sheet.getRange('A1').getValue();
  
  if (!searchTerm) {
    SpreadsheetApp.getUi().alert('A1セルに検索したい語句を入力してください。');
    return;
  }
  
  const baseUrl = 'https://auctions.yahoo.co.jp/search/search';
  const encodedTerm = encodeURIComponent(searchTerm);
  const searchUrl = `${baseUrl}?fr=auc_top&p=${encodedTerm}`;
  
  const html = HtmlService
    .createHtmlOutput(`
      <!DOCTYPE html>
      <html>
        <body>
          <script>
            window.open('${searchUrl}', '_blank');
            google.script.host.close();
          </script>
        </body>
      </html>
    `)
    .setWidth(1)
    .setHeight(1);
  
  SpreadsheetApp.getUi().showModelessDialog(html, '検索中...');
}