function fetchEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const apiKey = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'; // 提供されたAPIキー
  const cx = 'XXXXXXXXXXXXXXXXX'; // カスタム検索エンジンID

  for (let i = 1; i <= lastRow; i++) {
    const phoneNumber = sheet.getRange(i, 1).getValue();
    if (phoneNumber) {
      const query = `${phoneNumber} メールアドレス`;
      const url = `https://www.googleapis.com/customsearch/v1?q=${encodeURIComponent(query)}&key=${apiKey}&cx=${cx}&num=3`; // num=3で検索結果の上位3件を取得
      const response = UrlFetchApp.fetch(url);
      const results = JSON.parse(response.getContentText()).items;

      if (results && results.length > 0) {
        for (let j = 0; j < results.length; j++) {
          const pageUrl = results[j].link;
          const email = getEmailFromPage(pageUrl);
          if (email !== 'No email found') {
            sheet.getRange(i, 2).setValue(email);
            break; // メールアドレスが見つかったらループを抜ける
          }
        }
      } else {
        sheet.getRange(i, 2).setValue('No result found');
      }
    }
  }
}

function getEmailFromPage(url) {
  try {
    const response = UrlFetchApp.fetch(url);
    const html = response.getContentText();
    const emailRegex = /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/i;
    const matches = html.match(emailRegex);
    return matches ? matches[0] : 'No email found';
  } catch (e) {
    console.error('Error fetching page:', e);
    return 'Error fetching page';
  }
}
