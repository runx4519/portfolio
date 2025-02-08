function scrapeQoo10AndRakutenData() {
  const storeUrl = 'https://www.qoo10.jp/shop/yamada-denki';
  const sheetId = '1z0H2MuIKDstV1JOU9FO5QMyxPnlrdXl8usC273qK9p0';
  const sheetName = 'Sheet1';
  const rakutenAppId = 'XXXXXXXXXXXXXXXXXXX';
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);

  // Clear existing content
  sheet.clear();

  // Set headers
  sheet.getRange('A1').setValue('商品名');
  sheet.getRange('B1').setValue('商品URL');
  sheet.getRange('C1').setValue('価格(Qoo10)');
  sheet.getRange('D1').setValue('JANコード');
  sheet.getRange('E1').setValue('価格(楽天)');
  sheet.getRange('F1').setValue('利益');

  // Fetch the Qoo10 store page
  const response = UrlFetchApp.fetch(storeUrl);
  const html = response.getContentText();

  // Parse the HTML
  const $ = Cheerio.load(html);

  // Define selectors for product details
  const itemSelector = 'div.item';
  const productNameSelector = 'a.tt';
  const productUrlSelector = 'a.thmb';
  const productPriceSelector = 'div.prc > strong';

  // Limit the number of products to 3 for testing
  const maxProducts = 3;
  let productCount = 0;

  // Iterate over each product and extract details
  $(itemSelector).each(function (index, element) {
    if (productCount >= maxProducts) return false;

    const productName = $(element).find(productNameSelector).text().trim();
    const productUrl = $(element).find(productUrlSelector).attr('href');
    const productPriceText = $(element).find(productPriceSelector).text().trim();
    const productPrice = parseInt(productPriceText.replace('円', '').replace(/,/g, ''));

    // Fetch the product detail page to get the JAN code
    const productResponse = UrlFetchApp.fetch(productUrl);
    const productHtml = productResponse.getContentText();
    const $$ = Cheerio.load(productHtml);
    const janCode = $$('tr#tr_pan_industry td[itemprop="mpn"]').text().trim();

    // Call Rakuten API to get product price
    const rakutenApiUrl = `https://app.rakuten.co.jp/services/api/IchibaItem/Search/20170706?applicationId=${rakutenAppId}&keyword=${janCode}`;
    const rakutenResponse = UrlFetchApp.fetch(rakutenApiUrl);
    const rakutenData = JSON.parse(rakutenResponse.getContentText());
    let rakutenPrice = 0;

    if (rakutenData.Items && rakutenData.Items.length > 0) {
      rakutenPrice = rakutenData.Items[0].Item.itemPrice;
    }

    const profit = rakutenPrice - productPrice;

    // Add product details to the sheet
    sheet.getRange(productCount + 2, 1).setValue(productName);
    sheet.getRange(productCount + 2, 2).setValue(productUrl);
    sheet.getRange(productCount + 2, 3).setValue(productPrice);
    sheet.getRange(productCount + 2, 4).setValue(janCode);
    sheet.getRange(productCount + 2, 5).setValue(rakutenPrice);
    sheet.getRange(productCount + 2, 6).setValue(profit);

    productCount++;
  });
}
