function getSuggestKeywords() {
  deleteTriggers();
  
  const BATCH_SIZE = 50;
  const TIME_LIMIT = 5.5 * 60 * 1000;
  const START_ROW = 2;
  const startTime = new Date();
  
  try {
    // スプレッドシートの取得方法を修正
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error('スプレッドシートが開けませんでした。');
    }
    
    const sheet = spreadsheet.getActiveSheet();
    if (!sheet) {
      throw new Error('シートが見つかりませんでした。');
    }
    
    // 初回実行かどうかを確認
    const currentStatus = sheet.getRange('G2').getValue();
    const currentPosition = parseInt(sheet.getRange('J2').getValue()?.split('/')[0] || 1);
    const isFirstRun = !currentStatus || currentStatus === '完了';
    
    // フォルダとログファイルの準備
    const spreadsheetFolder = DriveApp.getFileById(spreadsheet.getId()).getParents().next();
    const outputFolder = createOrGetFolder(spreadsheetFolder, 'output');
    const logFolder = createOrGetFolder(spreadsheetFolder, 'log');
    const logFileName = Utilities.formatDate(startTime, 'Asia/Tokyo', 'yyyyMMdd') + '.txt';
    let logFile = createOrGetLogFile(logFolder, logFileName);
    
    appendToLog(logFile, '新規処理開始');
    
    if (isFirstRun) {
      sheet.getRange('G1:J1').setValues([['処理状態', '処理開始時刻', '処理終了時刻', '進捗状況']]);
      sheet.getRange('G2:J2').clearContent();
      const initialTime = Utilities.formatDate(startTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
      sheet.getRange('G2:H2').setValues([['実行中', initialTime]]);
    } else {
      sheet.getRange('G2').setValue('実行中');
      sheet.getRange('I2').clearContent();
    }
    
    // 出力ファイル名を生成
    let outputFileName = isFirstRun
      ? Utilities.formatDate(startTime, 'Asia/Tokyo', 'yyyyMMddHHmm') + '.txt'
      : Utilities.formatDate(new Date(sheet.getRange('H2').getValue()), 'Asia/Tokyo', 'yyyyMMddHHmm') + '.txt';
    
    let outputFile = isFirstRun
      ? outputFolder.createFile(outputFileName, '', MimeType.PLAIN_TEXT)
      : outputFolder.getFilesByName(outputFileName).next();
    
    // データ範囲の取得と処理
    const lastRow = sheet.getLastRow();
    const totalKeywords = sheet.getRange(`A${START_ROW}:A${lastRow}`)
      .getValues()
      .filter(k => k[0].toString().trim() !== '').length;
    
    let currentRow = isFirstRun ? START_ROW : parseInt(sheet.getRange('J2').getValue().split('/')[0]) + 1;
    
    while (currentRow <= lastRow) {
      if (new Date() - startTime > TIME_LIMIT) {
        handleTimeout(sheet, currentRow, totalKeywords, logFile);
        return;
      }
      
      const keyword = sheet.getRange(`A${currentRow}`).getValue().toString().trim();
      
      if (keyword) {
        try {
          sheet.getRange('J2').setValue(`${currentRow - 1}/${totalKeywords}`);
          appendToLog(logFile, `キーワード「${keyword}」の処理開始`);
          
          const processedKeyword = preprocessKeyword(keyword);
          let suggestions = [];

          // まず完全なキーワードで試行
          suggestions = fetchSuggestions(processedKeyword);

          // サジェストが得られなかった場合、キーワードを分割して試行
          if (suggestions.length === 0) {
            const words = processedKeyword.split(/\s+/);
            if (words.length > 1) {
              // 最初の単語で試行
              suggestions = fetchSuggestions(words[0]);
            }
          }

          if (suggestions.length > 0) {
            updateOutputFile(outputFile, keyword, suggestions);
            appendToLog(logFile, `キーワード「${keyword}」の処理完了: ${suggestions.length}件のサジェスト取得`);
          } else {
            appendToLog(logFile, `警告: キーワード「${keyword}」のサジェストは0件でした`);
          }
          
        } catch (error) {
          appendToLog(logFile, `エラー: キーワード「${keyword}」の処理中にエラーが発生: ${error.toString()}`);
        }
        
        Utilities.sleep(200); // APIレート制限のための最小限の待機時間
      }
      currentRow++;
    }
    
    completeProcessing(sheet, totalKeywords, logFile);
    
  } catch (error) {
    handleError(sheet, logFile, error);
    throw error;
  }
}

function fetchSuggestions(keyword) {
  const url = `http://suggestqueries.google.com/complete/search?output=toolbar&hl=ja&gl=JP&psi=1&q=${encodeURIComponent(keyword)}`;
  
  const options = {
    'method': 'get',
    'headers': {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
      'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
      'Accept-Language': 'ja,en-US;q=0.7,en;q=0.3',
    },
    'muteHttpExceptions': true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const xml = response.getContentText();
    return parseSuggestResponse(xml);
  } catch (error) {
    console.error('APIリクエストエラー:', error);
    return [];
  }
}

function preprocessKeyword(keyword) {
  return keyword
    .replace(/\s+/g, ' ') // 連続する空白を1つに
    .trim();              // 前後の空白を削除
}

function parseSuggestResponse(xml) {
  try {
    const document = XmlService.parse(xml);
    const root = document.getRootElement();
    const suggestions = [];
    
    const suggestionElements = root.getChildren('CompleteSuggestion');
    if (suggestionElements && suggestionElements.length > 0) {
      // 最大10件まで処理
      for (let i = 0; i < Math.min(10, suggestionElements.length); i++) {
        const element = suggestionElements[i];
        const suggestion = element.getChild('suggestion');
        if (suggestion) {
          const data = suggestion.getAttribute('data');
          if (data) {
            suggestions.push(data.getValue());
          }
        }
      }
    }
    
    return suggestions;
  } catch (error) {
    console.error('XML解析エラー:', error);
    return [];
  }
}

function updateOutputFile(outputFile, keyword, suggestions) {
  const currentContent = outputFile.getBlob().getDataAsString();
  const newContent = currentContent + 
    `\n【${keyword}】\n` + 
    suggestions.join('\n') + 
    '\n----------------------------------------\n';
  outputFile.setContent(newContent);
}

function handleTimeout(sheet, currentRow, totalKeywords, logFile) {
  sheet.getRange('G2').setValue('一時停止');
  sheet.getRange('J2').setValue(`${currentRow - 1}/${totalKeywords}`);
  
  // トリガーの作成方法を変更
  ScriptApp.newTrigger('getSuggestKeywords')
    .timeBased()
    .after(1 * 60 * 1000)  // 1分後
    .create();
  
  appendToLog(logFile, '実行時間制限により一時停止。1分後に自動再開します。');
}

function completeProcessing(sheet, totalKeywords, logFile) {
  sheet.getRange('G2').setValue('完了');
  sheet.getRange('I2').setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
  sheet.getRange('J2').setValue(`${totalKeywords}/${totalKeywords}`);
  
  appendToLog(logFile, '全ての処理が完了しました。', false);
}

function handleError(sheet, logFile, error) {
  sheet.getRange('G2').setValue('エラー');
  appendToLog(logFile, `エラーが発生しました: ${error.toString()}`);
}

function createOrGetFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parentFolder.createFolder(folderName);
  }
}

function createOrGetLogFile(logFolder, logFileName) {
  const existingLogFiles = logFolder.getFilesByName(logFileName);
  if (existingLogFiles.hasNext()) {
    return existingLogFiles.next();
  } else {
    return logFolder.createFile(logFileName, `処理開始\n`, MimeType.PLAIN_TEXT);
  }
}

function appendToLog(logFile, message, addTimestamp = true) {
  const currentContent = logFile.getBlob().getDataAsString();
  let newContent;
  
  if (addTimestamp) {
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    newContent = currentContent + `[${timestamp}] ${message}\n`;
  } else {
    newContent = currentContent + `${message}\n`;
  }
  
  logFile.setContent(newContent);
}

function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    ScriptApp.deleteTrigger(trigger);
  }
}