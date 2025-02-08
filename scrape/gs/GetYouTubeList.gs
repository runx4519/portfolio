function getYouTubeSubscribers() {
  // Replace with the name of your sheet
  const sheetName = 'シート1';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  // Check if the sheet is found
  if (!sheet) {
    return;
  }

  // Get all the URLs from the first column starting from the first row (A1)
  const urls = sheet.getRange(1, 1, sheet.getLastRow()).getValues();

  // Loop through each URL and get the channel ID
  for (let i = 0; i < urls.length; i++) {
    const url = urls[i][0];
    const channelName = getChannelNameFromURL(url);

    if (channelName) {
      const { subscribers, title } = getChannelDetails(channelName);
      sheet.getRange(i + 1, 2).setValue(subscribers);  // i + 1 to match the row in the sheet
      sheet.getRange(i + 1, 3).setValue(title);  // Record channel name in the third column
    }
  }
}

function getChannelNameFromURL(url) {
  const match = url.match(/https:\/\/www\.youtube\.com\/@([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

function getChannelDetails(channelName) {
  try {
    const searchResponse = YouTube.Search.list('id,snippet', {
      q: channelName,
      type: 'channel',
      maxResults: 1
    });

    if (searchResponse.items && searchResponse.items.length > 0) {
      const channelId = searchResponse.items[0].id.channelId;
      const title = searchResponse.items[0].snippet.title;
      const channelResponse = YouTube.Channels.list('statistics', {
        id: channelId
      });

      if (channelResponse.items && channelResponse.items.length > 0) {
        const subscribers = channelResponse.items[0].statistics.subscriberCount;
        return { subscribers, title };
      }
    }
  } catch (e) {
    // Handle error
  }
  return { subscribers: 'N/A', title: 'N/A' };
}
