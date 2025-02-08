const SPREADSHEET_ID = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX';
const API_KEY = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX';
const LOCATION = '池袋駅';
const BUSINESS_TYPE = 'restaurant';

function getPlacesData() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
  const places = fetchPlaces(LOCATION, BUSINESS_TYPE);
  
  // Clear the sheet before adding new data
  sheet.clear();
  
  // Set headers
  sheet.appendRow(['会社名/店舗名', 'ホームページ', '所在地', '電話番号', 'GoogleマップURL']);
  
  // Add fetched data to the sheet
  places.forEach(place => {
    sheet.appendRow([
      place.name || '',
      place.website || '',
      place.address || '',
      place.phone || '',
      place.mapsUrl || ''
    ]);
  });
}

function fetchPlaces(location, businessType) {
  const url = `https://maps.googleapis.com/maps/api/place/textsearch/json?query=${encodeURIComponent(businessType)}+in+${encodeURIComponent(location)}&key=${API_KEY}&language=ja`;
  const response = UrlFetchApp.fetch(url);
  const data = JSON.parse(response.getContentText());
  const places = data.results.map(result => {
    const details = fetchPlaceDetails(result.place_id);
    return {
      name: details.name || result.name,
      address: (details.formatted_address || result.formatted_address).replace(/^日本、/, ''),
      phone: details.formatted_phone_number || '',
      website: details.website || '',
      mapsUrl: `https://www.google.com/maps/place/?q=place_id:${result.place_id}`
    };
  });
  return places;
}

function fetchPlaceDetails(placeId) {
  const url = `https://maps.googleapis.com/maps/api/place/details/json?place_id=${placeId}&key=${API_KEY}&language=ja`;
  const response = UrlFetchApp.fetch(url);
  const data = JSON.parse(response.getContentText());
  return data.result;
}
