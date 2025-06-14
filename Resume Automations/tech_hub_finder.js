/**
 * Job Location Tech Hub Enhancer
 * Automatically appends nearby tech hub locations to job postings
 * Uses Gemini 1.5 Flash API for intelligent location mapping
 */

// Configuration constants
const CONFIG = {
  LOCATION_COLUMN: 'G',
  OUTPUT_COLUMN: 'K',
  START_ROW: 2,
  BATCH_SIZE: 15, // Reduced to respect 15 requests per minute limit
  API_DELAY: 1000, // 1 seconds between API calls (4 requests per minute max)
  DEFAULT_LOCATION: 'Los Angeles, CA',
  GEMINI_API_URL: 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent'
};

// Remote/generic location keywords that should default to Los Angeles
const REMOTE_KEYWORDS = [
  "remote", "usa", "united states", "remote, usa", "remote, united states"
];

// Predefined city mappings to avoid API calls for common locations
const CITY_MAPPINGS = {
  'san jose, ca': 'Mountain View, CA',
  'san francisco, ca': 'Mountain View, CA',
  'san francisco bay area': 'Mountain View, CA',
  'mountain view, ca': 'San Jose, CA',
  'san diego, ca': 'Los Angeles, CA',
  'los angeles, ca': 'Los Angeles, CA',
  'austin, tx': 'San Antonio, TX',
  'san antonio, tx': 'Austin, TX',
  'houston, tx': 'San Antonio, TX',
  'dallas, tx': 'Fort Worth, TX',
  'fort worth, tx': 'Dallas, TX',
  'frisco, tx': 'Plano, TX',
  'seattle, wa': 'Tacoma, WA',
  'tacoma, wa': 'Seattle, WA',
  'bellevue, wa': 'Tacoma, WA',
  'new york city, ny': 'Newark, NJ',
  'new york, ny': 'Newark, NJ',
  'newark, nj': 'New York, NY',
  'cambridge, ma': 'Worcester, MA',
  'worcester, ma': 'Cambridge, MA',
  'chicago, il': 'Naperville, IL',
  'naperville, il': 'Chicago, IL',
  'atlanta, ga': 'Alpharetta, GA',
  'alpharetta, ga': 'Atlanta, GA',
  'birmingham, nc': 'Alpharetta, GA',
  'raleigh, nc': 'Durham, NC',
  'charlotte, nc': 'Durham, NC',
  'durham, nc': 'Raleigh, NC',
  'denver, co': 'Boulder, CO',
  'boulder, co': 'Denver, CO',
  'arlington, va': 'Alexandria, VA',
  'alexandria, va': 'Arlington, VA',
  'miami, fl': 'Fort Lauderdale, FL',
  'tampa, fl': 'Fort Lauderdale, FL',
  'fort lauderdale, fl': 'Miami, FL',
  'philadelphia, pa': 'Wilmington, DE',
  'wilmington, de': 'Philadelphia, PA',
  'phoenix, az': 'Scottsdale, AZ',
  'tempe, az': 'Scottsdale, AZ',
  'scottsdale, az': 'Phoenix, AZ',
  'columbus, oh': 'Dayton, OH',
  'dayton, oh': 'Columbus, OH',
  'cleveland, oh': 'Lakewood, OH',
  'lakewood, oh': 'Cleveland, OH',
  'blue ash, oh': 'Cincinnati, OH',
  'cincinnati, oh': 'Dayton, OH',
  'mason, oh': 'Cincinnati, OH',
  'detroit, mi': 'Ann Arbor, MI',
  'ann arbor, mi': 'Detroit, MI',
  'minneapolis, mn': 'St. Paul, MN',
  'st. paul, mn': 'Minneapolis, MN',
  'portland, or': 'Beaverton, OR',
  'beaverton, or': 'Portland, OR',
  'salt lake city, ut': 'Provo, UT',
  'provo, ut': 'Salt Lake City, UT',
  'odgen, ut': 'Salt Lake City, UT',
  'draper, ut': 'Salt Lake City, UT',
  'st. louis, mo': 'Clayton, MO',
  'clayton, mo': 'St. Louis, MO',
  'nashville, tn': 'Murfreesboro, TN',
  'murfreesboro, tn': 'Nashville, TN',
  'indianapolis, in': 'Carmel, IN',
  'carmel, in': 'Indianapolis, IN',
  'madison, wi': 'Milwaukee, WI',
  'milwaukee, wi': 'Madison, WI',
  'huntsville, al': 'Decatur, AL',
  'decatur, al': 'Huntsville, AL',
  'new orleans, la': 'Baton Rouge, LA',
  'baton rouge, la': 'New Orleans, LA',
  'charleston, sc': 'Mount Pleasant, SC',
  'mount pleasant, sc': 'Charleston, SC',
  'las vegas, nv': 'Henderson, NV',
  'henderson, nv': 'Las Vegas, NV',
  'lexington, ky': 'Louisville, KY',
  'louisville, ky': 'Lexington, KY',
  'oklahoma city, ok': 'Norman, OK',
  'norman, ok': 'Oklahoma City, OK',
  'des moines, ia': 'Ames, IA',
  'ames, ia': 'Des Moines, IA',
  'kansas city, ks': 'Overland Park, KS',
  'overland park, ks': 'Kansas City, KS',
  'little rock, ar': 'Conway, AR',
  'conway, ar': 'Little Rock, AR',
  'albuquerque, nm': 'Santa Fe, NM',
  'santa fe, nm': 'Albuquerque, NM',
  'omaha, ne': 'Lincoln, NE',
  'lincoln, ne': 'Omaha, NE',
  'boise, id': 'Meridian, ID',
  'meridian, id': 'Boise, ID',
  'jackson, ms': 'Madison, MS',
  'madison, ms': 'Jackson, MS',
  'morgantown, wv': 'Fairmont, WV',
  'fairmont, wv': 'Morgantown, WV',
  'portland, me': 'Lewiston, ME',
  'lewiston, me': 'Portland, ME',
  'manchester, nh': 'Nashua, NH',
  'nashua, nh': 'Manchester, NH',
  'burlington, vt': 'Montpelier, VT',
  'montpelier, vt': 'Burlington, VT',
  'providence, ri': 'Warwick, RI',
  'warwick, ri': 'Providence, RI',
  'newark, de': 'Wilmington, DE',
  'anchorage, ak': 'Wasilla, AK',
  'wasilla, ak': 'Anchorage, AK',
  'honolulu, hi': 'Kailua, HI',
  'kailua, hi': 'Honolulu, HI',
  'irvine, ca': 'Los Angeles, CA',
  'los angeles, ca': 'Irvine, CA',
  'lehi, ut': 'Salt Lake City, UT',
  'baltimore, md': 'Washington, DC',
  'washington, dc': 'Baltimore, MD',
  'rockville, md': 'Washington, DC',
  'north bethesda, md': 'North Bethesda, MD',
  'warren, mi': 'Detroit, MI',
  'jersey city, nj': 'New York, NY',
  'fort mill, sc': 'Charlotte, NC',
  'somerville, ma': 'Boston, MA',
  'boston, ma': 'Somerville, MA',
  'madison, wi': 'Milwaukee, WI',
  'hartford, ct': 'New Haven, CT',
  'new haven, ct': 'Hartford, CT',
  'mclean, va': 'Arlington, VA',
  'reston, va': 'Tysons Corner, VA',
  'palo alto, ca': 'San Jose, CA',
  'district of columbia, united states': 'Arlington, VA',
  'columbia, sc': 'Greenville, SC',
  'tulsa, ok': 'Broken Arrow, OK'
};

/**
 * Main function to enhance job locations with nearby tech hubs
 */
function enhanceJobLocations() {
  try {
    Logger.log('Starting job location enhancement...');
    
    const sheet = SpreadsheetApp.getActiveSheet();
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    
    const dataToProcess = getRowsToProcess(sheet);
    
    if (dataToProcess.length === 0) {
      Logger.log('No rows found that need processing.');
      return;
    }
    
    Logger.log(`Found ${dataToProcess.length} locations to process`);
    
    // Process in batches
    const batchedData = batchArray(dataToProcess, CONFIG.BATCH_SIZE);
    let totalProcessed = 0;
    
    for (let batchIndex = 0; batchIndex < batchedData.length; batchIndex++) {
      const batch = batchedData[batchIndex];
      Logger.log(`Processing batch ${batchIndex + 1}/${batchedData.length} (${batch.length} items)`);
      
      const results = processBatch(batch, apiKey);
      updateSheet(sheet, results);
      
      totalProcessed += batch.length;
      
      // Add delay between batches (except for the last batch)
      if (batchIndex < batchedData.length - 1) {
        Utilities.sleep(1000);
      }
    }
    
    Logger.log(`Enhancement complete! Processed ${totalProcessed} locations.`);
    
  } catch (error) {
    Logger.log(`Error in enhanceJobLocations: ${error.toString()}`);
  }
}

/**
 * Get rows that need processing (non-empty location, empty output)
 */
function getRowsToProcess(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.START_ROW) {
    return [];
  }
  
  const locationRange = sheet.getRange(`${CONFIG.LOCATION_COLUMN}${CONFIG.START_ROW}:${CONFIG.LOCATION_COLUMN}${lastRow}`);
  const outputRange = sheet.getRange(`${CONFIG.OUTPUT_COLUMN}${CONFIG.START_ROW}:${CONFIG.OUTPUT_COLUMN}${lastRow}`);
  
  const locations = locationRange.getValues().flat();
  const outputs = outputRange.getValues().flat();
  
  const dataToProcess = [];
  
  for (let i = 0; i < locations.length; i++) {
    const location = locations[i];
    const output = outputs[i];
    const rowNumber = CONFIG.START_ROW + i;
    
    if (location && location.toString().trim() !== '' && 
        (!output || output.toString().trim() === '')) {
      dataToProcess.push({
        location: location.toString().trim(),
        rowNumber: rowNumber
      });
    }
  }
  
  return dataToProcess;
}

/**
 * Split array into batches
 */
function batchArray(array, batchSize) {
  const batches = [];
  for (let i = 0; i < array.length; i += batchSize) {
    batches.push(array.slice(i, i + batchSize));
  }
  return batches;
}

/**
 * Process a batch of locations
 */
function processBatch(batch, apiKey) {
  const results = [];
  
  for (let i = 0; i < batch.length; i++) {
    const item = batch[i];
    let techHub;
    
    try {
      if (shouldDefaultToLA(item.location)) {
        techHub = CONFIG.DEFAULT_LOCATION;
        Logger.log(`Row ${item.rowNumber}: "${item.location}" -> Default: ${techHub}`);
      } else {
        // First check predefined mappings
        const mappedCity = checkCityMappings(item.location);
        if (mappedCity) {
          techHub = mappedCity;
          Logger.log(`Row ${item.rowNumber}: "${item.location}" -> Mapped: ${techHub}`);
        } else {
          // Use Gemini API as fallback
          techHub = getTechHubFromGemini(item.location, apiKey);
          Logger.log(`Row ${item.rowNumber}: "${item.location}" -> API: ${techHub}`);
        }
      }
    } catch (error) {
      Logger.log(`Error processing row ${item.rowNumber} (${item.location}): ${error.toString()}`);
      techHub = CONFIG.DEFAULT_LOCATION;
    }
    
    results.push({
      rowNumber: item.rowNumber,
      techHub: techHub
    });
    
    if (i < batch.length - 1) {
      Utilities.sleep(CONFIG.API_DELAY);
    }
  }
  
  return results;
}

/**
 * Check if location has a predefined mapping
 */
function checkCityMappings(location) {
  const locationKey = location.toLowerCase().trim();
  return CITY_MAPPINGS[locationKey] || null;
}

/**
 * Check if location should default to Los Angeles
 */
function shouldDefaultToLA(location) {
  const locationLower = location.toLowerCase().trim();
  return REMOTE_KEYWORDS.some(keyword => locationLower.includes(keyword));
}

/**
 * Get tech hub location from Gemini API
 */
function getTechHubFromGemini(location, apiKey) {
  const prompt = `Answer in the format: "City, State" (e.g. San Jose, CA), no extra words. Find a popular city, town, or suburb within 30 miles of "${location}" Return only the location.`;
  
  const payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }],
    generationConfig: {
      temperature: 0.1,
      topK: 1,
      topP: 0.8,
      maxOutputTokens: 50
    }
  };
  
  const options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-goog-api-key': apiKey
    },
    payload: JSON.stringify(payload)
  };
  
  try {
    const response = UrlFetchApp.fetch(`${CONFIG.GEMINI_API_URL}?key=${apiKey}`, options);
    const responseData = JSON.parse(response.getContentText());
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`API Error: ${response.getResponseCode()}`);
    }
    
    if (!responseData.candidates || !responseData.candidates[0] || !responseData.candidates[0].content) {
      throw new Error('Invalid API response structure');
    }
    
    const generatedText = responseData.candidates[0].content.parts[0].text.trim();
    
    if (isValidCityStateFormat(generatedText)) {
      return generatedText;
    } else {
      Logger.log(`Invalid format from API for "${location}": "${generatedText}". Using default.`);
      return CONFIG.DEFAULT_LOCATION;
    }
    
  } catch (error) {
    Logger.log(`Gemini API error for "${location}": ${error.toString()}`);
    return CONFIG.DEFAULT_LOCATION;
  }
}

/**
 * Validate if the response is in "City, State" format
 */
function isValidCityStateFormat(text) {
  const trimmed = text.trim();
  if (trimmed.length > 50 || trimmed.length < 5) return false;
  if (!trimmed.includes(',')) return false;
  
  const parts = trimmed.split(',');
  if (parts.length !== 2) return false;
  
  const city = parts[0].trim();
  const state = parts[1].trim();
  
  if (city.length === 0 || state.length === 0) return false;
  if (state.length === 1 || state.length > 20) return false;
  
  return true;
}

/**
 * Update the sheet with results
 */
function updateSheet(sheet, results) {
  for (const result of results) {
    try {
      const cell = sheet.getRange(`${CONFIG.OUTPUT_COLUMN}${result.rowNumber}`);
      cell.setValue(result.techHub);
    } catch (error) {
      Logger.log(`Error updating row ${result.rowNumber}: ${error.toString()}`);
    }
  }
}