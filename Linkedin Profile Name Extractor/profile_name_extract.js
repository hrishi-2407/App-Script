/**
 * LinkedIn Name Extractor
 * 
 * Extracts candidate names from LinkedIn profile URLs in a Google Sheet
 * - Processes URLs in Column C and outputs names to Column D
 * - Handles both hyphenated URLs and single-word usernames
 * - Maintains proper error handling and rate limiting
 */

function extractLinkedInNames() {
  try {
    // Get active sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Find the last row with data in column C
    const lastRow = getLastRowWithData(sheet, 3); // Column C is index 3
    
    if (lastRow === 0) {
      SpreadsheetApp.getUi().alert("No LinkedIn URLs found in Column C!");
      return;
    }
    
    // Get all URLs and existing names
    const urlRange = sheet.getRange(1, 3, lastRow, 1);  // Column C
    const nameRange = sheet.getRange(1, 4, lastRow, 1); // Column D
    
    const urls = urlRange.getValues();
    const existingNames = nameRange.getValues();
    
    // Track results for final summary
    let processed = 0;
    let namesByHyphen = 0;
    let namesByFetch = 0;
    let namesNotFound = 0;
    let invalidUrls = 0;
    let skipped = 0;
    
    // Process each URL
    for (let i = 0; i < urls.length; i++) {
      const url = urls[i][0].toString().trim();
      const existingName = existingNames[i][0].toString().trim();
      
      // Skip empty cells, non-URLs or cells that already have a name
      if (!url || !isLinkedInUrl(url)) {
        if (!url) {
          // Skip silently for empty cells
          skipped++;
        } else {
          // Mark as invalid
          sheet.getRange(i + 1, 4).setValue("Invalid URL");
          invalidUrls++;
        }
        continue;
      }
      
      // Skip if name already exists
      if (existingName) {
        skipped++;
        continue;
      }
      
      // Extract LinkedIn username/path part
      const pathMatch = url.match(/linkedin\.com\/in\/([^\/\?]+)/i);
      if (!pathMatch || !pathMatch[1]) {
        sheet.getRange(i + 1, 4).setValue("Invalid URL");
        invalidUrls++;
        continue;
      }
      
      const profilePath = pathMatch[1];
      let name = "";
      
      // Handle hyphenated names
      if (profilePath.includes("-")) {
        name = extractNameFromHyphenatedPath(profilePath);
        sheet.getRange(i + 1, 4).setValue(name);
        namesByHyphen++;
      } else {
        // Single-word username - need to fetch the page
        try {
          name = fetchNameFromLinkedIn(url);
          if (name) {
            sheet.getRange(i + 1, 4).setValue(name);
            namesByFetch++;
          } else {
            sheet.getRange(i + 1, 4).setValue("Name Not Found");
            namesNotFound++;
          }
          // Sleep to avoid rate limiting, but only for actual HTTP requests
          Utilities.sleep(1000);
        } catch (error) {
          sheet.getRange(i + 1, 4).setValue("Name Not Found");
          namesNotFound++;
          // Still sleep on error to maintain rate limiting
          Utilities.sleep(1000);
        }
      }
      
      processed++;
      
      // Update every 5 URLs to avoid timeout issues
      if (processed % 5 === 0) {
        SpreadsheetApp.flush();
      }
    }
    
    // Show completion message
    const message = `
LinkedIn Name Extraction Complete:
- Total URLs processed: ${processed}
- Names extracted from hyphenated URLs: ${namesByHyphen}
- Names extracted via page fetch: ${namesByFetch}
- Names not found: ${namesNotFound}
- Invalid URLs: ${invalidUrls}
- Skipped (empty or already processed): ${skipped}
    `;
    
    SpreadsheetApp.getUi().alert(message);
  } catch (error) {
    SpreadsheetApp.getUi().alert("An error occurred: " + error.toString());
  }
}

/**
 * Check if a string is a valid LinkedIn URL
 */
function isLinkedInUrl(url) {
  return url && typeof url === 'string' && 
         url.toLowerCase().includes('linkedin.com/in/');
}

/**
 * Extract and format a name from a hyphenated LinkedIn profile path
 */
function extractNameFromHyphenatedPath(profilePath) {
  // Remove any trailing numeric ID (e.g., "540a831a0")
  const pathWithoutId = profilePath.replace(/-[0-9a-f]+\/?$/, '');
  
  // Split by hyphens and capitalize each part
  const nameParts = pathWithoutId.split('-');
  const capitalizedParts = nameParts.map(part => {
    return part.charAt(0).toUpperCase() + part.slice(1).toLowerCase();
  });
  
  // Join with spaces
  return capitalizedParts.join(' ');
}

/**
 * Fetch the LinkedIn profile page and extract the name from the title
 */
function fetchNameFromLinkedIn(url) {
  try {
    // Use a desktop browser user-agent
    const options = {
      "method": "get",
      "headers": {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
      },
      "muteHttpExceptions": true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const content = response.getContentText();
    
    // Check for HTTP success
    if (response.getResponseCode() !== 200) {
      return null;
    }
    
    // Extract title
    const titleMatch = content.match(/<title>(.*?)\s*\|\s*LinkedIn<\/title>/i);
    
    if (titleMatch && titleMatch[1] && !titleMatch[1].includes("Page Not Found")) {
      // Clean up the name (trim and handle special characters)
      return titleMatch[1].trim();
    }
    
    return null;
  } catch (error) {
    Logger.log("Error fetching LinkedIn page: " + error.toString());
    return null;
  }
}

/**
 * Find the last row with data in the specified column
 */
function getLastRowWithData(sheet, columnIndex) {
  const values = sheet.getRange(1, columnIndex, sheet.getLastRow(), 1).getValues();
  
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "") {
      return i + 1;
    }
  }
  
  return 0;
}

/**
 * Creates a menu item to run the extraction
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('LinkedIn Tools')
    .addItem('Extract Names from LinkedIn URLs', 'extractLinkedInNames')
    .addToUi();
}