function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Job Tools')
    .addItem('Resume', 'generateAndShareResumes')
    .addSeparator()
    .addItem('Location', 'enhanceJobLocations')
    .addToUi();
}

function generateAndShareResumes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Applications');
  const lastRow = sheet.getLastRow();
  const startRow = 4;
  const batchSize = 35;

  const baseResumeId = '1Kh8VauHOGSy3PDverq5ZLlEqGo8xOlP-bRuHowQol6g';
  const shareEmails = ['hrishibari2002@gmail.com', 'suhas112001@gmail.com'];

  const dataRange = sheet.getRange(startRow, 1, lastRow - startRow + 1, 11);
  const data = dataRange.getValues();

  let itemsProcessed = 0;
  const sheetUpdates = []; // Array to collect all sheet updates for batch operation
  const updateRows = []; // Array to track which rows need updates

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const companyName = row[5];  
    const changedLocation = row[10];
    const existingResumeLink = row[8];  // Column I
    const currentSheetRow = startRow + i;

    if (!companyName) {
      continue;
    }

    if (existingResumeLink) {
      // Already processed — skip
      continue;
    }

    if (!changedLocation) {
      sheetUpdates.push(['❌ No location provided']);
      updateRows.push(currentSheetRow);
      itemsProcessed++;
      continue;
    }

    try {
      // Create copy of the resume
      const newResume = DriveApp.getFileById(baseResumeId).makeCopy(`${companyName}_Yeswanth Koti Resume`);
      const newResumeId = newResume.getId();
      
      // Open document, replace text, and save
      const doc = DocumentApp.openById(newResumeId);
      const body = doc.getBody();
      body.replaceText('{{LOCATION}}', changedLocation);
      doc.saveAndClose();

      // Optimize sharing - get file object once and reuse
      const newResumeFile = DriveApp.getFileById(newResumeId);
      shareEmails.forEach(email => {
        newResumeFile.addEditor(email);
      });

      // Prepare sheet update data
      const newResumeUrl = `https://docs.google.com/document/d/${newResumeId}/edit`;
      sheetUpdates.push([newResumeUrl]);
      updateRows.push(currentSheetRow);

      itemsProcessed++;

      // Removed unnecessary sleep - Google Apps Script handles rate limiting automatically

      if (itemsProcessed >= batchSize) {
        // Perform batch update before breaking
        performBatchSheetUpdate(sheet, sheetUpdates, updateRows);
        return;
      }

    } catch (err) {
      // Collect error for batch update
      sheetUpdates.push([`❌ Error: ${err.message}`]);
      updateRows.push(currentSheetRow);
      itemsProcessed++;
    }
  }

  // Perform final batch update for all remaining items
  performBatchSheetUpdate(sheet, sheetUpdates, updateRows);
}

/**
 * Helper function to perform batch sheet updates
 * @param {Sheet} sheet - The sheet object
 * @param {Array} updates - Array of values to update
 * @param {Array} rows - Array of row numbers corresponding to updates
 */
function performBatchSheetUpdate(sheet, updates, rows) {
  if (updates.length === 0) return;
  
  // Group consecutive rows for more efficient batch updates
  const ranges = [];
  let currentRange = { startRow: rows[0], values: [updates[0]] };
  
  for (let i = 1; i < rows.length; i++) {
    if (rows[i] === rows[i-1] + 1) {
      // Consecutive row - add to current range
      currentRange.values.push(updates[i]);
    } else {
      // Non-consecutive row - finalize current range and start new one
      ranges.push(currentRange);
      currentRange = { startRow: rows[i], values: [updates[i]] };
    }
  }
  // Don't forget the last range
  ranges.push(currentRange);
  
  // Apply batch updates for each range
  ranges.forEach(range => {
    if (range.values.length === 1) {
      // Single cell update
      sheet.getRange(range.startRow, 9).setValue(range.values[0][0]);
    } else {
      // Multi-cell batch update
      sheet.getRange(range.startRow, 9, range.values.length, 1).setValues(range.values);
    }
  });
}