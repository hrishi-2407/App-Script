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
  const batchSize = 40;

  const baseResumeId = '1-HZj9OROYQ_ZiZFFmoPbqQCNk6GExToS6mjWj-aSOx8';
  const shareEmails = ['hrishibari2002@gmail.com', 'suhas112001@gmail.com'];

  const dataRange = sheet.getRange(startRow, 1, lastRow - startRow + 1, 11);
  const data = dataRange.getValues();

  let processedCount = 0;

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
      sheet.getRange(currentSheetRow, 9).setValue('❌ No location provided');
      continue;
    }

    try {
      const newResume = DriveApp.getFileById(baseResumeId).makeCopy(`${companyName}_Yeswanth Koti Resume`);
      const newResumeId = newResume.getId();
      const doc = DocumentApp.openById(newResumeId);
      const body = doc.getBody();
      body.replaceText('{{LOCATION}}', changedLocation);
      doc.saveAndClose();

      shareEmails.forEach(email => {
        DriveApp.getFileById(newResumeId).addEditor(email);
      });

      const newResumeUrl = `https://docs.google.com/document/d/${newResumeId}/edit`;
      sheet.getRange(currentSheetRow, 9).setValue(newResumeUrl);

      processedCount++;

      if (processedCount % 10 === 0) {
        Utilities.sleep(5000);
      }

      if (processedCount >= batchSize) {
        SpreadsheetApp.getUi().alert(`Processed ${batchSize} resumes. Run again for remaining.`);
        return;
      }

    } catch (err) {
      sheet.getRange(currentSheetRow, 9).setValue(`❌ Error: ${err.message}`);
    }
  }

  SpreadsheetApp.getUi().alert(`Resume generation completed. Total processed: ${processedCount}`);
}
