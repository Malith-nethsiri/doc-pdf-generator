// ⚙️ Replace these with your actual IDs
const TEMPLATE_ID = '1XPqMwcLANgqtdPXOc6Wh-HJEADGma8vHa9FbaCar61g';
const OUTPUT_FOLDER_ID = '1e7jAsyaKgvGJ3-Eqi4vGCXrkr8uqYwKs';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📄 Document Generator')
    .addItem('🔁 Generate Docs & PDFs', 'generateDocsFromSheet')
    .addToUi();
}

function generateDocsFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ClientData");
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ Sheet 'ClientData' not found. Please check the tab name.");
    return;
  }

  const data = sheet.getDataRange().getValues();
  const templateFile = DriveApp.getFileById(TEMPLATE_ID);
  const outputFolder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);

  for (let i = 1; i < data.length; i++) {
    const [name, amount, company, status] = data[i];

    // ⛔ Skip if already processed
    if (status === '✅ Done') continue;

    // 📝 Create a new doc from the template
    const newDocFile = templateFile.makeCopy(`Invoice - ${name}`, outputFolder);
    const newDoc = DocumentApp.openById(newDocFile.getId());
    const body = newDoc.getBody();

    // 🔁 Replace placeholders
    body.replaceText('{{name}}', name);
    body.replaceText('{{amount}}', amount);
    body.replaceText('{{company}}', company);
    newDoc.saveAndClose();

    // 📄 Convert to PDF
    const pdfBlob = newDocFile.getAs(MimeType.PDF);
    const pdfFile = outputFolder.createFile(pdfBlob).setName(`Invoice - ${name}.pdf`);

    // 🖊️ Mark as complete in the sheet
    sheet.getRange(i + 1, 4).setValue('✅ Done');
  }

  SpreadsheetApp.getUi().alert("✅ All documents and PDFs generated successfully!");
}
