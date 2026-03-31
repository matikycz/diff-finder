const ExcelJS = require('exceljs');
const path = require('path');

async function processExcelFile(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  workbook.eachSheet((worksheet, sheetId) => {
    console.log(`\nSheet ${sheetId}: ${worksheet.name}`);
    console.log(`Dimensions: ${worksheet.rowCount} rows x ${worksheet.columnCount} columns`);

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const values = row.values.slice(1); // slice off the leading undefined at index 0
      console.log(`  Row ${rowNumber}:`, values);
    });
  });
}

const filePath = path.join(__dirname, 'testfile.xlsx');

processExcelFile(filePath)
  .then(() => console.log('\nDone.'))
  .catch((err) => {
    console.error('Error processing file:', err.message);
    process.exit(1);
  });
