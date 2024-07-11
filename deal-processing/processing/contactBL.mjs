// processContacts.mjs

import ExcelJS from 'exceljs';
import readline from 'readline';

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

function question(query) {
  return new Promise((resolve) => rl.question(query, resolve));
}

async function processContacts(inputFile, validOutputFile, missingEmailOutputFile, emailColumn) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFile);

  const worksheet = workbook.getWorksheet(1);
  const validWorkbook = new ExcelJS.Workbook();
  const missingEmailWorkbook = new ExcelJS.Workbook();
  const validWorksheet = validWorkbook.addWorksheet('Valid Contacts');
  const missingEmailWorksheet = missingEmailWorkbook.addWorksheet('Missing Email Contacts');

  // Copy headers to both worksheets
  const headers = worksheet.getRow(1).values;
  validWorksheet.getRow(1).values = headers;
  missingEmailWorksheet.getRow(1).values = headers;

  let validRowIndex = 2; // Start from row 2 (after headers)
  let missingEmailRowIndex = 2;

  // Process data rows
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // Skip header row
      const emailCell = row.getCell(emailColumn);
      const email = emailCell.value ? emailCell.value.toString().trim() : '';

      if (email) {
        // Valid email
        copyRow(row, validWorksheet.getRow(validRowIndex));
        validRowIndex++;
      } else {
        // Missing email
        copyRow(row, missingEmailWorksheet.getRow(missingEmailRowIndex));
        missingEmailRowIndex++;
      }
    }
  });

  // Save the new workbooks
  await validWorkbook.xlsx.writeFile(validOutputFile);
  await missingEmailWorkbook.xlsx.writeFile(missingEmailOutputFile);
  console.log(`Contacts with valid emails saved to ${validOutputFile}`);
  console.log(`Contacts with missing emails saved to ${missingEmailOutputFile}`);
  rl.close();
}

function copyRow(sourceRow, targetRow) {
  sourceRow.eachCell((cell, colNumber) => {
    targetRow.getCell(colNumber).value = cell.value;
  });
}

async function main() {
  const inputFile = await question("Enter the input file name (e.g., contacts.xlsx): ");
  const validOutputFile = await question("Enter the file name for contacts with valid emails (e.g., valid_contacts.xlsx): ");
  const missingEmailOutputFile = await question("Enter the file name for contacts with missing emails (e.g., missing_email_contacts.xlsx): ");
  const emailColumn = await question("Enter the column letter for Contact Email (e.g., 'C'): ");
  
  await processContacts(inputFile, validOutputFile, missingEmailOutputFile, emailColumn);
}

main();