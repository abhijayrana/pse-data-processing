// processCompany.mjs

import ExcelJS from 'exceljs';
import readline from 'readline';

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

function question(query) {
  return new Promise((resolve) => rl.question(query, resolve));
}

async function processContacts(inputFile, validOutputFile, companyColumn) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFile);

  const worksheet = workbook.getWorksheet(1);
  const validWorkbook = new ExcelJS.Workbook();
  const validWorksheet = validWorkbook.addWorksheet('Valid Contacts');

  // Copy headers to the valid worksheet
  const headers = worksheet.getRow(1).values;
  validWorksheet.getRow(1).values = headers;

  let validRowIndex = 2; // Start from row 2 (after headers)

  // Process data rows
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // Skip header row
      const companyCell = row.getCell(companyColumn);
      const company = companyCell.value ? companyCell.value.toString().trim() : '';

      if (company) {
        // Valid company
        copyRow(row, validWorksheet.getRow(validRowIndex));
        validRowIndex++;
      }
    }
  });

  // Save the new workbook
  await validWorkbook.xlsx.writeFile(validOutputFile);
  console.log(`Contacts with valid companies saved to ${validOutputFile}`);
  rl.close();
}

function copyRow(sourceRow, targetRow) {
  sourceRow.eachCell((cell, colNumber) => {
    targetRow.getCell(colNumber).value = cell.value;
  });
}

async function main() {
  const inputFile = await question("Enter the input file name (e.g., contacts.xlsx): ");
  const validOutputFile = await question("Enter the file name for contacts with valid companies (e.g., valid_contacts.xlsx): ");
  const companyColumn = await question("Enter the column letter for Company Name (e.g., 'D'): ");
  
  await processContacts(inputFile, validOutputFile, companyColumn);
}

main();
