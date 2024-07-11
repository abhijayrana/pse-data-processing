import ExcelJS from 'exceljs';
import readline from 'readline';

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

function question(query) {
  return new Promise((resolve) => rl.question(query, resolve));
}

async function processDuplicateDeals(inputFile, uniqueOutputFile, duplicatesOutputFile, dealNameColumn) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFile);

  const worksheet = workbook.getWorksheet(1);
  const uniqueWorkbook = new ExcelJS.Workbook();
  const duplicatesWorkbook = new ExcelJS.Workbook();
  const uniqueWorksheet = uniqueWorkbook.addWorksheet('Unique Deals');
  const duplicatesWorksheet = duplicatesWorkbook.addWorksheet('Duplicate Deals');

  // Copy headers to both worksheets
  const headers = worksheet.getRow(1).values;
  uniqueWorksheet.getRow(1).values = headers;
  duplicatesWorksheet.getRow(1).values = headers;

  let uniqueRowIndex = 2; // Start from row 2 (after headers)
  let duplicateRowIndex = 2;
  const seenDeals = new Set();

  // Process data rows
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // Skip header row
      const dealNameCell = row.getCell(dealNameColumn);
      const dealName = dealNameCell.value ? dealNameCell.value.toString() : '';

      if (!seenDeals.has(dealName)) {
        // First occurrence of this deal name
        seenDeals.add(dealName);
        copyRow(row, uniqueWorksheet.getRow(uniqueRowIndex));
        uniqueRowIndex++;
      } else {
        // Duplicate deal name
        copyRow(row, duplicatesWorksheet.getRow(duplicateRowIndex));
        duplicateRowIndex++;
      }
    }
  });

  // Save the new workbooks
  await uniqueWorkbook.xlsx.writeFile(uniqueOutputFile);
  await duplicatesWorkbook.xlsx.writeFile(duplicatesOutputFile);
  console.log(`Unique deals saved to ${uniqueOutputFile}`);
  console.log(`Duplicate deals saved to ${duplicatesOutputFile}`);
  rl.close();
}

function copyRow(sourceRow, targetRow) {
  sourceRow.eachCell((cell, colNumber) => {
    targetRow.getCell(colNumber).value = cell.value;
  });
}

async function main() {
  const inputFile = await question("Enter the input file name (e.g., input.xlsx): ");
  const uniqueOutputFile = await question("Enter the file name for unique deals (e.g., unique_deals.xlsx): ");
  const duplicatesOutputFile = await question("Enter the file name for duplicate deals (e.g., duplicate_deals.xlsx): ");
  const dealNameColumn = await question("Enter the column letter for Deal Name (e.g., 'B'): ");
  
  await processDuplicateDeals(inputFile, uniqueOutputFile, duplicatesOutputFile, dealNameColumn);
}

main();