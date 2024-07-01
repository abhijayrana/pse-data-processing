import ExcelJS from 'exceljs';
import readline from 'readline';

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

function question(query) {
  return new Promise((resolve) => rl.question(query, resolve));
}

async function reformatOwnerColumns(inputFile, outputFile, ownerColumns) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFile);

  const worksheet = workbook.getWorksheet(1);
  const newWorkbook = new ExcelJS.Workbook();
  const newWorksheet = newWorkbook.addWorksheet('Formatted Data');

  // Copy headers
  const headers = worksheet.getRow(1).values;
  newWorksheet.getRow(1).values = headers;

  // Convert column letters to column numbers
  const ownerColumnNumbers = ownerColumns.map(col => worksheet.getColumn(col).number);

  // Process data rows
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // Skip header row
      const newRow = newWorksheet.getRow(rowNumber);
      
      row.eachCell((cell, colNumber) => {
        if (ownerColumnNumbers.includes(cell.col)) {
          let ownerValue = cell.value ? cell.value.toString() : '';
          // Extract email from between parentheses
          let email = ownerValue.match(/\((.*?)\)/);
          email = email ? email[1] : ''; // If match found, take the first captured group
          newRow.getCell(colNumber).value = email;
        } else {
          newRow.getCell(colNumber).value = cell.value;
        }
      });
    }
  });

  // Save the new workbook
  await newWorkbook.xlsx.writeFile(outputFile);
  console.log(`Owner columns formatted and saved to ${outputFile}`);
}

async function main() {
  const inputFile = await question("Enter the input file name (e.g., input.xlsx): ");
  const outputFile = await question("Enter the new file name for output (e.g., output.xlsx): ");
  const ownerColumnsInput = await question("Enter the column letters for Owner columns, separated by commas (e.g., C,E,G): ");
  const ownerColumns = ownerColumnsInput.split(',').map(col => col.trim().toUpperCase());

  await reformatOwnerColumns(inputFile, outputFile, ownerColumns);
  rl.close();
}

main();