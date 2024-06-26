import ExcelJS from 'exceljs';
import readline from 'readline';

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

function question(query) {
  return new Promise((resolve) => rl.question(query, resolve));
}

async function reformatPhoneNumbers(inputFile, outputFile, phoneNumberColumns) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFile);

  const worksheet = workbook.getWorksheet(1);
  const newWorkbook = new ExcelJS.Workbook();
  const newWorksheet = newWorkbook.addWorksheet('Formatted Data');

  // Copy headers
  const headers = worksheet.getRow(1).values;
  newWorksheet.getRow(1).values = headers;

  // Convert column letters to column numbers
  const phoneColumnNumbers = phoneNumberColumns.map(col => worksheet.getColumn(col).number);

  // Process data rows
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // Skip header row
      const newRow = newWorksheet.getRow(rowNumber);
      
      row.eachCell((cell, colNumber) => {
        if (phoneColumnNumbers.includes(cell.col)) {
          // Get the original value or formula
          let phoneNumber = cell.value;
          
          // If it's a formula (for cases like 123-456-7890 becoming -456)
          if (cell.formula) {
            phoneNumber = cell.formula;
          }
          
          // Convert to string and clean up
          phoneNumber = phoneNumber ? phoneNumber.toString() : '';
          phoneNumber = phoneNumber.replace(/[^\d-]/g, '');
          
          // Handle cases where it might be a negative number (like -456)
          if (phoneNumber.startsWith('-')) {
            // Attempt to reconstruct the original number
            const parts = phoneNumber.split('-');
            if (parts.length === 2) {
              phoneNumber = parts[1] + parts[0].slice(1);
            }
          }
          
          // Remove remaining hyphens and format
          phoneNumber = phoneNumber.replace(/-/g, '');
          
          if (phoneNumber.length === 10) {
            phoneNumber = '1' + phoneNumber;
          }
          
          // Store as text
          newRow.getCell(colNumber).value = phoneNumber;
          newRow.getCell(colNumber).numFmt = '@';
        } else {
          // Copy other cell values as-is
          newRow.getCell(colNumber).value = cell.value;
        }
      });
    }
  });

  // Save the new workbook
  await newWorkbook.xlsx.writeFile(outputFile);
  console.log(`Formatted data saved to ${outputFile}`);
  rl.close();
}

async function main() {
  const inputFile = await question("Enter the input file name (e.g., input.xlsx): ");
  const outputFile = await question("Enter the new file name for output (e.g., output.xlsx): ");
  const phoneColumnsInput = await question("Enter the column letters for phone numbers, separated by commas (e.g., B,D,F): ");
  const phoneNumberColumns = phoneColumnsInput.split(',').map(col => col.trim().toUpperCase());

  await reformatPhoneNumbers(inputFile, outputFile, phoneNumberColumns);
}

main();