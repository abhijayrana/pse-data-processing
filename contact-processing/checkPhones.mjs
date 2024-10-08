import ExcelJS from 'exceljs';
import readline from 'readline';

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

function question(query) {
  return new Promise((resolve) => rl.question(query, resolve));
}

async function validatePhoneNumbers(inputFile, outputFile, phoneColumn) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFile);

  const worksheet = workbook.getWorksheet(1);
  const newWorkbook = new ExcelJS.Workbook();
  const newWorksheet = newWorkbook.addWorksheet('Invalid Phone Numbers');

  // Copy headers
  const headers = worksheet.getRow(1).values;
  newWorksheet.getRow(1).values = headers;

  let newRowIndex = 2; // Start from row 2 (after headers)

  // Process data rows
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // Skip header row
      const phoneCell = row.getCell(phoneColumn);
      const phoneNumber = phoneCell.value ? phoneCell.value.toString().replace(/\D/g, '') : '';

      if (phoneNumber.length > 0 && !phoneNumber.startsWith('91') && phoneNumber.length !== 11) {
        // Copy the entire row to the new worksheet
        const newRow = newWorksheet.getRow(newRowIndex);
        row.eachCell((cell, colNumber) => {
          newRow.getCell(colNumber).value = cell.value;
        });
        newRowIndex++;
      }
    }
  });

  // Save the new workbook
  await newWorkbook.xlsx.writeFile(outputFile);
  console.log(`Rows with invalid phone numbers saved to ${outputFile}`);
  console.log(`Total invalid entries: ${newRowIndex - 2}`);
}

async function main() {
  const inputFile = await question("Enter the input file name (e.g., input.xlsx): ");
  const outputFile = await question("Enter the output file name for invalid entries (e.g., invalid_phones.xlsx): ");
  const phoneColumn = await question("Enter the column letter for phone numbers (e.g., 'C'): ");

  await validatePhoneNumbers(inputFile, outputFile, phoneColumn);
  rl.close();
}

main();