import ExcelJS from 'exceljs';
import readline from 'readline';

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

function question(query) {
  return new Promise((resolve) => rl.question(query, resolve));
}

async function removeLeadingOne(inputFile, outputFile, phoneColumn) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFile);

  const worksheet = workbook.getWorksheet(1);
  const phoneColumnNumber = worksheet.getColumn(phoneColumn).number;

  let modifiedCount = 0;

  // Process data rows
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // Skip header row
      const phoneCell = row.getCell(phoneColumnNumber);
      let phoneNumber = phoneCell.value ? phoneCell.value.toString() : '';

      // Remove non-digit characters
      phoneNumber = phoneNumber.replace(/\D/g, '');

      if (phoneNumber.startsWith('11')) {
        // Remove the leading '1'
        phoneNumber = phoneNumber.substring(1);
        phoneCell.value = phoneNumber;
        modifiedCount++;
      }
    }
  });

  // Save the modified workbook
  await workbook.xlsx.writeFile(outputFile);
  console.log(`Modified data saved to ${outputFile}`);
  console.log(`Total numbers modified: ${modifiedCount}`);
}

async function main() {
  const inputFile = await question("Enter the input file name (e.g., starts_with_11.xlsx): ");
  const outputFile = await question("Enter the output file name (e.g., modified_numbers.xlsx): ");
  const phoneColumn = await question("Enter the column letter for phone numbers (e.g., 'C'): ");

  await removeLeadingOne(inputFile, outputFile, phoneColumn);
  rl.close();
}

main();