import ExcelJS from 'exceljs';
import readline from 'readline';

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

function question(query) {
  return new Promise((resolve) => rl.question(query, resolve));
}

async function keepFirstEmail(inputFile, outputFile, emailColumn) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFile);

  const worksheet = workbook.getWorksheet(1);

  // Process data rows
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // Skip header row
      const emailCell = row.getCell(emailColumn);
      const emails = emailCell.value ? emailCell.value.toString().split(',') : [];

      if (emails.length > 1) {
        // Keep only the first email
        emailCell.value = emails[0].trim();
      }
    }
  });

  // Save the modified workbook
  await workbook.xlsx.writeFile(outputFile);
  console.log(`Modified data saved to ${outputFile}`);
  rl.close();
}

async function main() {
  const inputFile = await question("Enter the input file name (e.g., input.xlsx): ");
  const outputFile = await question("Enter the new file name for output (e.g., output_modified.xlsx): ");
  const emailColumn = await question("Enter the column letter for Email (e.g., 'E'): ");
  
  await keepFirstEmail(inputFile, outputFile, emailColumn);
}

main();