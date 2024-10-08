import ExcelJS from 'exceljs';
import readline from 'readline';

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

function question(query) {
  return new Promise((resolve) => rl.question(query, resolve));
}

async function generateEmails(inputFile, outputFile, emailColumn, phoneColumn) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFile);

  const worksheet = workbook.getWorksheet(1);
  const emailColumnNumber = worksheet.getColumn(emailColumn).number;
  const phoneColumnNumber = worksheet.getColumn(phoneColumn).number;

  let generatedCount = 0;

  // Process data rows
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // Skip header row
      const emailCell = row.getCell(emailColumnNumber);
      const phoneCell = row.getCell(phoneColumnNumber);

      let email = emailCell.value ? emailCell.value.toString().trim() : '';
      let phoneNumber = phoneCell.value ? phoneCell.value.toString().replace(/\D/g, '') : '';

      if (!email && phoneNumber) {
        // Generate email if the email is blank and there's a phone number
        email = `unknown+${phoneNumber}@prostructengineering.com`;
        emailCell.value = email;
        generatedCount++;
      }
    }
  });

  // Save the modified workbook
  await workbook.xlsx.writeFile(outputFile);
  console.log(`Modified data saved to ${outputFile}`);
  console.log(`Total emails generated: ${generatedCount}`);
}

async function main() {
  const inputFile = await question("Enter the input file name (e.g., input.xlsx): ");
  const outputFile = await question("Enter the output file name (e.g., output_with_emails.xlsx): ");
  const emailColumn = await question("Enter the column letter for email addresses (e.g., 'D'): ");
  const phoneColumn = await question("Enter the column letter for phone numbers (e.g., 'C'): ");

  await generateEmails(inputFile, outputFile, emailColumn, phoneColumn);
  rl.close();
}

main();