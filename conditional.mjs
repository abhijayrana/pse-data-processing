import ExcelJS from 'exceljs';
import readline from 'readline';

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

function question(query) {
  return new Promise((resolve) => rl.question(query, resolve));
}

async function splitCompanies(inputFile, outputFile, idColumn, companiesColumn, exceptionColumn) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFile);

  const inputWorksheet = workbook.getWorksheet(1);
  const outputWorkbook = new ExcelJS.Workbook();
  const outputWorksheet = outputWorkbook.addWorksheet('Processed Data');

  // Copy headers
  const headers = inputWorksheet.getRow(1).values;
  outputWorksheet.getRow(1).values = headers;

  let outputRowIndex = 2; // Start from row 2 (after headers)
  let splitCount = 0;
  let unchangedCount = 0;

  // Process data rows
  inputWorksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // Skip header row
      const id = row.getCell(idColumn).value;
      const companies = row.getCell(companiesColumn).value ? row.getCell(companiesColumn).value.toString() : '';
      const exceptionStatus = row.getCell(exceptionColumn).value;

      if (!exceptionStatus) {
        // Split companies if exception_status is blank
        const companyList = companies.split(',').map(company => company.trim()).filter(company => company);
        
        companyList.forEach(company => {
          const newRow = outputWorksheet.addRow([id, company, '']);
          outputRowIndex++;
        });
        
        splitCount += companyList.length > 0 ? 1 : 0;
      } else {
        // Copy the row as-is if there's an exception status
        outputWorksheet.addRow([id, companies, exceptionStatus]);
        outputRowIndex++;
        unchangedCount++;
      }
    }
  });

  // Save the new workbook
  await outputWorkbook.xlsx.writeFile(outputFile);
  console.log(`Processed data saved to ${outputFile}`);
  console.log(`Rows split: ${splitCount}`);
  console.log(`Rows unchanged: ${unchangedCount}`);
}

async function main() {
  try {
    const inputFile = await question("Enter the input file name (e.g., input.xlsx): ");
    const outputFile = await question("Enter the output file name (e.g., output_split.xlsx): ");
    const idColumn = await question("Enter the column letter for ID (e.g., 'A'): ");
    const companiesColumn = await question("Enter the column letter for Related Companies (e.g., 'B'): ");
    const exceptionColumn = await question("Enter the column letter for Exception Status (e.g., 'C'): ");

    await splitCompanies(inputFile, outputFile, idColumn, companiesColumn, exceptionColumn);
  } catch (error) {
    console.error("An error occurred:", error.message);
  } finally {
    rl.close();
  }
}

main();