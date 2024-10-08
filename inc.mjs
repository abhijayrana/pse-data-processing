import ExcelJS from 'exceljs';
import readline from 'readline';

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

function question(query) {
  return new Promise((resolve) => rl.question(query, resolve));
}

async function mergeCompanyNames(inputFile, outputFile, companyColumn) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFile);

  const worksheet = workbook.getWorksheet(1);
  const companyColumnNumber = worksheet.getColumn(companyColumn).number;

  let rowsToDelete = [];
  let mergeCount = 0;

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // Skip header row
      const currentCompanyCell = row.getCell(companyColumnNumber);
      const nextRow = worksheet.getRow(rowNumber + 1);

      const currentCompanyName = currentCompanyCell.value ? currentCompanyCell.value.toString().trim() : '';
      const nextCompanyName = nextRow.getCell(companyColumnNumber).value ? nextRow.getCell(companyColumnNumber).value.toString().trim() : '';

      // Check if the current row is not "Inc." and the next row is "Inc.", "inc.", or "inc"
      if (currentCompanyName && nextCompanyName.match(/^inc\.?$/i)) {
        // Merge the current company name with the exact next line text, adding a comma
        currentCompanyCell.value = `${currentCompanyName}, ${nextCompanyName}`;

        // Mark the next row for deletion
        rowsToDelete.push(rowNumber + 1);
        mergeCount++;
      }
    }
  });

  // Delete marked rows
  rowsToDelete.reverse().forEach((rowNumber) => {
    worksheet.spliceRows(rowNumber, 1);
  });

  // Save the modified workbook
  await workbook.xlsx.writeFile(outputFile);
  console.log(`Modified data saved to ${outputFile}`);
  console.log(`Merged ${mergeCount} company names`);
  console.log(`Deleted ${rowsToDelete.length} rows`);
}

async function main() {
  try {
    const inputFile = await question("Enter the input file name (e.g., input.xlsx): ");
    const outputFile = await question("Enter the output file name (e.g., output_merged.xlsx): ");
    const companyColumn = await question("Enter the column letter for company names (e.g., 'B'): ");

    await mergeCompanyNames(inputFile, outputFile, companyColumn);
  } catch (error) {
    console.error("An error occurred:", error.message);
  } finally {
    rl.close();
  }
}

main();