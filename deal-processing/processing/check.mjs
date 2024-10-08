import ExcelJS from 'exceljs';
import readline from 'readline';

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

function question(query) {
  return new Promise((resolve) => rl.question(query, resolve));
}

async function countUniqueRows(inputFile, idColumn) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFile);

  const worksheet = workbook.getWorksheet(1);
  const idColumnNumber = worksheet.getColumn(idColumn).number;

  const uniqueIds = new Set();

  // Process data rows
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // Skip header row
      const idCell = row.getCell(idColumnNumber);
      const id = idCell.value;
      if (id !== null && id !== undefined) {
        uniqueIds.add(id.toString());
      }
    }
  });

  return uniqueIds.size;
}

async function main() {
  const inputFile = await question("Enter the input file name (e.g., input.xlsx): ");
  const idColumn = await question("Enter the column letter for the ID (e.g., 'A'): ");

  try {
    const uniqueCount = await countUniqueRows(inputFile, idColumn);
    console.log(`Number of unique rows based on ID column: ${uniqueCount}`);
  } catch (error) {
    console.error("An error occurred:", error.message);
  }

  rl.close();
}

main();