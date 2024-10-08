import ExcelJS from 'exceljs';
import readline from 'readline';

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

function question(query) {
  return new Promise((resolve) => rl.question(query, resolve));
}

async function compareExcelFiles(file1, file2, outputFile, idColumn) {
  const workbook1 = new ExcelJS.Workbook();
  const workbook2 = new ExcelJS.Workbook();
  await workbook1.xlsx.readFile(file1);
  await workbook2.xlsx.readFile(file2);

  const worksheet1 = workbook1.getWorksheet(1);
  const worksheet2 = workbook2.getWorksheet(1);

  const outputWorkbook = new ExcelJS.Workbook();
  const outputWorksheet = outputWorkbook.addWorksheet('Common Rows');

  // Copy headers from the first file
  const headers = worksheet1.getRow(1).values;
  outputWorksheet.getRow(1).values = headers;

  // Create a map of IDs from the second file for quick lookup
  const idMap = new Map();
  worksheet2.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // Skip header row
      const id = row.getCell(idColumn).value;
      idMap.set(id, row);
    }
  });

  let outputRowIndex = 2; // Start from row 2 (after headers)

  // Compare rows and write common ones to the output file
  worksheet1.eachRow((row1, rowNumber) => {
    if (rowNumber > 1) { // Skip header row
      const id = row1.getCell(idColumn).value;
      if (idMap.has(id)) {
        // ID exists in both files, copy the row to the output
        copyRow(row1, outputWorksheet.getRow(outputRowIndex));
        outputRowIndex++;
      }
    }
  });

  // Save the output workbook
  await outputWorkbook.xlsx.writeFile(outputFile);
  console.log(`Common rows saved to ${outputFile}`);
  rl.close();
}

function copyRow(sourceRow, targetRow) {
  sourceRow.eachCell((cell, colNumber) => {
    targetRow.getCell(colNumber).value = cell.value;
  });
}

async function main() {
  const file1 = await question("Enter the name of the first Excel file: ");
  const file2 = await question("Enter the name of the second Excel file: ");
  const outputFile = await question("Enter the name for the output file: ");
  const idColumn = await question("Enter the column letter for the ID field: ");
  
  await compareExcelFiles(file1, file2, outputFile, idColumn);
}

main();