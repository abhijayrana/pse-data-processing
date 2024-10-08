import ExcelJS from 'exceljs';
import readline from 'readline';

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

function question(query) {
  return new Promise((resolve) => rl.question(query, resolve));
}

async function extractRowsWithCommas(inputFile, outputFile, relatedContactsColumn) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFile);

  const inputWorksheet = workbook.getWorksheet(1);
  const outputWorkbook = new ExcelJS.Workbook();
  const outputWorksheet = outputWorkbook.addWorksheet('Rows with Multiple Contacts');

  // Copy headers
  const headers = inputWorksheet.getRow(1).values;
  outputWorksheet.getRow(1).values = headers;

  let extractedRowCount = 0;
  let outputRowIndex = 2; // Start from row 2 (after headers)

  // Process data rows
  inputWorksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // Skip header row
      const relatedContactsCell = row.getCell(relatedContactsColumn);
      const relatedContacts = relatedContactsCell.value ? relatedContactsCell.value.toString() : '';

      if (relatedContacts.includes(',')) {
        // Copy the entire row to the output worksheet
        const newRow = outputWorksheet.getRow(outputRowIndex);
        row.eachCell((cell, colNumber) => {
          newRow.getCell(colNumber).value = cell.value;
        });
        outputRowIndex++;
        extractedRowCount++;
      }
    }
  });

  // Save the new workbook
  await outputWorkbook.xlsx.writeFile(outputFile);
  console.log(`Extracted rows saved to ${outputFile}`);
  console.log(`Total rows extracted: ${extractedRowCount}`);
}

async function main() {
  try {
    const inputFile = await question("Enter the input file name (e.g., input.xlsx): ");
    const outputFile = await question("Enter the output file name (e.g., multiple_contacts.xlsx): ");
    const relatedContactsColumn = await question("Enter the column letter for Related Contacts (e.g., 'C'): ");

    await extractRowsWithCommas(inputFile, outputFile, relatedContactsColumn);
  } catch (error) {
    console.error("An error occurred:", error.message);
  } finally {
    rl.close();
  }
}

main();