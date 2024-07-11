import ExcelJS from 'exceljs';
import readline from 'readline';

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

function question(query) {
  return new Promise((resolve) => rl.question(query, resolve));
}

async function splitRelatedContacts(inputFile, outputFile, relatedContactsColumn) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFile);

  const worksheet = workbook.getWorksheet(1);
  const newWorkbook = new ExcelJS.Workbook();
  const newWorksheet = newWorkbook.addWorksheet('Split Contacts');

  // Copy headers
  const headers = worksheet.getRow(1).values;
  newWorksheet.getRow(1).values = headers;

  let newRowIndex = 2; // Start from row 2 (after headers)

  // Process data rows
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // Skip header row
      const relatedContactsCell = row.getCell(relatedContactsColumn);
      const relatedContacts = relatedContactsCell.value ? relatedContactsCell.value.toString().split(',') : [];

      if (relatedContacts.length === 0) {
        // If no related contacts, just copy the row as is
        copyRow(row, newWorksheet.getRow(newRowIndex));
        newRowIndex++;
      } else {
        // For each related contact, create a new row
        relatedContacts.forEach(contact => {
          const newRow = newWorksheet.getRow(newRowIndex);
          copyRow(row, newRow);
          newRow.getCell(relatedContactsColumn).value = contact.trim();
          newRowIndex++;
        });
      }
    }
  });

  // Save the new workbook
  await newWorkbook.xlsx.writeFile(outputFile);
  console.log(`Split contacts data saved to ${outputFile}`);
  rl.close();
}

function copyRow(sourceRow, targetRow) {
  sourceRow.eachCell((cell, colNumber) => {
    targetRow.getCell(colNumber).value = cell.value;
  });
}

async function main() {
  const inputFile = await question("Enter the input file name (e.g., input.xlsx): ");
  const outputFile = await question("Enter the new file name for output (e.g., output_split.xlsx): ");
  const relatedContactsColumn = await question("Enter the column letter for Related Contacts (e.g., 'E'): ");
  
  await splitRelatedContacts(inputFile, outputFile, relatedContactsColumn);
}

main();