import ExcelJS from 'exceljs';
import readline from 'readline';

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

function question(query) {
  return new Promise((resolve) => rl.question(query, resolve));
}

async function addHubSpotIds(dictionaryFile, mainFile, outputFile, dictionaryNameColumn, dictionaryIdColumn, mainNameColumn) {
  // Read the dictionary file
  const dictionaryWorkbook = new ExcelJS.Workbook();
  await dictionaryWorkbook.xlsx.readFile(dictionaryFile);
  const dictionaryWorksheet = dictionaryWorkbook.getWorksheet(1);

  // Create a dictionary of deal names to HubSpot IDs
  const dealDictionary = {};
  dictionaryWorksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // Skip header row
      const dealName = row.getCell(dictionaryNameColumn).value;
      const hubSpotId = row.getCell(dictionaryIdColumn).value;
      dealDictionary[dealName] = hubSpotId;
    }
  });

  // Read the main file
  const mainWorkbook = new ExcelJS.Workbook();
  await mainWorkbook.xlsx.readFile(mainFile);
  const mainWorksheet = mainWorkbook.getWorksheet(1);

  // Insert a new column A for HubSpot IDs
  mainWorksheet.spliceColumns(1, 0, ['HubSpot ID']);

  // Process each row in the main file
  mainWorksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) {
      // Set header for the new column
      row.getCell(1).value = 'HubSpot ID';
    } else {
      const dealName = row.getCell(mainNameColumn + 1).value; // +1 because we inserted a new column
      const hubSpotId = dealDictionary[dealName] || '';
      row.getCell(1).value = hubSpotId;
    }
  });

  // Save the modified workbook
  await mainWorkbook.xlsx.writeFile(outputFile);
  console.log(`Updated file saved to ${outputFile}`);
  rl.close();
}

async function main() {
  const dictionaryFile = await question("Enter the dictionary file name (with deal names and HubSpot IDs): ");
  const mainFile = await question("Enter the main file name (with deal entries): ");
  const outputFile = await question("Enter the output file name: ");
  const dictionaryNameColumn = await question("Enter the column letter for deal names in the dictionary file: ");
  const dictionaryIdColumn = await question("Enter the column letter for HubSpot IDs in the dictionary file: ");
  const mainNameColumn = await question("Enter the column letter for deal names in the main file: ");

  await addHubSpotIds(
    dictionaryFile, 
    mainFile, 
    outputFile, 
    dictionaryNameColumn.toUpperCase().charCodeAt(0) - 64, // Convert letter to column number
    dictionaryIdColumn.toUpperCase().charCodeAt(0) - 64,
    mainNameColumn.toUpperCase().charCodeAt(0) - 64
  );
}

main();