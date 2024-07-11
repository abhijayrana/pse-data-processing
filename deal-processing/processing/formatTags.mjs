import ExcelJS from 'exceljs';
import readline from 'readline';

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

function question(query) {
  return new Promise((resolve) => rl.question(query, resolve));
}

function formatTags(tags) {
  const specialTag = "Hard Bounce, Spam, Invalid, or Bad Email";
  
  const hasSpecialTag = tags.includes(specialTag);
  
  if (hasSpecialTag) {
    tags = tags.replace(specialTag, "SPECIAL_TAG_PLACEHOLDER");
  }
  
  let formattedTags = tags.split(',')
                          .map(tag => tag.trim())
                          .join(';');
  
  formattedTags = ';' + formattedTags + ';';
  
  if (hasSpecialTag) {
    formattedTags = formattedTags.replace("SPECIAL_TAG_PLACEHOLDER", specialTag);
  }
  
  return formattedTags;
}

async function reformatTagsColumns(inputFile, outputFile, tagColumns) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFile);

  const worksheet = workbook.getWorksheet(1);
  const newWorkbook = new ExcelJS.Workbook();
  const newWorksheet = newWorkbook.addWorksheet('Formatted Data');

  // Copy headers
  const headers = worksheet.getRow(1).values;
  newWorksheet.getRow(1).values = headers;

  // Process data rows
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // Skip header row
      const newRow = newWorksheet.getRow(rowNumber);
      
      row.eachCell((cell, colNumber) => {
        if (tagColumns.includes(worksheet.getColumn(colNumber).letter)) {
          // Get the original value
          let tagsValue = cell.value ? cell.value.toString() : '';
          
          // Format the tags
          let formattedTags = formatTags(tagsValue);
          
          // Store the formatted tags
          newRow.getCell(colNumber).value = formattedTags;
        } else {
          // Copy other cell values as-is
          newRow.getCell(colNumber).value = cell.value;
        }
      });
    }
  });

  // Save the new workbook
  await newWorkbook.xlsx.writeFile(outputFile);
  console.log(`Formatted data saved to ${outputFile}`);
  rl.close();
}

async function main() {
  const inputFile = await question("Enter the input file name (e.g., input.xlsx): ");
  const outputFile = await question("Enter the new file name for output (e.g., output.xlsx): ");
  const tagColumnsInput = await question("Enter the column letters for the Tags columns (e.g., 'D,F,H'): ");
  const tagColumns = tagColumnsInput.split(',').map(col => col.trim().toUpperCase());
  
  await reformatTagsColumns(inputFile, outputFile, tagColumns);
}

main();