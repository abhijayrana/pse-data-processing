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
  // Special case tag
  const specialTag = "Hard Bounce, Spam, Invalid, or Bad Email";
  
  // Check if the special tag exists in the string
  const hasSpecialTag = tags.includes(specialTag);
  
  // If the special tag exists, replace it temporarily
  if (hasSpecialTag) {
    tags = tags.replace(specialTag, "SPECIAL_TAG_PLACEHOLDER");
  }
  
  // Split the tags and format them
  let formattedTags = tags.split(',')
                          .map(tag => tag.trim())
                          .join(';');
  
  // Add semicolon at the start and end
  formattedTags = ';' + formattedTags + ';';
  
  // If the special tag was present, put it back
  if (hasSpecialTag) {
    formattedTags = formattedTags.replace("SPECIAL_TAG_PLACEHOLDER", specialTag);
  }
  
  return formattedTags;
}

async function reformatTagsColumn(inputFile, outputFile) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFile);

  const worksheet = workbook.getWorksheet(1);
  const newWorkbook = new ExcelJS.Workbook();
  const newWorksheet = newWorkbook.addWorksheet('Formatted Data');

  const tagsColumn = await question("Enter the column letter for the Tags column (e.g., 'D'): ");

  // Copy headers
  const headers = worksheet.getRow(1).values;
  newWorksheet.getRow(1).values = headers;

  // Process data rows
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // Skip header row
      const newRow = newWorksheet.getRow(rowNumber);
      
      row.eachCell((cell, colNumber) => {
        if (cell.col === worksheet.getColumn(tagsColumn).number) {
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
  await reformatTagsColumn(inputFile, outputFile);
}

main();