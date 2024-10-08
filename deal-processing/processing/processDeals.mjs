import ExcelJS from 'exceljs';
import readline from 'readline';

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

function question(query) {
  return new Promise((resolve) => rl.question(query, resolve));
}

async function reformatOwnerColumns(worksheet, ownerColumns) {
  const ownerColumnNumbers = ownerColumns.map(col => worksheet.getColumn(col).number);

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // Skip header row
      ownerColumnNumbers.forEach(colNumber => {
        const cell = row.getCell(colNumber);
        let ownerValue = cell.value ? cell.value.toString() : '';
        let email = ownerValue.match(/\((.*?)\)/);
        email = email ? email[1] : ''; // If match found, take the first captured group
        cell.value = email;
      });
    }
  });
}

async function reformatPhoneNumbers(worksheet, phoneNumberColumns) {
  const phoneColumnNumbers = phoneNumberColumns.map(col => worksheet.getColumn(col).number);

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // Skip header row
      phoneColumnNumbers.forEach(colNumber => {
        const cell = row.getCell(colNumber);
        let phoneNumber = cell.value;
        
        if (cell.formula) {
          phoneNumber = cell.formula;
        }
        
        phoneNumber = phoneNumber ? phoneNumber.toString() : '';
        phoneNumber = phoneNumber.replace(/[^\d-]/g, '');
        
        if (phoneNumber.startsWith('-')) {
          const parts = phoneNumber.split('-');
          if (parts.length === 2) {
            phoneNumber = parts[1] + parts[0].slice(1);
          }
        }
        
        phoneNumber = phoneNumber.replace(/-/g, '');
        
        if (phoneNumber.length === 10) {
          phoneNumber = '1' + phoneNumber;
        }
        
        cell.value = phoneNumber;
        cell.numFmt = '@';
      });
    }
  });
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

async function reformatTagsColumns(worksheet, tagColumns) {
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // Skip header row
      tagColumns.forEach(col => {
        const cell = row.getCell(col);
        let tagsValue = cell.value ? cell.value.toString() : '';
        let formattedTags = formatTags(tagsValue);
        cell.value = formattedTags;
      });
    }
  });
}

async function splitRelatedContacts(worksheet, relatedContactsColumn) {
  const newRows = [];

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // Skip header row
      const relatedContactsCell = row.getCell(relatedContactsColumn);
      const relatedContacts = relatedContactsCell.value ? relatedContactsCell.value.toString().split(',') : [];

      if (relatedContacts.length === 0) {
        newRows.push(row.values);
      } else {
        relatedContacts.forEach(contact => {
          const newRow = [...row.values];
          newRow[worksheet.getColumn(relatedContactsColumn).number] = contact.trim();
          newRows.push(newRow);
        });
      }
    }
  });

  // Clear existing rows (except header) and add new rows
  while (worksheet.rowCount > 1) {
    worksheet.spliceRows(2, 1);
  }
  worksheet.addRows(newRows);
}

async function processExcelFile(inputFile, outputFile, ownerColumns, phoneColumns, tagColumns) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFile);
  const worksheet = workbook.getWorksheet(1);

  console.log("Reformatting owner columns...");
  await reformatOwnerColumns(worksheet, ownerColumns);

  console.log("Reformatting phone numbers...");
  await reformatPhoneNumbers(worksheet, phoneColumns);

  console.log("Reformatting tags columns...");
  await reformatTagsColumns(worksheet, tagColumns);



  await workbook.xlsx.writeFile(outputFile);
  console.log(`Processed data saved to ${outputFile}`);
}

async function main() {
  const inputFile = await question("Enter the input file name (e.g., input.xlsx): ");
  const outputFile = await question("Enter the new file name for output (e.g., output.xlsx): ");

  const ownerColumnsInput = await question("Enter the column letters for Owner columns, separated by commas (e.g., C,E,G): ");
  const ownerColumns = ownerColumnsInput.split(',').map(col => col.trim().toUpperCase());

  const phoneColumnsInput = await question("Enter the column letters for phone numbers, separated by commas (e.g., B,D,F): ");
  const phoneColumns = phoneColumnsInput.split(',').map(col => col.trim().toUpperCase());

  const tagColumnsInput = await question("Enter the column letters for the Tags columns (e.g., D,F,H): ");
  const tagColumns = tagColumnsInput.split(',').map(col => col.trim().toUpperCase());


  await processExcelFile(inputFile, outputFile, ownerColumns, phoneColumns, tagColumns);

  rl.close();
}

main();