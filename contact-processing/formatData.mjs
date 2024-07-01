import ExcelJS from "exceljs";
import readline from "readline";

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
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

  let formattedTags = tags
    .split(",")
    .map((tag) => tag.trim())
    .join(";");

  formattedTags = ";" + formattedTags + ";";

  if (hasSpecialTag) {
    formattedTags = formattedTags.replace(
      "SPECIAL_TAG_PLACEHOLDER",
      specialTag
    );
  }

  return formattedTags;
}

async function reformatTagsColumn(inputFile, outputFile, tagsColumns) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFile);

  const worksheet = workbook.getWorksheet(1);
  const newWorkbook = new ExcelJS.Workbook();
  const newWorksheet = newWorkbook.addWorksheet("Formatted Data");

  // Copy headers
  const headers = worksheet.getRow(1).values;
  newWorksheet.getRow(1).values = headers;

  // Convert column letters to column numbers
  const tagsColumnNumbers = tagsColumns.map(
    (col) => worksheet.getColumn(col).number
  );

  // Process data rows
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      // Skip header row
      const newRow = newWorksheet.getRow(rowNumber);

      row.eachCell((cell, colNumber) => {
        if (tagsColumnNumbers.includes(cell.col)) {
          let tagsValue = cell.value ? cell.value.toString() : "";
          let formattedTags = formatTags(tagsValue);
          newRow.getCell(colNumber).value = formattedTags;
        } else {
          newRow.getCell(colNumber).value = cell.value;
        }
      });
    }
  });

  // Save the new workbook
  await newWorkbook.xlsx.writeFile(outputFile);
  console.log(`Tags formatted and saved to ${outputFile}`);
}

async function reformatPhoneNumbers(inputFile, outputFile, phoneNumberColumns) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFile);

  const worksheet = workbook.getWorksheet(1);
  const newWorkbook = new ExcelJS.Workbook();
  const newWorksheet = newWorkbook.addWorksheet("Formatted Data");

  // Copy headers
  const headers = worksheet.getRow(1).values;
  newWorksheet.getRow(1).values = headers;

  // Convert column letters to column numbers
  const phoneColumnNumbers = phoneNumberColumns.map(
    (col) => worksheet.getColumn(col).number
  );

  // Process data rows
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      // Skip header row
      const newRow = newWorksheet.getRow(rowNumber);

      row.eachCell((cell, colNumber) => {
        if (phoneColumnNumbers.includes(cell.col)) {
          let phoneNumber = cell.value;

          if (cell.formula) {
            phoneNumber = cell.formula;
          }

          phoneNumber = phoneNumber ? phoneNumber.toString() : "";
          phoneNumber = phoneNumber.replace(/[^\d-]/g, "");

          if (phoneNumber.startsWith("-")) {
            const parts = phoneNumber.split("-");
            if (parts.length === 2) {
              phoneNumber = parts[1] + parts[0].slice(1);
            }
          }

          phoneNumber = phoneNumber.replace(/-/g, "");

          if (phoneNumber.length === 10) {
            phoneNumber = "1" + phoneNumber;
          }

          newRow.getCell(colNumber).value = phoneNumber;
          newRow.getCell(colNumber).numFmt = "@";
        } else {
          newRow.getCell(colNumber).value = cell.value;
        }
      });
    }
  });

  // Save the new workbook
  await newWorkbook.xlsx.writeFile(outputFile);
  console.log(`Phone numbers formatted and saved to ${outputFile}`);
}

async function reformatOwnerColumns(inputFile, outputFile, ownerColumns) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFile);

  const worksheet = workbook.getWorksheet(1);
  const newWorkbook = new ExcelJS.Workbook();
  const newWorksheet = newWorkbook.addWorksheet("Formatted Data");

  // Copy headers
  const headers = worksheet.getRow(1).values;
  newWorksheet.getRow(1).values = headers;

  // Convert column letters to column numbers
  const ownerColumnNumbers = ownerColumns.map(
    (col) => worksheet.getColumn(col).number
  );

  // Process data rows
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      // Skip header row
      const newRow = newWorksheet.getRow(rowNumber);

      row.eachCell((cell, colNumber) => {
        if (ownerColumnNumbers.includes(cell.col)) {
          let ownerValue = cell.value ? cell.value.toString() : "";
          // Extract email from between parentheses
          let email = ownerValue.match(/\((.*?)\)/);
          email = email ? email[1] : ""; // If match found, take the first captured group
          newRow.getCell(colNumber).value = email;
        } else {
          newRow.getCell(colNumber).value = cell.value;
        }
      });
    }
  });

  // Save the new workbook
  await newWorkbook.xlsx.writeFile(outputFile);
  console.log(`Owner columns formatted and saved to ${outputFile}`);
}
async function processExcelFile() {
  const inputFile = await question(
    "Enter the input file name (e.g., input.xlsx): "
  );
  const outputFile = await question(
    "Enter the new file name for output (e.g., output.xlsx): "
  );

  const tagsColumnsInput = await question(
    "Enter the column letters for Tags, separated by commas (e.g., D,E,F): "
  );
  const tagsColumns = tagsColumnsInput
    .split(",")
    .map((col) => col.trim().toUpperCase());
  const phoneColumnsInput = await question(
    "Enter the column letters for phone numbers, separated by commas (e.g., B,D,F): "
  );
  const phoneNumberColumns = phoneColumnsInput
    .split(",")
    .map((col) => col.trim().toUpperCase());
  const ownerColumnsInput = await question(
    "Enter the column letters for Owner columns, separated by commas (e.g., C,E,G): "
  );
  const ownerColumns = ownerColumnsInput
    .split(",")
    .map((col) => col.trim().toUpperCase());

  // Step 1: Format Tags
  await reformatTagsColumn(inputFile, "temp_tags.xlsx", tagsColumns);

  // Step 2: Format Phone Numbers
  await reformatPhoneNumbers(
    "temp_tags.xlsx",
    "temp_phone.xlsx",
    phoneNumberColumns
  );

  // Step 3: Format Owner Columns
  await reformatOwnerColumns("temp_phone.xlsx", outputFile, ownerColumns);

  // Clean up temporary files
  const fs = await import("fs");
  fs.unlink("temp_tags.xlsx", (err) => {
    if (err) console.error("Error deleting temp_tags.xlsx:", err);
  });
  fs.unlink("temp_phone.xlsx", (err) => {
    if (err) console.error("Error deleting temp_phone.xlsx:", err);
  });

  console.log(`Processing complete. Final output saved to ${outputFile}`);
  rl.close();
}

processExcelFile();
