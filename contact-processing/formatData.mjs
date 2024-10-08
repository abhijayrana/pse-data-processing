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

async function processColumns(
  inputFile,
  outputFile,
  ownerColumns,
  phoneColumns,
  tagColumns
) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFile);

  const worksheet = workbook.getWorksheet(1);
  const newWorkbook = new ExcelJS.Workbook();
  const newWorksheet = newWorkbook.addWorksheet("Formatted Data");

  const headers = worksheet.getRow(1).values;
  newWorksheet.getRow(1).values = headers;

  const ownerColumnNumbers = ownerColumns.map(
    (col) => worksheet.getColumn(col).number
  );
  const phoneColumnNumbers = phoneColumns.map(
    (col) => worksheet.getColumn(col).number
  );
  const tagColumnNumbers = tagColumns.map(
    (col) => worksheet.getColumn(col).number
  );

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      const newRow = newWorksheet.getRow(rowNumber);

      row.eachCell((cell, colNumber) => {
        if (ownerColumnNumbers.includes(cell.col)) {
          let ownerValue = cell.value ? cell.value.toString() : "";
          //extract only the email (everything between the two parentheses)
          let email = ownerValue.match(/\((.*?)\)/);
          email = email ? email[1] : ""; // If match found, take the first captured group

          newRow.getCell(colNumber).value = email;
        } else if (phoneColumnNumbers.includes(cell.col)) {
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
        } else if (tagColumnNumbers.includes(cell.col)) {
          let tagsValue = cell.value ? cell.value.toString() : "";
          let formattedTags = formatTags(tagsValue);
          newRow.getCell(colNumber).value = formattedTags;
        } else {
          newRow.getCell(colNumber).value = cell.value;
        }
      });
    }
  });

  await newWorkbook.xlsx.writeFile(outputFile);
  console.log(`Formatted data saved to ${outputFile}`);
  rl.close();
}

async function main() {
  const inputFile = await question(
    "Enter the input file name (e.g., input.xlsx): "
  );
  const outputFile = await question(
    "Enter the new file name for output (e.g., output.xlsx): "
  );
  const ownerColumnsInput = await question(
    "Enter the column letters for Owner columns, separated by commas (e.g., C,E,G): "
  );
  const phoneColumnsInput = await question(
    "Enter the column letters for phone numbers, separated by commas (e.g., B,D,F): "
  );
  const tagColumnsInput = await question(
    "Enter the column letters for Tags columns, separated by commas (e.g., A,B,C): "
  );

  const ownerColumns = ownerColumnsInput
    .split(",")
    .map((col) => col.trim().toUpperCase());
  const phoneColumns = phoneColumnsInput
    .split(",")
    .map((col) => col.trim().toUpperCase());
  const tagColumns = tagColumnsInput
    .split(",")
    .map((col) => col.trim().toUpperCase());

  await processColumns(
    inputFile,
    outputFile,
    ownerColumns,
    phoneColumns,
    tagColumns
  );
}

main();
