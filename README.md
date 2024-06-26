# PSE Data Processor

This project contains scripts to format Engagebay data in Excel sheets, addressing issues with inconsistent formatting and Excel's automatic calculations, useful for cleanup before Hubspot import.

## Prerequisites

- Node.js (version 12 or higher)
- npm

## Installation

1. Clone this repository or download the script files.
2. Navigate to the script's directory in your terminal.
3. Run `npm install` to install the required dependencies.

## Usage

Make sure your original excel file is in the same directory as the .mjs files

### Master Processor (`formatData.mjs`)

This script combines all three formatting functions (phone numbers, contact owners, and tags) into a single process.

To use:
1. Run the script with the command: `node formatData.mjs`
2. Follow the prompts to enter:
   - The input file name (e.g., input.xlsx)
   - The output file name (e.g., output.xlsx)
   - The column letter for the Tags column (e.g., 'D')
   - The column letters for phone numbers, separated by commas (e.g., B,D,F)
   - The column letters for Owner columns, separated by commas (e.g., C,E,G)

Notes:
- This script will process the Excel file through all three formatting steps in sequence.
- Temporary files are created during the process but are automatically deleted upon completion.

### Individual Formatters

#### 1. Phone Number Formatter (`formatPhoneNumber.mjs`)

This script reformats phone numbers in specified columns.

To use:
1. Run the script with the command: `node formatPhoneNumber.mjs`
2. Follow the prompts to enter:
   - The input file name (e.g., input.xlsx)
   - The output file name (e.g., output.xlsx)
   - The column letters for phone numbers, separated by commas (e.g., B,D,F)

Notes:
- All phone numbers will be converted to the format: 11234567890 (no spaces or special characters).
- If a phone number doesn't start with '1', it will be added automatically.

#### 2. Contact Owner Formatter (`formatContactOwner.mjs`)

This script formats the "Owner" column (or any column with eb/hubspot user) by extracting only the name and removing the email address.

To use:
1. Run the script with the command: `node formatContactOwner.mjs`
2. Follow the prompts to enter:
   - The input file name (e.g., input.xlsx)
   - The output file name (e.g., output.xlsx)
   - The column letters for the Owner column (or any column with Hubspot User as desired property type) (e.g., C,AW,BA)

Notes:
- The script will keep only the name part, removing everything after and including the opening parenthesis.
  For example, "Abhijay Rana (abhijayrana@domain.com)" will become "Abhijay Rana".

#### 3. Tags Formatter (`formatTags.mjs`)

This script formats "Tags" in Hubspot import format. Changes "taga, tagb, tagc" to ";taga;tagb;tagc;" for Hubspot multiselect property

To use:
1. Run the script with the command: `node formatContactOwner.mjs`
2. Follow the prompts to enter:
   - The input file name (e.g., input.xlsx)
   - The output file name (e.g., output.xlsx)
   - The column letter for the Tag column (e.g., I)

Notes:
- Double check pls
- Has an exception for some stuff specific to the initial PSE eb -> hs migration but can be changed in code

## General Notes

- Both scripts assume that the first row of your Excel sheet contains headers.
- The original Excel files are not modified; new files are created with the formatted data.
- You can run these scripts separately or in sequence depending on your needs.

