# PSE Data Processor

This script formats Engagebay data in an Excel sheet, addressing issues with inconsistent formatting and Excel's automatic calculations, one usecase is cleanup for Hubspot import.


## Prerequisites

- Node.js (version 12 or higher)
- npm

## Installation

1. Clone this repository or download the script files.
2. Navigate to the script's directory in your terminal.
3. Run `npm install` to install the required dependencies.

## Usage

1. Open a terminal and navigate to the script's directory.
2. Run the script with the command: `node formatContactData.mjs`

3. Follow the prompts to enter:
- The input file name (e.g., input.xlsx)
- The output file name (e.g., output.xlsx)
- The column letter for phone numbers (e.g., B)

The script will create a new Excel file with the formatted phone numbers.

## Notes

- The script assumes that the first row of your Excel sheet contains headers.
- All phone numbers will be converted to the format: 11234567890 (no spaces or special characters).
- The original Excel file is not modified; a new file is created with the formatted data.
