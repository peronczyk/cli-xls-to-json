const fs = require('fs');
const { Command } = require('commander');
const ExcelJS = require('exceljs');

// Define the command-line options
const program = new Command();
program
  .option('-f, --file <file>', 'Excel file to convert')
  .option('-k, --key <key>', 'Column to use as keys for JSON')
  .option('-v, --value <value>', 'Column to use as values for JSON')
  .parse(process.argv);

const options = program.opts();

// Check if required options are present
if (!options.file || !options.key || !options.value) {
  console.error('Missing required options!');
  program.help();
}

// Read the Excel file
const workbook = new ExcelJS.Workbook();
workbook.xlsx.readFile(options.file)
  .then(() => {
    const worksheet = workbook.getWorksheet(1);

    // Loop through the rows and create the JSON object
    const jsonObj = {};
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) { // Skip the header row
        const key = row.getCell(options.key).value;
        const value = row.getCell(options.value).value;
        if (key && value) {
          if (jsonObj[key]) {
            console.warn(` - Duplicated key detected: ${key}`);
          } else {
            jsonObj[key] = value.trim();
          }
        }
      }
    });

    // Write the JSON object to a file
    const jsonStr = JSON.stringify(jsonObj, null, 2);
    const outputFile = options.file.replace('.xlsx', '.json');
    fs.writeFileSync(outputFile, jsonStr);
    console.log(`JSON file written to ${outputFile}`);
  })
  .catch((err) => {
    console.error('Error reading Excel file:', err);
  });