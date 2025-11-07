// Test pivot tables with XML special characters in data
// This reproduces the XML parsing error when data contains &, <, >, ", '
// Issue: XML special characters in pivot table source data cause malformed XML

/* eslint-disable */

function main(filepath) {
  const Excel = require('../lib/exceljs.nodejs.js');

  const workbook = new Excel.Workbook();

  // Create data with XML special characters
  const dataSheet = workbook.addWorksheet('Data');
  dataSheet.addRows([
    ['Company', 'Product', 'Region', 'Sales'],
    ['Johnson & Johnson', 'Drug A', 'East', 1000],
    ['Pfizer & Co', 'Drug B', 'West', 1500],
    ['BioTech <Special>', 'Drug C', 'North', 1200],
    ['PharmaCorp "Elite"', 'Drug D', 'South', 1800],
    ["Gene's Labs", 'Drug E', 'East', 2000],
    ['MedCo & Partners', 'Drug F', 'West', 2200],
    ['Ages: < 65 years', 'Drug A', 'North', 1100],
    ['Dose: 10mg > 5mg', 'Drug B', 'South', 1900],
  ]);

  // Create pivot table with data containing XML special characters
  const pivotSheet = workbook.addWorksheet('Pivot with Special Chars');
  pivotSheet.addPivotTable({
    sourceSheet: dataSheet,
    rows: ['Company', 'Product'], // Company column has XML special chars
    columns: ['Region'],
    values: ['Sales'],
    metric: 'sum',
  });

  save(workbook, filepath);
}

function save(workbook, filepath) {
  const HrStopwatch = require('./utils/hr-stopwatch');
  const stopwatch = new HrStopwatch();
  stopwatch.start();

  workbook.xlsx
    .writeFile(filepath)
    .then(() => {
      const microseconds = stopwatch.microseconds;
      console.log(
        'Done. Successfully generated pivot table with XML special characters.'
      );
      console.log('Time taken:', microseconds);
    })
    .catch(err => {
      console.error('ERROR:', err.message);
      process.exit(1);
    });
}

const [, , filepath] = process.argv;
main(filepath || 'test-pivot-xml-escape.xlsx');
