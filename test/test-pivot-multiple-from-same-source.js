// Test multiple pivot tables from same source with different field configurations
// This is a regression test for issue #5
// Bug: Second pivot table incorrectly reused field configuration from first table
// Fix: Changed splice() to slice() to prevent state mutation

/* eslint-disable */

function main(filepath) {
  const Excel = require('../lib/exceljs.nodejs.js');

  const workbook = new Excel.Workbook();

  // Create source data with multiple dimensions
  const sourceData = workbook.addWorksheet('Sales Data');
  sourceData.addRows([
    ['Region', 'Product', 'Salesperson', 'Quarter', 'Revenue', 'Units'],
    ['North', 'Widget A', 'Alice', 'Q1', 10000, 100],
    ['South', 'Widget B', 'Bob', 'Q1', 15000, 150],
    ['North', 'Widget A', 'Alice', 'Q2', 12000, 120],
    ['South', 'Widget B', 'Bob', 'Q2', 18000, 180],
    ['East', 'Widget C', 'Charlie', 'Q1', 20000, 200],
    ['West', 'Widget C', 'Diana', 'Q2', 22000, 220],
  ]);

  // First pivot table: Revenue by Region and Product
  const pivot1 = workbook.addWorksheet('Pivot 1 - Region x Product');
  pivot1.addPivotTable({
    sourceSheet: sourceData,
    rows: ['Region', 'Product'],
    columns: ['Quarter'],
    values: ['Revenue'],
    metric: 'sum',
  });

  // Second pivot table: Units by Salesperson (completely different fields)
  const pivot2 = workbook.addWorksheet('Pivot 2 - Salesperson x Quarter');
  pivot2.addPivotTable({
    sourceSheet: sourceData,
    rows: ['Salesperson'],
    columns: ['Quarter'],
    values: ['Units'],
    metric: 'sum',
  });

  // Third pivot table: Another different configuration
  const pivot3 = workbook.addWorksheet('Pivot 3 - Product x Region');
  pivot3.addPivotTable({
    sourceSheet: sourceData,
    rows: ['Product'],
    columns: ['Region'],
    values: ['Revenue'],
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
        'Done. Successfully generated 3 pivot tables from same source.'
      );
      console.log('Time taken:', microseconds);
    })
    .catch(err => {
      console.error('ERROR:', err.message);
      process.exit(1);
    });
}

const [, , filepath] = process.argv;
main(filepath || 'test-pivot-multiple-same-source.xlsx');
