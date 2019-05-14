const fs = require('fs');

const HrStopwatch = require('./utils/hr-stopwatch');

const Workbook = require('../excel').Workbook;

const filenameIn = process.argv[2];
const filenameOut = process.argv[3];

const stopwatch = new HrStopwatch();
const wb = new Workbook();
stopwatch.start();
wb.xlsx
  .readFile(filenameIn)
  .then(() => wb.xlsx.writeFile(filenameOut))
  .then(() => {
    const micros = stopwatch.microseconds;
    console.log('Done.');
    console.log('Time taken:', micros);
  })
  .catch(error => {
    console.error('Error', error.stack);
  });
