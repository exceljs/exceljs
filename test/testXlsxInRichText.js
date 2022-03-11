'use strict';

/**
 * Minimum reproducible program
 */

const fs = require('fs');
const stream = require('stream');

const Excel = require('../excel');

const filename = process.argv[2];
const buffer = fs.readFileSync(filename);
const workbook = new Excel.Workbook();

const streamReadable = new stream.Readable();
streamReadable.push(buffer);
streamReadable.push(null);

workbook.xlsx
  .read(streamReadable)
  .then(allWorksheet => {
    console.log('allWorksheet: ', allWorksheet);
  })
  .catch(error => {
    console.error('something went wrong', error.stack);
  });
