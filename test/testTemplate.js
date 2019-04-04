'use strict';

const Excel = require('../excel');

const inputFile = process.argv[2];
const outputFile = process.argv[3];

let passed = true;
const assert = function(value, failMessage, passMessage) {
  if (!value) {
    if (failMessage) {
      console.error(failMessage);
    }
    passed = false;
  } else if (passMessage) {
    console.log(passMessage);
  }
};

const workbook = new Excel.Workbook();
workbook.xlsx
  .readFile('./out/template.xlsx')
  .then(stream => {
    const options = {
      useSharedStrings: true,
      useStyles: true,
    };

    return stream.xlsx
      .writeFile('./out/template-out.xlsx', options)
      .then(() => {
        console.log('Done.');
      });
  })
  .catch(error => {
    console.error(error.message);
    console.error(error.stack);
  });
