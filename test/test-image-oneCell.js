const path = require('path');

const HrStopwatch = require('./utils/hr-stopwatch');

const {Workbook} = require('../lib/exceljs.nodejs');

const filename = process.argv[2];

const wb = new Workbook();
const ws = wb.addWorksheet('blort');

const imageId = wb.addImage({
  filename: path.join(__dirname, 'data/image2.png'),
  extension: 'png',
});
ws.addImage(imageId, {
  tl: {col: 0.1125, row: 0.4},
  br: {col: 2.101046875, row: 3.4},
  editAs: 'oneCell',
});

// Add another image with twoCell editAs
const imageId2 = wb.addImage({
  filename: path.join(__dirname, 'data/image2.png'),
  extension: 'png',
});
ws.addImage(imageId2, {
  tl: {col: 3.1125, row: 0.4},
  br: {col: 5.101046875, row: 3.4},
  editAs: 'twoCell',
});

const stopwatch = new HrStopwatch();
stopwatch.start();
wb.xlsx
  .writeFile(filename)
  .then(() => {
    const micros = stopwatch.microseconds;
    console.log('Done.');
    console.log('Time taken:', micros);
  })
  .catch(error => {
    console.error(error.stack);
  });
