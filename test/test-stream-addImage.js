const path = require('path');
const Excel = require('../lib/exceljs.nodejs');
const HrStopwatch = require('./utils/hr-stopwatch');

const filename = process.argv[2];

const wb = new Excel.stream.xlsx.WorkbookWriter({filename});
const ws1 = wb.addWorksheet('Foo');
const imageId1 = wb.addImage({
    filename: path.join(__dirname, 'data/image2.png'),
    extension: 'png',
  });
const imageId2 = wb.addImage({
  filename:  path.join(__dirname, 'data/bubbles.jpg'),
  extension: 'jpg',
});
  
ws1.addImage(imageId1, {tl: {col: 0.25, row: 0.7},
  ext: {width: 160, height: 60}});

ws1.addImage(imageId2, 'C1:F10');

const ws2 = wb.addWorksheet('Fooo2');
ws2.addImage(imageId1, 'A1:B4');
const imageId3 = wb.addImage({
  filename: path.join(__dirname, 'data/image2.png'),
  extension: 'png',
});
ws2.addImage(imageId3, 'B4:D10');
const stopwatch = new HrStopwatch();
stopwatch.start();

wb.commit()
  .then(() => {
    const micros = stopwatch.microseconds;
    console.log('Done.');
    console.log('Time taken:', micros);
  })
  .catch(error => {
    console.log(error.message);
  });

