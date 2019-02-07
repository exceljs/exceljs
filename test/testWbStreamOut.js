var Excel = require('../excel');
var utils = require('../spec/utils/index');

var filename = process.argv[2];
var styles = {
  filename: filename,
  useStyles: true
};
var wb = new Excel.stream.xlsx.WorkbookWriter(styles);
var ws = wb.addWorksheet('blort');

var style = {
  font: utils.styles.fonts.comicSansUdB16,
  alignment: utils.styles.alignments[1].alignment
};
ws.columns = [
  { header: 'A1', width: 10 },
  { header: 'B1', width: 20, style: style },
  { header: 'C1', width: 30 }
];

ws.getRow(2).font = utils.styles.fonts.broadwayRedOutline20;

ws.getCell('A2').value = 'A2';
ws.getCell('B2').value = 'B2';
ws.getCell('C2').value = 'C2';
ws.getCell('A3').value = 'A3';
ws.getCell('B3').value = 'B3';
ws.getCell('C3').value = 'C3';

wb.commit()
  .then(function() {
    console.log('Done');
    // var wb2 = new Excel.Workbook();
    // return wb2.xlsx.readFile('./wb.test2.xlsx');
  });