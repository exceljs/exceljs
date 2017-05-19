
var Excel = require('../excel');

var filename = process.argv[2];

var wb = new Excel.Workbook();

var ws = wb.addWorksheet('printHeader');


ws.getCell('A1').value = 'This is a header row repeated on every printed page';
ws.getCell('B2').value = 'This is a header row too';

for (var i=0 ; i < 100 ; i++){
  ws.addRow(['not header row']);
};


ws.pageSetup.printTitlesRow = '1:2';


wb.xlsx.writeFile(filename)
  .then(function() {
    console.log('Done.');
  })
  .catch(function(error) {
    console.log(error.message);
  });