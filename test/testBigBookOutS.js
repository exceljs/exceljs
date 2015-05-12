var fs = require('fs');
var _ = require('underscore');
var Promise = require('bluebird');

var Excel = require('../excel');
var Workbook = Excel.Workbook;
var WorkbookWriter = Excel.stream.xlsx.WorkbookWriter;

var filename = process.argv[2];
var count = parseInt(process.argv[3]);

var wb = new WorkbookWriter({filename: filename, useSharedStrings: true});
var ws = wb.addWorksheet("blort");

var fonts = {
    arialBlackUI14: { name: "Arial Black", family: 2, size: 14, underline: true, italic: true },
    comicSansUdB16: { name: "Comic Sans MS", family: 4, size: 8, underline: "double", bold: true }
};

ws.columns = [
    { header: "Col 1", key:"key", width: 25 },
    { header: "Col 2", key:"name", width: 32 },
    { header: "Col 3", key:"age", width: 21 },
    { header: "Col 4", key:"addr1", width: 18 },
    { header: "Col 5", key:"addr2", width: 8 },
    { header: "Col 6", key:"num1", width: 8 },
    { header: "Col 7", key:"num2", width: 8 },
    { header: "Col 8", key:"num3", width: 32, style: { font: fonts.comicSansUdB16 } }
];

function randomName() {
    var text = [];
    var possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";

    for( var i=0; i < 5; i++ )
        text.push(possible.charAt(Math.floor(Math.random() * possible.length)));

    return text.join('');
}
function randomNum(d) {
    return Math.round(Math.random()*d);
}

ws.getRow(1).font = fonts.arialBlackUI14;

for (var i = 0; i < count; i++) {
    ws.addRow({
        key: i,
        name: randomName(),
        age: randomNum(100),
        addr1: randomName(),
        addr2: randomName(),
        num1: randomNum(10000),
        num2: randomNum(100000),
        num3: randomNum(1000000)
    }).commit();
}

wb.commit()
    .then(function(){
        console.log("Done.");
    })
    .catch(function(error) {
        console.log(error.message);
    });
