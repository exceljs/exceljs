var fs = require('fs');
var _ = require('underscore');
var Promise = require('bluebird');

var HrStopwatch = require('./utils/hr-stopwatch');

var Excel = require('../excel');
var Workbook = Excel.Workbook;
var WorkbookWriter = Excel.stream.xlsx.WorkbookWriter;

if (process.argv[2] == 'help') {
    console.log("Usage:");
    console.log("    node testBigBookOut filename count writer strings styles");
    console.log("Where:");
    console.log("    writer is one of [stream, document]");
    console.log("    strings is one of [shared, own]");
    console.log("    styles is one of [styled, plain]");
    process.exit(0);
}

var filename = process.argv[2];
var count = process.argv.length > 3 ? parseInt(process.argv[3]) : 1000;
var writer = process.argv.length > 4 ? process.argv[4] : "stream";
var strings = process.argv.length > 5 ? process.argv[5] : "own";
var styles = process.argv.length > 6 ? process.argv[6] : "styled";

var useStream = (writer === "stream");
var useSharedStrings = (strings === "shared");
var useStyles = (styles === "styled");

var stopwatch = new HrStopwatch();
stopwatch.start();

var wb = useStream ?
    new WorkbookWriter({filename: filename, useSharedStrings: useSharedStrings, useStyles: useStyles}) :
    new Workbook();
    
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

function randomName(length) {
    length = length || 5;
    var text = [];
    var possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";

    for( var i=0; i < length; i++ )
        text.push(possible.charAt(Math.floor(Math.random() * possible.length)));

    return text.join('');
}
function randomNum(d) {
    return Math.round(Math.random()*d);
}

ws.getRow(1).font = fonts.arialBlackUI14;
var i = 0;
var deferred = Promise.defer();
var timer = setInterval(function() {
    if (i++ < count) {
        ws.addRow({
            key: i,
            name: randomName(5),
            age: randomNum(100),
            addr1: randomName(16),
            addr2: randomName(10),
            num1: randomNum(10000),
            num2: randomNum(100000),
            num3: randomNum(1000000)
        }).commit();
    } else {
        clearInterval(timer);
        deferred.resolve();
    }
}, 2);

deferred.promise.then(function() {
        return useStream ?
            wb.commit() :
            wb.xlsx.writeFile(filename);    
    })
    .then(function(){
        stopwatch.stop();        
        console.log("Done.");
        console.log("Time: " + stopwatch);
    })
    .catch(function(error) {
        console.log(error.message);
    });
