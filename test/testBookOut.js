var fs = require('fs');
var _ = require('underscore');
var Promise = require('bluebird');

var Workbook = require('../lib/workbook');

var filename = process.argv[2];

var wb = new Workbook()
var ws = wb.addWorksheet("blort");

var arialBlackUI14 = { name: "Arial Black", family: 2, size: 14, underline: true, italic: true };
var comicSansUdB16 = { name: "Comic Sans MS", family: 4, size: 16, underline: "double", bold: true };

ws.columns = [
    { header: "Col 1", key:"key", width: 25 },
    { header: "Col 2", key:"name", width: 25 },
    { header: "Col 3", key:"age", width: 21 },
    { header: "Col 4", key:"addr1", width: 18 },
    { header: "Col 5", key:"addr2", width: 8 },
];

ws.getCell("A2").value = 7;
ws.getCell("B2").value = "Hello, World!";
ws.getCell("B2").font = comicSansUdB16;

ws.getCell("C2").value = -5.55;
ws.getCell("C2").numFmt = "&quot;£&quot;#,##0.00;[Red]\-&quot;£&quot;#,##0.00";
ws.getCell("C2").font = arialBlackUI14;

ws.getCell("D2").value = 3.14;
ws.getCell("D2").value = new Date();
ws.getCell("D2").numFmt = "d-mmm-yyyy";
ws.getCell("D2").font = comicSansUdB16;

ws.getCell("E2").value = ["Hello", "World"].join(", ") + "!";

ws.getCell("A3").value = {text: "www.google.com", hyperlink:"http://www.google.com"};
ws.getCell("A4").value = "Boo!";
ws.getCell("C4").value = "Hoo!";
ws.mergeCells("A4", "C4");

ws.getCell("A5").value = 1;
ws.getCell("B5").value = 2;
ws.getCell("C5").value = {formula:"A5+B5", result:3};

ws.getCell("A6").value = "Hello";
ws.getCell("B6").value = "World";
ws.getCell("C6").value = {formula:'CONCATENATE(A6,", ",B6,"!")', result:'Hello, World!'};

ws.getCell("A7").value = 1;
ws.getCell("B7").value = 2;
ws.getCell("C7").value = {formula:"A7+B7"};

var now = new Date();
ws.getCell("A8").value = now;
ws.getCell("B8").value = 0;
ws.getCell("C8").value = {formula:"A8+B8", result: now};

ws.getCell("A9").value = 1.6;
ws.getCell("A9").numFmt = "# ?/?";
ws.getCell("B9").value = 1.6;
ws.getCell("B9").numFmt = "h:mm:ss";
ws.getCell("C9").value = 0.016;
ws.getCell("C9").numFmt = "0.00%";
ws.getCell("D9").value = 1.6;
ws.getCell("D9").numFmt = "[Green]#,##0 ;[Red](#,##0)";
ws.getCell("E9").value = 1.6;
ws.getCell("E9").numFmt = "#0.000";
ws.getCell("F9").value = 0.016;
ws.getCell("F9").numFmt = "# ?/?%";

ws.getCell("A10").value = "<";
ws.getCell("B10").value = ">";
ws.getCell("C10").value = "<a>";
ws.getCell("D10").value = "><";


wb.xlsx.writeFile(filename)
    .then(function(){
        console.log("Done.");
    })
    //.catch(function(error) {
    //    console.log(error.message);
    //})
