var fs = require('fs');
var _ = require('underscore');
var Promise = require('bluebird');

var Workbook = require('../lib/workbook');

var filename = process.argv[2];

var wb = new Workbook()
var ws = wb.addWorksheet("blort");

ws.columns = [
    { header: "Col 1", key:"key", width: 25 },
    { header: "Col 2", key:"name", width: 25 },
    { header: "Col 3", key:"age", width: 21 },
    { header: "Col 4", key:"addr1", width: 18 },
    { header: "Col 5", key:"addr2", width: 8 },
];

ws.getCell("A2").value = 7;
ws.getCell("B2").value = "Hello, World!";
ws.getCell("C2").value = 5;
ws.getCell("D2").value = 3.14;
ws.getCell("D2").value = new Date();
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


wb.xlsx.writeFile(filename)
    .then(function(){
        console.log("Done.");
    })
    //.catch(function(error) {
    //    console.log(error.message);
    //})
