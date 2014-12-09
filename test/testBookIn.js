var fs = require('fs');
var _ = require('underscore');
var Promise = require('bluebird');

var Excel = require('../excel');

var filename = process.argv[2];

var wb = new Excel.Workbook();

var passed = true;
var assert = function(value, failMessage, passMessage) {
    if (!value) {
        if (failMessage) console.log(failMessage);
        passed = false;
    } else {
        if (passMessage) console.log(passMessage);        
    }
}

// assuming file created by testBookOut
wb.xlsx.readFile(filename)
    .then(function() {
        var ws = wb.getWorksheet("blort");
        
        assert(ws, "Expected to find a worksheet called blort");
        
        var cols = ws.columns;
        assert(cols[0] && (cols[0].width == 25), "Expected column width of col 1 to be 25, was " + cols[0].width);
        
        assert(ws.getCell("A2").value == 7, "Expected A2 == 7");
        assert(ws.getCell("B2").value == "Hello, World!", 'Expected B2 == "Hello, World!", was "' + ws.getCell("B2").value + '"');
        assert(ws.getCell("D2").value instanceof Date, "expected D2 to be a Date, was " + ws.getCell("D2").value);
        
        assert(ws.getCell("C5").value.formula, "Expected C5 to be a formula, was " + JSON.stringify(ws.getCell("C5").value));
        
        assert(passed, "Something went wrong", "All tests passed!");
    });


