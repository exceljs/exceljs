var fs = require('fs');
var _ = require('underscore');
var Promise = require('bluebird');

var Workbook = require('../lib/workbook');

var filename = process.argv[2];

var wb = new Workbook();

// assuming file created by testBookOut
wb.xlsx.readFile(filename)
    .then(function() {
        var wss = wb.worksheets;
        console.log("Worksheets: " + wss.length);
        _.each(wss, function(ws) {
            console.log("Sheet " + ws.id + " - " + ws.name);
            console.log("    Dimensions: " + ws.dimensions);
        });
    })
    .catch(function(error) {
        console.log(error.message);
    });
