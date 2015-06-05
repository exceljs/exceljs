var fs = require('fs');
var _ = require('underscore');
var Promise = require('bluebird');

var utils = require('./utils/utils');
var HrStopwatch = require('./utils/hr-stopwatch');

var Excel = require('../excel');
var Workbook = Excel.Workbook;
var WorkbookReader = Excel.stream.xlsx.WorkbookReader;

if (process.argv[2] == 'help') {
    console.log("Usage:");
    console.log("    node testBigBookIn filename reader passes");
    console.log("Where:");
    console.log("    reader is one of [stream, document]");
    console.log("    passes is one of [one, two]");
    process.exit(0);
}

var filename = process.argv[2];
var reader = process.argv.length > 3 ? process.argv[3] : "stream";
var passes = process.argv.length > 4 ? process.argv[4] : "one";

var useStream = (reader === "stream");
var passCounts = {
    zero: 0,
    one: 1,
    two: 2
};

var options = {
    reader: (useStream ? "stream" : "document"),
    filename: filename,
    passCount: passCounts[passes] || 0
};
console.log(JSON.stringify(options, null, "  "));

var stopwatch = new HrStopwatch();
stopwatch.start();

var cols = [3,6,7,8];
var sums = [0,0,0,0,0,0,0,0,0,0,0];
var count = 0;
function checkRow(row) {
    if (row.number > 1) {
        _.each(cols, function(col) {
            sums[col] += row.getCell(col).value;
        });
    }
    count++;
    if (count % 1000 === 0) {
        process.stdout.write("Count:" + count + "\033[0G");
    }
}
function report() {
    console.log("Count: " + count);
    console.log(sums.join(', '));
    
    stopwatch.stop();
    console.log("Time: " + stopwatch);
}

if (useStream) {
    var wb = new WorkbookReader(options);
    wb.on("worksheet", function(worksheet) {
        worksheet.on("row", checkRow);
        worksheet.on("end", report);
    });
    wb.on("entry", function(entry) {
        console.log(JSON.stringify(entry));
    });
    switch (options.passCount) {
        case 0:
            wb.read(filename, {
                entries: "emit"
            });
            break;
        case 1:
            wb.read(filename, {
                entries: "emit",
                worksheets: "emit"
            });
            break;
        case 2:
            break;
    }
} else {
    var wb = new Workbook();
    wb.xlsx.readFile(filename)
        .then(function() {
            var ws = wb.getWorksheet("blort");
            ws.eachRow(checkRow);
        })
        .then(report);
}


