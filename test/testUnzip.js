var fs = require("fs");
var unzip = require("unzip");
var _ = require("underscore");

var ss = null;

var filename = process.argv[2];
console.log("Reading " + filename);
fs.createReadStream(filename)
    .pipe(unzip.Parse())
    .on('entry',function (entry) {
        console.log(entry.path);
    })
    .on('close', function () {
        console.log("Finished");
    })
    .on('error', function (error) {
        console.log("Error: " + error.message);
    });
    
