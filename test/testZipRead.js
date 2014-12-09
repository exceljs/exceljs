var fs = require('fs');
var unzip = require("unzip");
var Sax = require("sax");

var filename = process.argv[2];

var zipEntries = {};
var zipStream = unzip.Parse();
zipStream.on('entry',function (entry) {
    console.log(entry.path)
    //zipEntries[entry.path] = entry;
    entry.autodrain();
});
zipStream.on('close', function () {
    console.log("zip close")
});

var readStream = fs.createReadStream(filename);
readStream.pipe(zipStream);