var fs = require('fs');
var unzip = require('unzip2');
var _ = require('underscore');

var ss = null;

var filename = process.argv[2];
console.log('Reading ' + filename);
fs.createReadStream(filename)
  .pipe(unzip.Parse())
  .on('entry',function (entry) {
    var buf = entry.read();
    console.log(entry.path, buf.length);
  })
  .on('close', function () {
    console.log('Finished');
  })
  .on('error', function (error) {
    console.log('Error: ' + error.message);
  });

