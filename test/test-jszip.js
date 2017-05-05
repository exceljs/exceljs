'use strict';

var JSZip = require('jszip');
var Bluebird = require('bluebird');
var fs = require('fs');

var fsp = Bluebird.promisifyAll(fs);

var filename = process.argv[2];

var jsZip = new JSZip();

fsp.readFileAsync(filename)
  .then(function(data) {
    console.log('data', data)
    return jsZip.loadAsync(data);
  })
  .then(function(zip) {
    zip.forEach(function(path, entry) {
      if (!entry.dir) {
        // console.log(path, entry)
        console.log(path, entry.name, entry._data.compressedSize, entry._data.uncompressedSize)
      }
    });
  });
