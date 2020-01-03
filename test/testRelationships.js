var fs = require('fs');
var _ = require('underscore');
var Relationships = require('../lib/doc/relationships');

var testUtils = require('./utils/utils');

var mode = process.argv[2];
var out = process.argv[3];
var filename = process.argv[4];

var rt;
var tests = {
  document: function() {
    rt = new Relationships({
      filename: '',
      path: '',
      templatename: '',
      templatepath: ''
    });
    rt.addOfficeDocument();
  },
  sheet: function() {
    rt = new Relationships({
      filename: 'sheet1.xml',
      path: out + '/xl/worksheets',
      templatename: 'sheet.xml',
      templatepath: '../templates/xl/worksheets'
    });

    var lst = ['http://www.google.com', 'https://www.facebook.com'];
    _.each(lst, function(item) {
      rt.addHyperlink(item);
    });
  },
  book: function() {
    rt = new Relationships({
      filename: 'workbook.xml',
      path: out + '/xl',
      templatename: 'workbook.xml',
      templatepath: '../templates/xl'
    });

    rt.addWorksheet(1);
    rt.addTheme(1);
    rt.addStyles();
    rt.addSharedStrings();
  }
}

testUtils.cleanDir(out)
  .then(function() {
    // init test
    tests[mode]();

    // serialize
    var fullpath = out + '/' + filename;
    console.log('Writing relationship table to ' + fullpath);
    var stream = fs.createWriteStream(fullpath);
    rt.write(stream)
      .then(function(){
        stream.close();
        console.log('Done.');
      });
  })
  .catch(function(err) {
    console.log('Caught Error: ' + err);
  });

