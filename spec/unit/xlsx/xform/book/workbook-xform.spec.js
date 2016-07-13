'use strict';

var fs = require('fs');

var WorkbookXform = require('../../../../../lib/xlsx/xform/book/workbook-xform');
var testXformHelper = require('../test-xform-helper');

var expectations = [
  {
    title: 'book.1',
    create:  function() { return new WorkbookXform()},
    preparedModel: require('./data/book.1.json'),
    xml: fs.readFileSync(__dirname + '/data/book.1.xml').toString().replace(/\r\n/g, '\n'),
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'parse']
  }
];

describe('WorkbookXform', function () {
  testXformHelper(expectations);
});
