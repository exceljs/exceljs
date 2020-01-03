'use strict';

var fs = require('fs');

var SharedStringsXform = require('../../../../../lib/xlsx/xform/strings/shared-strings-xform');
var testXformHelper = require('../test-xform-helper');

var expectations = [
  {
    title: 'Shared Strings',
    create: function() { return new SharedStringsXform(); },
    preparedModel: require('./data/sharedStrings.json'),
    xml: fs.readFileSync(__dirname + '/data/sharedStrings.xml').toString(),
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('SharedStringsXform', function() {
  testXformHelper(expectations);
});
