'use strict';

var IntegerXform = require('../../../../../lib/xlsx/xform/simple/integer-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'five',
    create: function() { return new IntegerXform({tag: 'integer', attr: 'val'}); },
    preparedModel: 5,
    xml: '<integer val="5"/>',
    parsedModel: 5,
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'zero',
    create: function() { return new IntegerXform({tag: 'integer', attr: 'val'}); },
    preparedModel: 0,
    xml: '',
    tests: ['render', 'renderIn']
  },
  {
    title: 'undefined',
    create: function() { return new IntegerXform({tag: 'integer', attr: 'val'}); },
    preparedModel: undefined,
    xml: '',
    tests: ['render', 'renderIn']
  }
];

describe('IntegerXform', function() {
  testXformHelper(expectations);
});
