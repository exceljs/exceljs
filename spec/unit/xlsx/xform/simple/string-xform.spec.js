'use strict';

var StringXform = require('../../../../../lib/xlsx/xform/simple/string-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'hello',
    create:  function() { return new StringXform({tag: 'string', attr: 'val'})},
    preparedModel: "Hello, World!",
    xml: '<string val="Hello, World!"/>',
    parsedModel: "Hello, World!",
    tests: ['render', 'parse']
  },
  {
    title: 'empty',
    create:  function() { return new StringXform({tag: 'string', attr: 'val'})},
    preparedModel: '',
    xml: '<string val=""/>',
    parsedModel: '',
    tests: ['render', 'parse']
  },
  {
    title: 'undefined',
    create:  function() { return new StringXform({tag: 'string', attr: 'val'})},
    preparedModel: undefined,
    xml: '',
    tests: ['render']
  }
];

describe('StringXform', function () {
  testXformHelper(expectations);
});
