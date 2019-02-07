'use strict';

var ListXform = require('../../../../lib/xlsx/xform/list-xform');
var IntegerXform = require('../../../../lib/xlsx/xform/simple/integer-xform');
var testXformHelper = require('./test-xform-helper');

var expectations = [
  {
    title: 'Tagged',
    create: function() { return new ListXform({tag: 'ints', childXform: new IntegerXform({tag: 'int', attr: 'val'})}); },
    preparedModel: [1, 2, 3],
    get parsedModel() { return this.preparedModel; },
    xml: '<ints><int val="1"/><int val="2"/><int val="3"/></ints>',
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Tagged and Counted',
    create: function() { return new ListXform({tag: 'ints', count: true, childXform: new IntegerXform({tag: 'int', attr: 'val'})}); },
    preparedModel: [1, 2, 3],
    get parsedModel() { return this.preparedModel; },
    xml: '<ints count="3"><int val="1"/><int val="2"/><int val="3"/></ints>',
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('ListXform', function() {
  testXformHelper(expectations);
});
