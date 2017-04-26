'use strict';

var AutoFilterXform = require('../../../../../lib/xlsx/xform/sheet/auto-filter-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'Row and Column Adress',
    create: function() { return new AutoFilterXform(); },
    preparedModel: {from: {row: 1, column: 1}, to: {row: 1, column: 3}},
    xml: '<autoFilter ref="A1:C1"/>',
    parsedModel: {from: {row: 1, column: 1}, to: {row: 1, column: 3}},
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'String address',
    create: function() { return new AutoFilterXform(); },
    preparedModel: {from: 'A1', to: 'C1'},
    xml: '<autoFilter ref="A1:C1"/>',
    parsedModel: {from: {row: 1, column: 1}, to: {row: 1, column: 3}},
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('AutoFilterXform', function() {
  testXformHelper(expectations);
});