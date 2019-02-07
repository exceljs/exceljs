'use strict';

var PrintOptionsXform = require('../../../../../lib/xlsx/xform/sheet/print-options-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'empty',
    create: function() { return new PrintOptionsXform(); },
    preparedModel: {},
    xml: '',
    parsedModel: {},
    tests: ['render', 'renderIn']
  },
  {
    title: 'gridlines',
    create: function() { return new PrintOptionsXform(); },
    preparedModel: {showGridLines: true},
    xml: '<printOptions gridLines="1"/>',
    parsedModel: {showGridLines: true, showRowColHeaders: false, horizontalCentered: false, verticalCentered: false},
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'headers',
    create: function() { return new PrintOptionsXform(); },
    preparedModel: {showRowColHeaders: true},
    xml: '<printOptions headings="1"/>',
    parsedModel: {showGridLines: false, showRowColHeaders: true, horizontalCentered: false, verticalCentered: false},
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'centered',
    create: function() { return new PrintOptionsXform(); },
    preparedModel: {horizontalCentered: true, verticalCentered: true},
    xml: '<printOptions horizontalCentered="1" verticalCentered="1"/>',
    parsedModel: {showGridLines: false, showRowColHeaders: false, horizontalCentered: true, verticalCentered: true},
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('PrintOptionsXform', function() {
  testXformHelper(expectations);
});
