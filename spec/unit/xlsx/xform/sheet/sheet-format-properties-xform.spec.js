'use strict';

var SheetFormatPropertiesXform = require('../../../../../lib/xlsx/xform/sheet/sheet-format-properties-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'full',
    create: function() { return new SheetFormatPropertiesXform(); },
    preparedModel: {defaultRowHeight: 14.4, dyDescent: 0.55, outlineLevelRow: 5, outlineLevelCol: 2},
    xml: '<sheetFormatPr defaultRowHeight="14.4" outlineLevelRow="5" outlineLevelCol="2" x14ac:dyDescent="0.55"/>',
    parsedModel: {defaultRowHeight: 14.4, dyDescent: 0.55, outlineLevelRow: 5, outlineLevelCol: 2},
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'default',
    create: function() { return new SheetFormatPropertiesXform(); },
    preparedModel: {defaultRowHeight: 14.4, dyDescent: 0.55},
    xml: '<sheetFormatPr defaultRowHeight="14.4" x14ac:dyDescent="0.55"/>',
    parsedModel: {defaultRowHeight: 14.4, dyDescent: 0.55, outlineLevelRow: 0, outlineLevelCol: 0},
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('SheetFormatPropertiesXform', function() {
  testXformHelper(expectations);
});
