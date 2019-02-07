'use strict';

var OutlinePropertiesXform = require('../../../../../lib/xlsx/xform/sheet/outline-properties-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'empty',
    create: function() { return new OutlinePropertiesXform(); },
    preparedModel: {},
    xml: '',
    parsedModel: {},
    tests: ['render', 'renderIn']
  },
  {
    title: 'summaryBelow',
    create: function() { return new OutlinePropertiesXform(); },
    preparedModel: { summaryBelow: false },
    xml: '<outlinePr summaryBelow="0"/>',
    parsedModel: { summaryBelow: false },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'summaryRight',
    create: function() { return new OutlinePropertiesXform(); },
    preparedModel: { summaryRight: false },
    xml: '<outlinePr summaryRight="0"/>',
    parsedModel: { summaryRight: false },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'summaryRight',
    create: function() { return new OutlinePropertiesXform(); },
    preparedModel: { summaryBelow: true, summaryRight: false },
    xml: '<outlinePr summaryBelow="1" summaryRight="0"/>',
    parsedModel: { summaryBelow: true, summaryRight: false },
    tests: ['render', 'renderIn', 'parse']
  },
];

describe('OutlinePropertiesXform', function() {
  testXformHelper(expectations);
});
