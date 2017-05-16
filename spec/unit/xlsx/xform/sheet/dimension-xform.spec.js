'use strict';

var DimensionXform = require('../../../../../lib/xlsx/xform/sheet/dimension-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'Dimension',
    create: function() { return new DimensionXform(); },
    preparedModel: 'A1:F5',
    get parsedModel() { return this.preparedModel; },
    xml: '<dimension ref="A1:F5"/>',
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('DimensionXform', function() {
  testXformHelper(expectations);
});
