'use strict';

var ColXform = require('../../../../../lib/xlsx/xform/sheet/col-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'Best Fit',
    create: () => new ColXform(),
    preparedModel: {min: 2, max: 2, width: 10.15625, bestFit: true},
    get parsedModel() { return this.preparedModel; },
    xml: '<col min="2" max="2" width="10.15625" bestFit="1" customWidth="1"/>',
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Outline',
    create: () => new ColXform(),
    preparedModel: {min: 2, max: 2, width: 10.15625, bestFit: true, outlineLevel: 1, collapsed: true},
    xml: '<col min="2" max="2" width="10.15625" bestFit="1" customWidth="1" outlineLevel="1" collapsed="1"/>',
    parsedModel: {min: 2, max: 2, width: 10.15625, bestFit: true, outlineLevel: 1, collapsed: true},
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('ColXform', function() {
  testXformHelper(expectations);
});
