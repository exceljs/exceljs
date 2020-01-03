'use strict';

var BlipXform = require('../../../../../lib/xlsx/xform/drawing/blip-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'full',
    create: function() { return new BlipXform(); },
    preparedModel: { rId: 'rId1' },
    xml: '<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1" cstate="print" />',
    parsedModel: { rId: 'rId1' },
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('BlipXform', function() {
  testXformHelper(expectations);
});
