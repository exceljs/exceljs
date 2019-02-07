'use strict';

var PageSetupPropertiesXform = require('../../../../../lib/xlsx/xform/sheet/page-setup-properties-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'fitToPage',
    create: function() { return new PageSetupPropertiesXform(); },
    preparedModel: {fitToPage: true},
    xml: '<pageSetUpPr fitToPage="1"/>',
    parsedModel: {fitToPage: true},
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('PageSetupPropertiesXform', function() {
  testXformHelper(expectations);
});
