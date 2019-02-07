'use strict';

var MergeCellXform = require('../../../../../lib/xlsx/xform/sheet/merge-cell-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'Merge',
    create: function() { return new MergeCellXform(); },
    preparedModel: 'B2:C4',
    xml: '<mergeCell ref="B2:C4"/>',
    parsedModel: 'B2:C4',
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('MergeCellXform', function() {
  testXformHelper(expectations);
});
