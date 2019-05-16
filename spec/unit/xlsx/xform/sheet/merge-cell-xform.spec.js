'use strict';

const MergeCellXform = require('../../../../../lib/xlsx/xform/sheet/merge-cell-xform');
const testXformHelper = require('./../test-xform-helper');

const expectations = [
  {
    title: 'Merge',
    create() {
      return new MergeCellXform();
    },
    preparedModel: 'B2:C4',
    xml: '<mergeCell ref="B2:C4"/>',
    parsedModel: 'B2:C4',
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('MergeCellXform', () => {
  testXformHelper(expectations);
});
