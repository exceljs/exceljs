'use strict';

const SheetProtectionXform = require('../../../../../lib/xlsx/xform/sheet/sheet-protection-xform');
const testXformHelper = require('../test-xform-helper');

const expectations = [
  {
    title: 'Normal',
    create() {
      return new SheetProtectionXform();
    },
    preparedModel: {
    },
    xml:
      '<sheetProtection/>',
    parsedModel: {
    },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('SheetProtectionXform', () => {
  testXformHelper(expectations);
});
