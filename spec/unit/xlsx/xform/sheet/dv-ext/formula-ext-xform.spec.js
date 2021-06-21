const testXformHelper = require('../../test-xform-helper');

const FormulaExtXform = verquire('xlsx/xform/sheet/dv-ext/formula-ext-xform');

const expectations = [
  {
    title: 'formula',
    create() {
      return new FormulaExtXform();
    },
    preparedModel: ['ExternalSheet!$A$1:$A$2'],
    xml: `
        <x14:formula1>
            <xm:f>ExternalSheet!$A$1:$A$2</xm:f>
        </x14:formula1>
    `,
    parsedModel: {
      'xm:f': 'ExternalSheet!$A$1:$A$2',
    },
    tests: ['render', 'parse'],
  },
];

describe('FormulaExtXform', () => {
  testXformHelper(expectations);
});
