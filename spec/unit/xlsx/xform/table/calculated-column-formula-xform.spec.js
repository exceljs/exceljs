const testXformHelper = require('../test-xform-helper');

const CalculatedColumnFormulaXform = verquire(
  'xlsx/xform/table/calculated-column-formula-xform'
);

const expectations = [
  {
    title: 'text',
    create() {
      return new CalculatedColumnFormulaXform();
    },
    preparedModel: 'A2*2',
    xml: '<calculatedColumnFormula>A2*2</calculatedColumnFormula>',
    parsedModel: 'A2*2',
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('CalculatedColumnFormulaXform', () => {
  testXformHelper(expectations);
});
