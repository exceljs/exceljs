const testXformHelper = require('../test-xform-helper');

const BrkXform = verquire('xlsx/xform/sheet/brk-xform');

const expectations = [
  {
    title: 'one page break',
    create() {
      return new BrkXform();
    },
    initialModel: {id: 2, max: 3, min: 1, man: 1},
    preparedModel: {id: 2, max: 3, min: 1, man: 1},
    xml: '<brk id="2" max="3" min="1" man="1"/>',
    parsedModel: {id: 2, max: 3, min: 1, man: 1},
    tests: ['prepare', 'render', 'renderIn'],
  },
];

describe('BrkXform', () => {
  testXformHelper(expectations);
});
