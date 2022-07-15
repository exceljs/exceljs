const testXformHelper = require('../test-xform-helper');

const BooleanXform = verquire('xlsx/xform/simple/boolean-xform');

const expectations = [
  {
    title: 'true',
    create() {
      return new BooleanXform({tag: 'boolean', attr: 'val'});
    },
    preparedModel: true,
    get parsedModel() {
      return this.preparedModel;
    },
    xml: '<boolean/>',
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'zero',
    create() {
      return new BooleanXform({tag: 'boolean', attr: 'val'});
    },
    preparedModel: false,
    xml: '<boolean val="0"/>',
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'one',
    create() {
      return new BooleanXform({tag: 'boolean', attr: 'val'});
    },
    preparedModel: true,
    xml: '<boolean val="1"/>',
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'undefined',
    create() {
      return new BooleanXform({tag: 'boolean', attr: 'val'});
    },
    preparedModel: undefined,
    xml: '',
    tests: ['render', 'renderIn'],
  },
];

describe('BooleanXform', () => {
  testXformHelper(expectations);
});
