'use strict';

const IntegerXform = require('../../../../../lib/xlsx/xform/simple/integer-xform');
const testXformHelper = require('./../test-xform-helper');

const expectations = [
  {
    title: 'five',
    create() {
      return new IntegerXform({ tag: 'integer', attr: 'val' });
    },
    preparedModel: 5,
    xml: '<integer val="5"/>',
    parsedModel: 5,
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'zero',
    create() {
      return new IntegerXform({ tag: 'integer', attr: 'val' });
    },
    preparedModel: 0,
    xml: '',
    tests: ['render', 'renderIn'],
  },
  {
    title: 'undefined',
    create() {
      return new IntegerXform({ tag: 'integer', attr: 'val' });
    },
    preparedModel: undefined,
    xml: '',
    tests: ['render', 'renderIn'],
  },
];

describe('IntegerXform', () => {
  testXformHelper(expectations);
});
