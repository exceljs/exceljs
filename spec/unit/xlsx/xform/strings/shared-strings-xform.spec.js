'use strict';

const fs = require('fs');

const SharedStringsXform = require('../../../../../lib/xlsx/xform/strings/shared-strings-xform');
const testXformHelper = require('../test-xform-helper');

const expectations = [
  {
    title: 'Shared Strings',
    create() {
      return new SharedStringsXform();
    },
    preparedModel: require('./data/sharedStrings.json'),
    xml: fs.readFileSync(`${__dirname}/data/sharedStrings.xml`).toString(),
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('SharedStringsXform', () => {
  testXformHelper(expectations);
});
