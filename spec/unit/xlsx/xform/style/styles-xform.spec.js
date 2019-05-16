'use strict';

const fs = require('fs');

const StylesXform = require('../../../../../lib/xlsx/xform/style/styles-xform');
const testXformHelper = require('./../test-xform-helper');

const expectations = [
  {
    title: 'Styles',
    create() {
      return new StylesXform();
    },
    preparedModel: require('./data/styles.1.json'),
    xml: fs.readFileSync(`${__dirname}/data/styles.1.xml`).toString(),
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('StylesXform', () => {
  testXformHelper(expectations);
});
