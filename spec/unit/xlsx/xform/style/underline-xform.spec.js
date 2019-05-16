'use strict';

const UnderlineXform = require('../../../../../lib/xlsx/xform/style/underline-xform');
const testXformHelper = require('./../test-xform-helper');

const expectations = [
  {
    title: 'single',
    create() {
      return new UnderlineXform();
    },
    preparedModel: true,
    get parsedModel() {
      return this.preparedModel;
    },
    xml: '<u/>',
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'double',
    create() {
      return new UnderlineXform();
    },
    preparedModel: 'double',
    get parsedModel() {
      return this.preparedModel;
    },
    xml: '<u val="double"/>',
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'false',
    create() {
      return new UnderlineXform();
    },
    preparedModel: false,
    get parsedModel() {
      return this.preparedModel;
    },
    xml: '',
    tests: ['render', 'renderIn'],
  },
];

describe('UnderlineXform', () => {
  testXformHelper(expectations);
});
