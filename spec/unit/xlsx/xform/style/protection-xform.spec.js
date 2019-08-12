'use strict';

const ProtectionXform = require('../../../../../lib/xlsx/xform/style/protection-xform');
const testXformHelper = require('../test-xform-helper');

const expectations = [
  {
    title: 'Empty',
    create: () => new ProtectionXform(),
    preparedModel: {},
    xml: '',
    tests: ['render', 'renderIn'],
  },
  {
    title: 'Locked',
    create: () => new ProtectionXform(),
    preparedModel: {
      locked: true,
    },
    xml: '',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn'],
  },
  {
    title: 'Unlocked',
    create: () => new ProtectionXform(),
    preparedModel: {
      locked: false,
    },
    xml: '<protection locked="0"/>',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('ProtectionXform', () => {
  testXformHelper(expectations);
});
