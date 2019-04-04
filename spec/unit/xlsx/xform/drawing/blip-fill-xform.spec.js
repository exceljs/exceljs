'use strict';

const BlipFillXform = require('../../../../../lib/xlsx/xform/drawing/blip-fill-xform');
const testXformHelper = require('./../test-xform-helper');

const expectations = [
  {
    title: 'normal',
    create() {
      return new BlipFillXform();
    },
    preparedModel: { rId: 'rId1' },
    xml:
      '<xdr:blipFill>' +
      '<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1" cstate="print" />' +
      '<a:stretch><a:fillRect /></a:stretch>' +
      '</xdr:blipFill>',
    parsedModel: { rId: 'rId1' },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('BlipFillXform', () => {
  testXformHelper(expectations);
});
