'use strict';

var BlipFillXform = require('../../../../../lib/xlsx/xform/drawing/blip-fill-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'normal',
    create: function() { return new BlipFillXform(); },
    preparedModel: { rId: 'rId1' },
    xml: '<xdr:blipFill>' +
           '<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1" cstate="print" />' +
           '<a:stretch><a:fillRect /></a:stretch>' +
         '</xdr:blipFill>',
    parsedModel: { rId: 'rId1' },
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('BlipFillXform', function() {
  testXformHelper(expectations);
});
