'use strict';

var HyperlinkXform = require('../../../../../lib/xlsx/xform/sheet/hyperlink-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'Web Link',
    create: function() { return new HyperlinkXform(); },
    preparedModel: {address: 'B6', rId: 'rId1' },
    get parsedModel() { return this.preparedModel; },
    xml: '<hyperlink ref="B6" r:id="rId1"/>',
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('HyperlinkXform', function() {
  testXformHelper(expectations);
});
