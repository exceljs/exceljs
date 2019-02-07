'use strict';

var PageBreaksXform = require('../../../../../lib/xlsx/xform/sheet/page-breaks-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'one page break',
    create: function() { return new PageBreaksXform(); },
    initialModel: {id: 2, max: 3, min: 1, man: 1},
    preparedModel: {id: 2, max: 3, min: 1, man: 1},
    xml: '<brk id="2" max="3" min="1" man="1"/>',
    parsedModel: {id: 2, max: 3, min: 1, man: 1},
    tests: ['prepare', 'render', 'renderIn']
  }
];

describe('PageBreaksXform', function() {
  testXformHelper(expectations);
});
