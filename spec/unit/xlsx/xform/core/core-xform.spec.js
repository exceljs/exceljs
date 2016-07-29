'use strict';

var fs = require('fs');

var CoreXform = require('../../../../../lib/xlsx/xform/core/core-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'core.xml',
    create:  function() { return new CoreXform()},
    preparedModel: { creator: 'Guyon Roche', lastModifiedBy: 'Guyon Roche', created: new Date('2016-04-20T16:26:46Z'), modified: new Date('2016-05-12T06:52:49Z')},
    xml: fs.readFileSync(__dirname + '/data/core.xml').toString().replace(/\r\n/g, '\n'),
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('CoreXform', function () {
  testXformHelper(expectations);
});
