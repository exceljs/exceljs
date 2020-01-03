'use strict';

var fs = require('fs');

var AppXform = require('../../../../../lib/xlsx/xform/core/app-xform');
var testXformHelper = require('../test-xform-helper');

var expectations = [
  {
    title: 'app.01',
    create: function() { return new AppXform(); },
    preparedModel: { worksheets: [{name: 'Sheet1'}] },
    xml: fs.readFileSync(__dirname + '/data/app.01.xml').toString().replace(/\r\n/g, '\n'),
    tests: ['render', 'renderIn']
  },
  {
    title: 'app.02',
    create: function() { return new AppXform(); },
    preparedModel: { worksheets: [{name: 'Sheet1'}, {name: 'Sheet2'}], company: 'Cyber Sapiens, Ltd.', manager: 'Guyon Roche' },
    xml: fs.readFileSync(__dirname + '/data/app.02.xml').toString().replace(/\r\n/g, '\n'),
    tests: ['render', 'renderIn']
  }
];

describe('AppXform', function() {
  testXformHelper(expectations);
});
