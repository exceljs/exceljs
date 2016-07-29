'use strict';

var fs = require('fs');

var AppXform = require('../../../../../lib/xlsx/xform/core/app-xform');
var testXformHelper = require('../test-xform-helper');

var expectations = [
  {
    title: 'app.1',
    create:  function() { return new AppXform()},
    preparedModel: require('./data/app.1.json'),
    xml: fs.readFileSync(__dirname + '/data/app.1.xml').toString().replace(/\r\n/g, '\n'),
    tests: ['render', 'renderIn']
  }
];

describe('AppXform', function () {
  testXformHelper(expectations);
});
