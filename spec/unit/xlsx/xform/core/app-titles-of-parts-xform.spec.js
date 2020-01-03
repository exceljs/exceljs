'use strict';

var AppTitlesOfPartsXform = require('../../../../../lib/xlsx/xform/core/app-titles-of-parts-xform');
var testXformHelper = require('../test-xform-helper');

var expectations = [
  {
    title: 'app.01',
    create: function() { return new AppTitlesOfPartsXform(); },
    preparedModel: [{name: 'Sheet1'}],
    xml: '<TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr></vt:vector></TitlesOfParts>',
    tests: ['render', 'renderIn']
  },
  {
    title: 'app.02',
    create: function() { return new AppTitlesOfPartsXform(); },
    preparedModel: [{name: 'Sheet1'}, {name: 'Sheet2'}],
    xml: '<TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr><vt:lpstr>Sheet2</vt:lpstr></vt:vector></TitlesOfParts>',
    tests: ['render', 'renderIn']
  }
];

describe('AppTitlesOfPartsXform', function() {
  testXformHelper(expectations);
});
