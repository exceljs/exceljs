'use strict';

var WorkbookViewXform = require('../../../../../lib/xlsx/xform/book/workbook-view-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'Normal',
    create:  function() { return new WorkbookViewXform()},
    preparedModel: {},
    xml:  '<workbookView xWindow="0" yWindow="0" windowWidth="12000" windowHeight="24000"/>',
    parsedModel: {x: 0, y: 0, width: 12000, height: 24000, active: false, visibility: 'visible'},
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Hidden',
    create:  function() { return new WorkbookViewXform()},
    preparedModel: {visibility: 'hidden'},
    xml:  '<workbookView visibility="hidden" xWindow="0" yWindow="0" windowWidth="12000" windowHeight="24000"/>',
    parsedModel: {visibility: 'hidden', x: 0, y: 0, width: 12000, height: 24000, active: false},
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Active',
    create:  function() { return new WorkbookViewXform()},
    preparedModel: {active: true},
    xml:  '<workbookView xWindow="0" yWindow="0" windowWidth="12000" windowHeight="24000" activeTab="1"/>',
    parsedModel: {visibility: 'visible', x: 0, y: 0, width: 12000, height: 24000, active: true},
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('WorkbookViewXform', function () {
  testXformHelper(expectations);
});
