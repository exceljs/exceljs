'use strict';

var AlignmentXform = require('../../../../../lib/xlsx/xform/style/alignment-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'Empty',
    create:  function() { return new AlignmentXform()},
    preparedModel: {},
    xml: '',
    tests: ['render', 'renderIn']
  },
  {
    title: 'Top Left',
    create:  function() { return new AlignmentXform()},
    preparedModel: { horizontal: 'left', vertical: 'top' },
    xml: '<alignment horizontal="left" vertical="top"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Middle Centre',
    create:  function() { return new AlignmentXform()},
    preparedModel: { horizontal: 'center', vertical: 'middle' },
    xml: '<alignment horizontal="center" vertical="center"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Bottom Right',
    create:  function() { return new AlignmentXform()},
    preparedModel: { horizontal: 'right', vertical: 'bottom'},
    xml: '<alignment horizontal="right" vertical="bottom"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Wrap Text',
    create:  function() { return new AlignmentXform()},
    preparedModel: { wrapText: true },
    xml: '<alignment wrapText="1"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Indent 1',
    create:  function() { return new AlignmentXform()},
    preparedModel: { indent: 1 },
    xml: '<alignment indent="1"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Indent 2',
    create:  function() { return new AlignmentXform()},
    preparedModel: { indent: 2 },
    xml: '<alignment indent="2"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Rotate 15',
    create:  function() { return new AlignmentXform()},
    preparedModel: { horizontal: 'right', vertical: 'bottom', textRotation: 15 },
    xml: '<alignment horizontal="right" vertical="bottom" textRotation="15"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Rotate 30',
    create:  function() { return new AlignmentXform()},
    preparedModel: { horizontal: 'right', vertical: 'bottom', textRotation: 30 },
    xml: '<alignment horizontal="right" vertical="bottom" textRotation="30"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Rotate 45',
    create:  function() { return new AlignmentXform()},
    preparedModel: { horizontal: 'right', vertical: 'bottom', textRotation: 45 },
    xml: '<alignment horizontal="right" vertical="bottom" textRotation="45"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Rotate 60',
    create:  function() { return new AlignmentXform()},
    preparedModel: { horizontal: 'right', vertical: 'bottom', textRotation: 60 },
    xml: '<alignment horizontal="right" vertical="bottom" textRotation="60"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Rotate 75',
    create:  function() { return new AlignmentXform()},
    preparedModel: { horizontal: 'right', vertical: 'bottom', textRotation: 75 },
    xml: '<alignment horizontal="right" vertical="bottom" textRotation="75"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Rotate 90',
    create:  function() { return new AlignmentXform()},
    preparedModel: { horizontal: 'right', vertical: 'bottom', textRotation: 90 },
    xml: '<alignment horizontal="right" vertical="bottom" textRotation="90"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Rotate -15',
    create:  function() { return new AlignmentXform()},
    preparedModel: { horizontal: 'right', vertical: 'bottom', textRotation: -15 },
    xml: '<alignment horizontal="right" vertical="bottom" textRotation="105"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Rotate -30',
    create:  function() { return new AlignmentXform()},
    preparedModel: { horizontal: 'right', vertical: 'bottom', textRotation: -30 },
    xml: '<alignment horizontal="right" vertical="bottom" textRotation="120"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Rotate -45',
    create:  function() { return new AlignmentXform()},
    preparedModel: { horizontal: 'right', vertical: 'bottom', textRotation: -45 },
    xml: '<alignment horizontal="right" vertical="bottom" textRotation="135"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Rotate -60',
    create:  function() { return new AlignmentXform()},
    preparedModel: { horizontal: 'right', vertical: 'bottom', textRotation: -60 },
    xml: '<alignment horizontal="right" vertical="bottom" textRotation="150"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Rotate -75',
    create:  function() { return new AlignmentXform()},
    preparedModel: { horizontal: 'right', vertical: 'bottom', textRotation: -75 },
    xml: '<alignment horizontal="right" vertical="bottom" textRotation="165"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Rotate -90',
    create:  function() { return new AlignmentXform()},
    preparedModel: { horizontal: 'right', vertical: 'bottom', textRotation: -90 },
    xml: '<alignment horizontal="right" vertical="bottom" textRotation="180"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Vertical Text',
    create:  function() { return new AlignmentXform()},
    preparedModel: { horizontal: 'right', vertical: 'bottom', textRotation: 'vertical' },
    xml: '<alignment horizontal="right" vertical="bottom" textRotation="255"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('AlignmentXform', function () {
  testXformHelper(expectations);
});
