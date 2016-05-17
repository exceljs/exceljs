'use strict';

var UnderlineXform = require('../../../../lib/xlsx/xform/underline-xform');
var testXformHelper = require('./test-xform-helper');

var expectations = [
  {
    title: 'single',
    create:  function() { return new UnderlineXform()},
    model: true,
    xml: '<u/>'
  },
  {
    title: 'double',
    create:  function() { return new UnderlineXform()},
    model: 'double',
    xml: '<u val="double"/>'
  },
  {
    title: 'false',
    create:  function() { return new UnderlineXform(false)},
    model: false,
    xml: ''
  }
];

testXformHelper('UnderlineXform', expectations);