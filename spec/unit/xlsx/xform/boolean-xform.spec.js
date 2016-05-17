'use strict';

var BooleanXform = require('../../../../lib/xlsx/xform/boolean-xform');
var testXformHelper = require('./test-xform-helper');

var expectations = [
  {
    title: 'true',
    create:  function() { return new BooleanXform(undefined, {tag: 'boolean', attr: 'val'})},
    model: true,
    xml: '<boolean/>'
  },
  {
    title: 'false',
    create:  function() { return new BooleanXform(false, {tag: 'boolean', attr: 'val'})},
    model: false,
    xml: ''
  },
  {
    title: 'undefined',
    create:  function() { return new BooleanXform(undefined, {tag: 'boolean', attr: 'val'})},
    model: undefined,
    xml: ''
  }
];

testXformHelper('BooleanXform', expectations);