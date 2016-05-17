'use strict';

var ColorXform = require('../../../../lib/xlsx/xform/color-xform');
var testXformHelper = require('./test-xform-helper');

var expectations = [
  {
    title: 'RGB',
    create:  function() { return new ColorXform()},
    model: {argb:'FF00FF00'},
    xml: '<color rgb="FF00FF00"/>'
  },
  {
    title: 'Theme',
    create:  function() { return new ColorXform()},
    model: {theme:1},
    xml: '<color theme="1"/>'
  },
  {
    title: 'Theme with Tint',
    create:  function() { return new ColorXform()},
    model: {theme:1, tint: 0.5},
    xml: '<color theme="1" tint="0.5"/>'
  },
  {
    title: 'Indexed',
    create:  function() { return new ColorXform()},
    model: {indexed: 1},
    xml: '<color indexed="1"/>'
  },
  {
    title: 'Empty',
    create:  function() { return new ColorXform()},
    model: undefined,
    xml: '<color auto="1"/>'
  }
];

testXformHelper('ColorXform', expectations);