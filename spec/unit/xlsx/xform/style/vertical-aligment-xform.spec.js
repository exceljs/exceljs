'use strict';

const StringXform = require('../../../../../lib/xlsx/xform/simple/string-xform');
const testXformHelper = require('./../test-xform-helper');

const expectations = [
  {
    title: 'superscript',
    create() {
      return new StringXform({ tag: 'vertAlign', attr: 'val' });
    },
    preparedModel: 'superscript',
    xml: '<vertAlign val="superscript"/>',
    parsedModel: 'superscript',
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'subscript',
    create() {
      return new StringXform({ tag: 'vertAlign', attr: 'val' });
    },
    preparedModel: 'subscript ',
    xml: '<vertAlign val="subscript"/>',
    parsedModel: 'subscript',
    tests: ['render', 'renderIn', 'parse'],
  },
];
