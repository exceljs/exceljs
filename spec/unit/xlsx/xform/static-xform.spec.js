'use strict';

var StaticXform = require('../../../../lib/xlsx/xform/static-xform');
var testXformHelper = require('./test-xform-helper');

// var model = {
//   tag: 'name',
//   $: {attr: 'value'},
//   c: [
//     { tag: 'child' }
//   ],
//   t: 'some text'
// };

var expectations = [
  {
    title: 'Leaf',
    create: function() { return new StaticXform({tag: 'root', $: {attr: 'val'}}); },
    preparedModel: undefined,
    get parsedModel() { return this.preparedModel; },
    xml: '<root attr="val"/>',
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Nested',
    create: function() { return new StaticXform({tag: 'root', $: {attr: 'val'}, c: [{tag: 'child1', $: {attr: 5}}, {tag: 'child2', $: {attr: true}}]}); },
    preparedModel: undefined,
    get parsedModel() { return this.preparedModel; },
    xml: '<root attr="val"><child1 attr="5"/><child2 attr="true"/></root>',
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Texted',
    create: function() { return new StaticXform({tag: 'root', $: {attr: 'val'}, c: [{tag: 'child1', $: {attr: 5}, t: 'Hello, World!'}]}); },
    preparedModel: undefined,
    get parsedModel() { return this.preparedModel; },
    xml: '<root attr="val"><child1 attr="5">Hello, World!</child1></root>',
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('StaticXform', function() {
  testXformHelper(expectations);
});
