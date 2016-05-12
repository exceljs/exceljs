'use strict';

var fs = require('fs');
var Sax = require('sax');
var Bluebird = require('bluebird');
var _ = require('underscore');

var expect = require('chai').expect;

var XmlStream = require('../../../../lib/utils/xml-stream');
var ColorXform = require('../../../../lib/xlsx/xform/color-xform');
var TestXform = ColorXform;

var expectations = [
  {
    title: 'RGB',
    model: {argb:'FF00FF00'},
    xml: '<color rgb="FF00FF00"/>'
  },
  {
    title: 'Theme',
    model: {theme:1},
    xml: '<color theme="1"/>'
  },
  {
    title: 'Theme with Tint',
    model: {theme:1, tint: 0.5},
    xml: '<color theme="1" tint="0.5"/>'
  },
  {
    title: 'Indexed',
    model: {indexed: 1},
    xml: '<color indexed="1"/>'
  },
  {
    title: 'Empty',
    model: undefined,
    xml: '<color auto="1"/>',
  }
];

describe('ColorXform', function() {
  _.each(expectations, function(expectation) {
    describe(expectation.title, function() {
      it('Translate to XML', function() {
        return new Bluebird(function(resolve, reject) {
          var xform = new TestXform();
          var xmlStream = new XmlStream();
          xform.write(xmlStream, expectation.model);
          expect(xmlStream.xml).to.equal(expectation.xml);
          resolve();
        });
      });

      it('Translate to Model', function() {
        return new Bluebird(function(resolve, reject) {
          var parser = Sax.createStream(true);
          var xform = new TestXform();

          xform.parse(parser)
            .then(function(model) {
              expect(model).to.deep.equal(expectation.model);
              resolve();
            })
            .catch(reject);

          parser.write(expectation.xml);
        });
      });
    });
  });
});