'use strict';

var fs = require('fs');
var Sax = require('sax');
var Bluebird = require('bluebird');

var expect = require('chai').expect;

var XmlStream = require('../../../../lib/utils/xml-stream');
var FontXform = require('../../../../lib/xlsx/xform/font-xform');

var expectation = {
  model: {bold: true, size: 14, color: {argb:'FF00FF00'}, name: 'Calibri', family: 2, scheme: 'minor'},
  xml: '<font><b/><color rgb="FF00FF00"/><family val="2"/><scheme val="minor"/><sz val="14"/><name val="Calibri"/></font>'
};

describe('FontXform', function() {
  it('translate to xml', function() {
    return new Bluebird(function(resolve, reject) {
      var xform = new FontXform();
      var xmlStream = new XmlStream();
      xform.write(xmlStream, expectation.model);
      expect(xmlStream.xml).to.equal(expectation.xml);
      resolve();
    });
  });

  it('translate from xml', function() {
    return new Bluebird(function(resolve, reject) {
      var parser = Sax.createStream(true);
      var xform = new FontXform();

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