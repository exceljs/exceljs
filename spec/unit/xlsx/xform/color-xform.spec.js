'use strict';

var fs = require('fs');
var Sax = require('sax');
var Bluebird = require('bluebird');

var expect = require('chai').expect;

var XmlStream = require('../../../../lib/utils/xml-stream');
var FontXform = require('../../../../lib/xlsx/xform/font-xform');

describe.only('FontXform', function() {
  it('translate to xml', function() {
    return new Bluebird(function(resolve, reject) {
      var xform = new FontXform();
      var xmlStream = new XmlStream();
      xform.write(xmlStream, {bold: true, size: 14, color: {argb:'ff00ff00'}, name: 'Calibri', family: 2, scheme: 'minor'});
      expect(xmlStream.xml).to.equal('<font><b/><sz val="14"/><color rgb="FF00FF00"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>');
    });
  });

  it.skip('translate from xml', function() {
    return new Bluebird(function(resolve, reject) {
      var parser = Sax.createStream(true);
      var xform = new FontXform();
      
      parser.on('opentag', function(node) {
        xform.parseOpen(node);
      });
      parser.on('text', function(text) {
        xform.parseText(text);
      });
      parser.on('closetag', function(name) {
        if (!xform.parseClose(name)) {
          expect(xform.model).to.deep.equal(expectedModel);
          resolve();
        }
      });

      parser.write(sharedStringsXml);
    });
  });
});