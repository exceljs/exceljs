'use strict';

var fs = require('fs');
var Sax = require('sax');
var Bluebird = require('bluebird');

var expect = require('chai').expect;

var XmlStream = require('../../../../lib/utils/xml-stream');
var SharedStringsXform = require('../../../../lib/xlsx/xform/shared-strings-xform');
var sharedStringsXml = fs.readFileSync(__dirname + '/data/sharedStrings.xml').toString();
var sharedStringsModel = require('./data/sharedStrings.json');


describe('SharedStringsXform', function() {
  it.only('translate to xml', function() {
    return new Bluebird(function(resolve, reject) {
      var ssx = new SharedStringsXform();
      var xmlStream = new XmlStream();
      ssx.write(xmlStream, sharedStringsModel);
      expect(xmlStream.xml).to.equal(sharedStringsXml);
    });
  });

  it('translate from xml', function() {
    return new Bluebird(function(resolve, reject) {
      var parser = Sax.createStream(true);
      var ssx = new SharedStringsXform();
      
      parser.on('opentag', function(node) {
        ssx.parseOpen(node);
      });
      parser.on('text', function(text) {
        ssx.parseText(text);
      });
      parser.on('closetag', function(name) {
        if (!ssx.parseClose(name)) {
          expect(ssx.model).to.deep.equal(sharedStringsModel);
          resolve();
        }
      });

      parser.write(sharedStringsXml);
    });
  });
});