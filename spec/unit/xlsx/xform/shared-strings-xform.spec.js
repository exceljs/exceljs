'use strict';

var fs = require('fs');
var Sax = require('sax');
var Bluebird = require('bluebird');
var _ = require('underscore');


var chai    = require('chai');
var chaiXml = require('chai-xml');

var expect  = chai.expect;

//loads the plugin
chai.use(chaiXml);

var XmlStream = require('../../../../lib/utils/xml-stream');
var SharedStringsXform = require('../../../../lib/xlsx/xform/shared-strings-xform');
var TestXform = SharedStringsXform;

var expectations = [
  {
    title: 'Shared Strings',
    model: require('./data/sharedStrings.json'),
    xml: fs.readFileSync(__dirname + '/data/sharedStrings.xml').toString()
  }
];

describe('SharedStringsXform', function() {
  _.each(expectations, function(expectation) {
    describe(expectation.title, function() {
      it('Translate to XML', function() {
        return new Bluebird(function(resolve, reject) {
          var xform = new TestXform();
          var xmlStream = new XmlStream();
          xform.write(xmlStream, expectation.model);
          expect(xmlStream.xml).xml.to.equal(expectation.xml);
          resolve();
        });
      });

      it('Translate to Model', function() {
        return new Bluebird(function(resolve, reject) {
          var parser = Sax.createStream(true);
          var xform = new TestXform();

          xform.parse(parser)
            .then(function(model) {
              expect(model).to.deep.equal(expectation.model2 || expectation.model);
              resolve();
            })
            .catch(reject);

          parser.write(expectation.xml);
        });
      });
    });
  });
});