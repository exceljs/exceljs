'use strict';

var Sax = require('sax');
var Bluebird = require('bluebird');
var _ = require('underscore');

var chai    = require('chai');
var chaiXml = require('chai-xml');

var expect  = chai.expect;

chai.use(chaiXml);

var XmlStream = require('../../../../lib/utils/xml-stream');

function testXform(name, expectations) {
  describe(name, function () {
    _.each(expectations, function (expectation) {
      describe(expectation.title, function () {
        it('Translate to XML', function () {
          return new Bluebird(function (resolve) {
            var xform = expectation.create();
            var xmlStream = new XmlStream();
            xform.write(xmlStream, expectation.model);
            expect(xmlStream.xml).to.equal(expectation.xml);
            resolve();
          });
        });

        if (expectation.xml) {
          it('Translate to Model', function () {
            return new Bluebird(function (resolve, reject) {
              var parser = Sax.createStream(true);
              var xform = expectation.create();

              xform.parse(parser)
                .then(function (model) {
                  expect(model).to.deep.equal(expectation.model);
                  resolve();
                })
                .catch(reject);

              parser.write(expectation.xml);
            });
          });
        }
      });
    });
  });
}

module.exports = testXform;