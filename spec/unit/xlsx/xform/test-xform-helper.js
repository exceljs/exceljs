'use strict';

var Sax = require('sax');
var Bluebird = require('bluebird');
var _ = require('lodash');

var chai    = require('chai');
var chaiXml = require('chai-xml');

var expect  = chai.expect;

chai.use(chaiXml);

var XmlStream = require('../../../../lib/utils/xml-stream');

function getExpectation(expectation, name) {
  if (!expectation.hasOwnProperty(name)) {
    throw new Error('Expectation missing required field: ' + name);
  }
  return _.cloneDeep(expectation[name]);
}

// provides boilerplate examples for the four transform steps: prepare, write,  parse and reconcile
//  prepare: model => preparedModel
//  write:  preparedModel => xml
//  parse:  xml => parsedModel
//  reconcile: parsedModel => reconciledModel

var its = {
  prepare: function(expectation) {
    it('Prepare Model', function() {
      return new Bluebird(function(resolve) {
        var model = getExpectation(expectation, 'initialModel');
        var result = getExpectation(expectation, 'preparedModel');

        var xform = expectation.create();
        xform.prepare(model, expectation.options);
        expect(model).to.deep.equal(result);
        resolve();
      });
    });
  },

  write: function(expectation) {
    it('Translate to XML', function () {
      return new Bluebird(function (resolve) {
        var model = getExpectation(expectation, 'preparedModel');
        var result = getExpectation(expectation, 'xml');

        var xform = expectation.create();
        var xmlStream = new XmlStream();
        xform.write(xmlStream, model);
        // console.log(xmlStream.xml);
        expect(xmlStream.xml).xml.to.equal(result);
        resolve();
      });
    });
  },

  parse: function(expectation) {
    it('Translate to Model', function () {
      return new Bluebird(function (resolve, reject) {
        var xml = getExpectation(expectation, 'xml');
        var result = getExpectation(expectation, 'parsedModel');

        var parser = Sax.createStream(true);
        var xform = expectation.create();

        xform.parse(parser)
          .then(function (model) {
            //console.log(JSON.stringify(model));
            expect(model).to.deep.equal(result);
            resolve();
          })
          .catch(reject);

        parser.write(xml);
      });
    });
  },

  reconcile: function(expectation) {
    it('Reconcile Model', function() {
      return new Bluebird(function(resolve) {
        var model = getExpectation(expectation, 'parsedModel');
        var result = getExpectation(expectation, 'reconciledModel');

        var xform = expectation.create();
        xform.reconcile(model, expectation.options);
        expect(model).to.deep.equal(result);
        resolve();
      });
    });
  }
};

function testXform(expectations) {
  _.each(expectations, function (expectation) {
    var tests = getExpectation(expectation, 'tests');
    describe(expectation.title, function () {
      _.each(tests, function(test) {
        its[test](expectation);
      });
    });
  });
}

module.exports = testXform;