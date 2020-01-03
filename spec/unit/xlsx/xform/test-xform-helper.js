'use strict';

var Sax = require('sax');
var _ = require('../../../utils/under-dash');

var chai = require('chai');

var XmlStream = require('../../../../lib/utils/xml-stream');
var CompositeXform = require('../../../../lib/xlsx/xform/composite-xform');
var BooleanXform = require('../../../../lib/xlsx/xform/simple/boolean-xform');

var expect = chai.expect;

function getExpectation(expectation, name) {
  if (!expectation.hasOwnProperty(name)) {
    throw new Error('Expectation missing required field: ' + name);
  }
  return _.cloneDeep(expectation[name]);
}

// ===============================================================================================================
// provides boilerplate examples for the four transform steps: prepare, render,  parse and reconcile
//  prepare: model => preparedModel
//  render:  preparedModel => xml
//  parse:  xml => parsedModel
//  reconcile: parsedModel => reconciledModel

var its = {
  prepare: function(expectation) {
    it('Prepare Model', function() {
      return new Promise(function(resolve) {
        var model = getExpectation(expectation, 'initialModel');
        var result = getExpectation(expectation, 'preparedModel');

        var xform = expectation.create();
        xform.prepare(model, expectation.options);
        expect(model).to.deep.equal(result);
        resolve();
      });
    });
  },

  render: function(expectation) {
    it('Render to XML', function() {
      return new Promise(function(resolve) {
        var model = getExpectation(expectation, 'preparedModel');
        var result = getExpectation(expectation, 'xml');

        var xform = expectation.create();
        var xmlStream = new XmlStream();
        xform.render(xmlStream, model);
        // console.log(xmlStream.xml);
        
        expect(xmlStream.xml).xml.to.equal(result);
        resolve();
      });
    });
  },

  'prepare-render': function(expectation) {
    // when implementation details get in the way of testing the prepared result
    it('Prepare and Render to XML', function() {
      return new Promise(function(resolve) {
        var model = getExpectation(expectation, 'initialModel');
        var result = getExpectation(expectation, 'xml');

        var xform = expectation.create();
        var xmlStream = new XmlStream();

        xform.prepare(model, expectation.options);
        xform.render(xmlStream, model);

        expect(xmlStream.xml).xml.to.equal(result);
        resolve();
      });
    });
  },

  renderIn: function(expectation) {
    it('Render in Composite to XML ', function() {
      return new Promise(function(resolve) {
        var model = {
          pre: true,
          child: getExpectation(expectation, 'preparedModel'),
          post: true
        };
        var result =
          '<compy>' +
            '<pre/>' +
            getExpectation(expectation, 'xml') +
            '<post/>' +
          '</compy>';

        var xform = new CompositeXform({
          tag: 'compy',
          children: [
            { name: 'pre', xform: new BooleanXform({tag: 'pre', attr: 'val'}) },
            { name: 'child', xform: expectation.create() },
            { name: 'post', xform: new BooleanXform({tag: 'post', attr: 'val'}) }
          ]
        });
        
        var xmlStream = new XmlStream();
        xform.render(xmlStream, model);
        // console.log(xmlStream.xml);
        
        expect(xmlStream.xml).xml.to.equal(result);
        resolve();
      });
    });
  },
  
  parseIn: function(expectation) {
    it('Parse within composite', function() {
      return new Promise(function(resolve, reject) {
        var xml = '<compy><pre/>' + getExpectation(expectation, 'xml') + '<post/></compy>';
        var childXform = expectation.create();
        var result = {pre: true};
        result[childXform.tag] = getExpectation(expectation, 'parsedModel');
        result.post = true;
        var xform = new CompositeXform({
          tag: 'compy',
          children: [
            {name: 'pre', xform: new BooleanXform({tag: 'pre', attr: 'val'})},
            {name: childXform.tag, xform: childXform},
            {name: 'post', xform: new BooleanXform({tag: 'post', attr: 'val'})}
          ]
        });
        var parser = Sax.createStream(true);

        xform.parse(parser)
          .then(function(model) {
            // console.log('parsed Model', JSON.stringify(model));
            // console.log('expected Model', JSON.stringify(result));

            // eliminate the undefined
            var clone = _.cloneDeep(model, false);

            // console.log('result', JSON.stringify(clone));
            // console.log('expect', JSON.stringify(result));
            expect(clone).to.deep.equal(result);
            resolve();
          })
          .catch(reject);
        parser.write(xml);
      });
    });
  },

  parse: function(expectation) {
    it('Parse to Model', function() {
      return new Promise(function(resolve, reject) {
        var xml = getExpectation(expectation, 'xml');
        var result = getExpectation(expectation, 'parsedModel');

        var parser = Sax.createStream(true);
        var xform = expectation.create();

        xform.parse(parser)
          .then(function(model) {
            // eliminate the undefined
            var clone = _.cloneDeep(model, false);

            // console.log('result', JSON.stringify(clone));
            // console.log('expect', JSON.stringify(result));
            expect(clone).to.deep.equal(result);
            resolve();
          })
          .catch(reject);

        parser.write(xml);
      });
    });
  },

  reconcile: function(expectation) {
    it('Reconcile Model', function() {
      return new Promise(function(resolve) {
        var model = getExpectation(expectation, 'parsedModel');
        var result = getExpectation(expectation, 'reconciledModel');

        var xform = expectation.create();
        xform.reconcile(model, expectation.options);

        // eliminate the undefined
        var clone = _.cloneDeep(model, false);

        expect(clone).to.deep.equal(result);
        resolve();
      });
    });
  }
};

function testXform(expectations) {
  _.each(expectations, function(expectation) {
    var tests = getExpectation(expectation, 'tests');
    describe(expectation.title, function() {
      _.each(tests, function(test) {
        its[test](expectation);
      });
    });
  });
}

module.exports = testXform;
