'use strict';

const Sax = require('sax');
const {expect} = require('chai');
const _ = require('../../../utils/under-dash');

const XmlStream = require('../../../../lib/utils/xml-stream');
const CompositeXform = require('../../../../lib/xlsx/xform/composite-xform');
const BooleanXform = require('../../../../lib/xlsx/xform/simple/boolean-xform');

function getExpectation(expectation, name) {
  if (!expectation.hasOwnProperty(name)) {
    throw new Error(`Expectation missing required field: ${name}`);
  }
  return _.cloneDeep(expectation[name]);
}

// ===============================================================================================================
// provides boilerplate examples for the four transform steps: prepare, render,  parse and reconcile
//  prepare: model => preparedModel
//  render:  preparedModel => xml
//  parse:  xml => parsedModel
//  reconcile: parsedModel => reconciledModel

const its = {
  prepare(expectation) {
    it('Prepare Model', () =>
      new Promise(resolve => {
        const model = getExpectation(expectation, 'initialModel');
        const result = getExpectation(expectation, 'preparedModel');

        const xform = expectation.create();
        xform.prepare(model, expectation.options);
        expect(_.cloneDeep(model)).to.deep.equal(result);
        resolve();
      }));
  },

  render(expectation) {
    it('Render to XML', () =>
      new Promise(resolve => {
        const model = getExpectation(expectation, 'preparedModel');
        const result = getExpectation(expectation, 'xml');

        const xform = expectation.create();
        const xmlStream = new XmlStream();
        xform.render(xmlStream, model, 0);
        // console.log(xmlStream.xml);
        // console.log(result);

        expect(xmlStream.xml).xml.to.equal(result);
        resolve();
      }));
  },

  'prepare-render': function(expectation) {
    // when implementation details get in the way of testing the prepared result
    it('Prepare and Render to XML', () =>
      new Promise(resolve => {
        const model = getExpectation(expectation, 'initialModel');
        const result = getExpectation(expectation, 'xml');

        const xform = expectation.create();
        const xmlStream = new XmlStream();

        xform.prepare(model, expectation.options);
        xform.render(xmlStream, model);

        expect(xmlStream.xml).xml.to.equal(result);
        resolve();
      }));
  },

  renderIn(expectation) {
    it('Render in Composite to XML ', () =>
      new Promise(resolve => {
        const model = {
          pre: true,
          child: getExpectation(expectation, 'preparedModel'),
          post: true,
        };
        const result =
          `<compy><pre/>${getExpectation(expectation, 'xml')}<post/></compy>`;

        const xform = new CompositeXform({
          tag: 'compy',
          children: [
            {
              name: 'pre',
              xform: new BooleanXform({tag: 'pre', attr: 'val'}),
            },
            {name: 'child', xform: expectation.create()},
            {
              name: 'post',
              xform: new BooleanXform({tag: 'post', attr: 'val'}),
            },
          ],
        });

        const xmlStream = new XmlStream();
        xform.render(xmlStream, model);
        // console.log(xmlStream.xml);

        expect(xmlStream.xml).xml.to.equal(result);
        resolve();
      }));
  },

  parseIn(expectation) {
    it('Parse within composite', () =>
      new Promise((resolve, reject) => {
        const xml = `<compy><pre/>${getExpectation(
          expectation,
          'xml'
        )}<post/></compy>`;
        const childXform = expectation.create();
        const result = {pre: true};
        result[childXform.tag] = getExpectation(expectation, 'parsedModel');
        result.post = true;
        const xform = new CompositeXform({
          tag: 'compy',
          children: [
            {
              name: 'pre',
              xform: new BooleanXform({tag: 'pre', attr: 'val'}),
            },
            {name: childXform.tag, xform: childXform},
            {
              name: 'post',
              xform: new BooleanXform({tag: 'post', attr: 'val'}),
            },
          ],
        });
        const parser = Sax.createStream(true);

        xform
          .parse(parser)
          .then(model => {
            // console.log('parsed Model', JSON.stringify(model));
            // console.log('expected Model', JSON.stringify(result));

            // eliminate the undefined
            const clone = _.cloneDeep(model, false);

            // console.log('result', JSON.stringify(clone));
            // console.log('expect', JSON.stringify(result));
            expect(clone).to.deep.equal(result);
            resolve();
          })
          .catch(reject);
        parser.write(xml);
      }));
  },

  parse(expectation) {
    it('Parse to Model', () =>
      new Promise((resolve, reject) => {
        const xml = getExpectation(expectation, 'xml');
        const result = getExpectation(expectation, 'parsedModel');

        const parser = Sax.createStream(true);
        const xform = expectation.create();

        xform
          .parse(parser)
          .then(model => {
            // eliminate the undefined
            const clone = _.cloneDeep(model, false);

            // console.log('result', JSON.stringify(clone));
            // console.log('expect', JSON.stringify(result));
            expect(clone).to.deep.equal(result);
            resolve();
          })
          .catch(reject);

        parser.write(xml);
      }));
  },

  reconcile(expectation) {
    it('Reconcile Model', () =>
      new Promise(resolve => {
        const model = getExpectation(expectation, 'parsedModel');
        const result = getExpectation(expectation, 'reconciledModel');

        const xform = expectation.create();
        xform.reconcile(model, expectation.options);

        // eliminate the undefined
        const clone = _.cloneDeep(model, false);

        expect(clone).to.deep.equal(result);
        resolve();
      }));
  },
};

function testXform(expectations) {
  _.each(expectations, expectation => {
    const tests = getExpectation(expectation, 'tests');
    describe(expectation.title, () => {
      _.each(tests, test => {
        its[test](expectation);
      });
    });
  });
}

module.exports = testXform;
