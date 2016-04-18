'use strict';

var expect = require('chai').expect;

var _ = require('underscore');
var DataValidationsXform = require('../../../lib/xlsx/data-validations-xform');

var expectedXML = [
  {
    title: 'list type',
    model:{ E1: { type: 'list', allowBlank: true, showInputMessage: true, showErrorMessage: true, formulae:['Ducks']} },
    xml: '<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="E1"><formula1>Ducks</formula1></dataValidation></dataValidations>'
  },
  {
    title: 'whole type',
    model:{ A1: { type: 'whole', operator:'between', allowBlank: true, showInputMessage: true, showErrorMessage: true, formulae:['5', '10']} },
    xml: '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1"><formula1>5</formula1><formula2>10</formula2></dataValidation></dataValidations>'
  },
  {
    title: 'decimal type',
    model:{ A1: { type: 'decimal', operator:'notBetween', allowBlank: true, showInputMessage: true, showErrorMessage: true, formulae:['5', '10']} },
    xml: '<dataValidations count="1"><dataValidation type="decimal" operator="notBetween" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1"><formula1>5</formula1><formula2>10</formula2></dataValidation></dataValidations>'
  },
  {
    title: 'custom type',
    model:{ A1: { type: 'custom', allowBlank: true, showInputMessage: true, showErrorMessage: true, formulae:['OR(C21=5,C21=7)']} },
    xml: '<dataValidations count="1"><dataValidation type="custom" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1"><formula1>OR(C21=5,C21=7)</formula1></dataValidation></dataValidations>'
  }
];

describe('DataValidationsXform', function() {
  _.each(expectedXML, function(expectation) {
    describe(expectation.title, function() {
      it('translate to xml', function() {
          var dvx = new DataValidationsXform(expectation.model);
          expect(dvx.xml).to.equal(expectation.xml);
      });
    });
  });
});