'use strict';

var DataValidationsXform = require('../../../../../lib/xlsx/xform/sheet/data-validations-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'list type',
    create: function() { return new DataValidationsXform(); },
    preparedModel:{ E1: { type: 'list', allowBlank: true, showInputMessage: true, showErrorMessage: true, formulae:['Ducks']} },
    get parsedModel() { return this.preparedModel; },
    xml: '<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="E1"><formula1>Ducks</formula1></dataValidation></dataValidations>',
    tests: ['render', 'parse']
  },
  {
    title: 'whole type',
    create: function() { return new DataValidationsXform(); },
    preparedModel:{ A1: { type: 'whole', operator:'between', allowBlank: true, showInputMessage: true, showErrorMessage: true, formulae:[5, 10]} },
    get parsedModel() { return this.preparedModel; },
    xml: '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1"><formula1>5</formula1><formula2>10</formula2></dataValidation></dataValidations>',
    tests: ['render', 'parse']
  },
  {
    title: 'decimal type',
    create: function() { return new DataValidationsXform(); },
    preparedModel:{ A1: { type: 'decimal', operator:'notBetween', allowBlank: true, showInputMessage: true, showErrorMessage: true, formulae:[5, 10]} },
    get parsedModel() { return this.preparedModel; },
    xml: '<dataValidations count="1"><dataValidation type="decimal" operator="notBetween" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1"><formula1>5</formula1><formula2>10</formula2></dataValidation></dataValidations>',
    tests: ['render', 'parse']
  },
  {
    title: 'custom type',
    create: function() { return new DataValidationsXform(); },
    preparedModel:{ A1: { type: 'custom', allowBlank: true, showInputMessage: true, showErrorMessage: true, formulae:['OR(C21=5,C21=7)']} },
    get parsedModel() { return this.preparedModel; },
    xml: '<dataValidations count="1"><dataValidation type="custom" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1"><formula1>OR(C21=5,C21=7)</formula1></dataValidation></dataValidations>',
    tests: ['render', 'parse']
  }
];

describe('DataValidationsXform', function() {
  testXformHelper(expectations);
});