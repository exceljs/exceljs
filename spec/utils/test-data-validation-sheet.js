'use strict';

var expect = require('chai').expect;

var tools = require('./tools');

var self = module.exports = {
  dataValidations: tools.fix(require('./data/data-validations.json')),
  createDataValidations: function(type, operator) {
    var dataValidation = {
      type: type,
      operator: operator,
      allowBlank: true,
      showInputMessage: true,
      showErrorMessage: true,
      formulae: [self.dataValidations.values[type].v1]
    };
    switch (operator) {
      case 'between':
      case 'notBetween':
        dataValidation.formulae.push(self.dataValidations.values[type].v2);
        break;
      default:
        break;
    }
    return dataValidation;
  },

  addSheet: function(wb) {
    var ws = wb.addWorksheet('data-validations');

    // named list
    ws.getCell('D1').value = 'Hewie';
    ws.getCell('D1').name = 'Nephews';
    ws.getCell('E1').value = 'Dewie';
    ws.getCell('E1').name = 'Nephews';
    ws.getCell('F1').value = 'Louie';
    ws.getCell('F1').name = 'Nephews';
    ws.getCell('A1').value = tools.concatenateFormula('Named List');
    ws.getCell('B1').dataValidation = self.dataValidations.B1;

    ws.getCell('A3').value = tools.concatenateFormula('Literal List');
    ws.getCell('B3').dataValidation = self.dataValidations.B3;

    ws.getCell('D5').value = 'Tom';
    ws.getCell('E5').value = 'Dick';
    ws.getCell('F5').value = 'Harry';
    ws.getCell('A5').value = tools.concatenateFormula('Range List');
    ws.getCell('B5').dataValidation = self.dataValidations.B5;

    self.dataValidations.operators.forEach(function(operator, cIndex) {
      var col = 3 + cIndex;
      ws.getCell(7, col).value = tools.concatenateFormula(operator);
    });
    self.dataValidations.types.forEach(function(type, rIndex) {
      var row = 8 + rIndex;
      ws.getCell(row, 1).value = tools.concatenateFormula(type);
      self.dataValidations.operators.forEach(function(operator, cIndex) {
        var col = 3 + cIndex;
        ws.getCell(row, col).dataValidation = self.createDataValidations(type, operator);
      });
    });

    ws.getCell('A13').value = tools.concatenateFormula('Prompt');
    ws.getCell('B13').dataValidation = self.dataValidations.B13;

    ws.getCell('D13').value = tools.concatenateFormula('Error');
    ws.getCell('E13').dataValidation = self.dataValidations.E13;

    ws.getCell('A15').value = tools.concatenateFormula('Terse');
    ws.getCell('B15').dataValidation = self.dataValidations.B15;

    ws.getCell('A17').value = tools.concatenateFormula('Decimal');
    ws.getCell('B17').dataValidation = self.dataValidations.B17;

    ws.getCell('A19').value = tools.concatenateFormula('Any');
    ws.getCell('B19').dataValidation = self.dataValidations.B19;
  },

  checkSheet: function(wb) {
    var ws = wb.getWorksheet('data-validations');
    expect(ws).to.not.be.undefined();

    expect(ws.getCell('B1').dataValidation).to.deep.equal(self.dataValidations.B1);
    expect(ws.getCell('B3').dataValidation).to.deep.equal(self.dataValidations.B3);
    expect(ws.getCell('B5').dataValidation).to.deep.equal(self.dataValidations.B5);

    self.dataValidations.types.forEach(function(type, rIndex) {
      var row = 8 + rIndex;
      ws.getCell(row, 1).value = tools.concatenateFormula(type);
      self.dataValidations.operators.forEach(function(operator, cIndex) {
        var col = 3 + cIndex;
        expect(ws.getCell(row, col).dataValidation).to.deep.equal(self.createDataValidations(type, operator));
      });
    });

    expect(ws.getCell('B13').dataValidation).to.deep.equal(self.dataValidations.B13);
    expect(ws.getCell('E13').dataValidation).to.deep.equal(self.dataValidations.E13);
    expect(ws.getCell('B15').dataValidation).to.deep.equal(self.dataValidations.B15);
    expect(ws.getCell('B17').dataValidation).to.deep.equal(self.dataValidations.B17);
    expect(ws.getCell('B19').dataValidation).to.deep.equal(self.dataValidations.B19);
  }
};
