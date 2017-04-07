'use strict';

var expect = require('chai').expect;

var SharedFormula = require('../../../lib/utils/shared-formula');

describe('SharedFormula', function() {

  it("replaces in formulas", () => {
    assertFormulaReplace({formula: 'B2', from: 'A1', to: 'B4'}, 'C5');
    assertFormulaReplace({formula: 'B2', from: 'A1', to: 'C20'}, 'D21');
    assertFormulaReplace({formula: 'r22', from: 'A1', to: 'B4'}, 'S25');

    assertFormulaReplace({formula: 'min(loan!S45 * AD47,1)', from: 'B4', to: 'A5'}, 'min(loan!R46 * AC48,1)');

    assertFormulaReplace({formula: 'longer', from: 'B4', to: 'A5'}, 'longer');
    
  });

  var assertFormulaReplace = (test, expected, explain) => {
    expect(SharedFormula.slideFormula(
        test.formula, 
        test.from, 
        test.to
      ))
      .to.equal(expected, "when the formula '"+ test.formula + "' is copied from "+ test.from + " to " + test.to + ", the formula should become "+ expected + (explain || ''));
  }

  it("replaces in formulas but leaves the $ signed", () => {
    assertFormulaReplace({formula: '$B$2', from: 'A1', to: 'B4'}, '$B$2');
    assertFormulaReplace({formula: 'B$2', from: 'A1', to: 'C20'}, 'D$2');
    assertFormulaReplace({formula: '$R22', from: 'A1', to: 'B4'}, '$R25');

    assertFormulaReplace({formula: 'min(loan!$S45 * ZD$47,1)', from: 'B4', to: 'A5'}, 'min(loan!$S46 * ZC$47,1)');
  });

  it("replaces in formulas but leaves range names alone", () => {
    assertFormulaReplace({formula: 'XDF67', from: 'B4', to: 'A5'}, 'XDE68');
    assertFormulaReplace({formula: 'LONGER', from: 'B4', to: 'A5'}, 'LONGER');
    assertFormulaReplace({formula: 'LOGO10', from: 'B4', to: 'A5'}, 'LOGO10');
  });

  it("replaces in formulas but leaves functions and named ranges alone", () => {
    assertFormulaReplace({formula: 'a2b4', from: 'A1', to: 'B4'}, 'a2b4');
    assertFormulaReplace({formula: '2b4', from: 'A1', to: 'B4'}, '2b4');
    assertFormulaReplace({formula: 'a2b', from: 'A1', to: 'B4'}, 'a2b');
    assertFormulaReplace({formula: 'a2(2)', from: 'A1', to: 'B4'}, 'a2(2)',' - It cannot be a cell reference when followed by open paren');
    var longI25a = 'OR(INDEX(cbalSI1CodeR1,1)=C25,INDEX(cbalSI2CodeR1,1)=C25,INDEX(cbalSI3CodeR1,1)=C25,INDEX(cbalSI4CodeR1,1)=C25,INDEX(cbalSI5CodeR1,1)=C25,INDEX(cbalSI1CodeR2,1)=C25,INDEX(cbalSI2CodeR2,1)=C25,INDEX(cbalSI3CodeR2,1)=C25,INDEX(cbalSI4CodeR2,1)=C25,INDEX(cbalSI5CodeR2,1)=C25,INDEX(cbalSI1CodeR3,1)=C25,INDEX(cbalSI2CodeR3,1)=C25,INDEX(cbalSI3CodeR3,1)=C25,INDEX(cbalSI4CodeR3,1)=C25,INDEX(cbalSI5CodeR3,1)=C25)';
    var longI25b = 'OR(INDEX(cbalSI1CodeR1,1)=D27,INDEX(cbalSI2CodeR1,1)=D27,INDEX(cbalSI3CodeR1,1)=D27,INDEX(cbalSI4CodeR1,1)=D27,INDEX(cbalSI5CodeR1,1)=D27,INDEX(cbalSI1CodeR2,1)=D27,INDEX(cbalSI2CodeR2,1)=D27,INDEX(cbalSI3CodeR2,1)=D27,INDEX(cbalSI4CodeR2,1)=D27,INDEX(cbalSI5CodeR2,1)=D27,INDEX(cbalSI1CodeR3,1)=D27,INDEX(cbalSI2CodeR3,1)=D27,INDEX(cbalSI3CodeR3,1)=D27,INDEX(cbalSI4CodeR3,1)=D27,INDEX(cbalSI5CodeR3,1)=D27)';
    assertFormulaReplace({formula: longI25a, from: 'A1', to: 'B3'}, longI25b);
  });

});