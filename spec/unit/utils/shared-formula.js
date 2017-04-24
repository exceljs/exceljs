'use strict';

var expect = require('chai').expect;

var { slideFormula } = require('../../../lib/utils/shared-formula');

describe('shared-formula', () => {
  describe('slideFormula', () => {
    const expectations = [
      { args: ['A1+1', 'A2', 'A3'], result: 'A2+1' },
      { args: ['SUM(A1:A10)', 'A11', 'B11'], result: 'SUM(B1:B10)' },
    ];
    expectations.forEach(({ args, result }) => {
      it(`${args[0]} from ${args[1]} to ${args[2]}`, () => {
        expect(slideFormula(...args)).to.equal(result);
      });
    });
  });
});