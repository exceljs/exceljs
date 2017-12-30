/**
 * Copyright (c) 2017 Morten Ulrik Sørensen
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var colCache = require('./col-cache');

// var cellRefRegex = /(([a-z_\-0-9]*)!)?[$]?([a-z]+)[$]?([1-9][0-9]*)/i;
var replacementCandidateRx = /(([a-z_\-0-9]*)!)?([a-z0-9_$]{2,})([(])?/ig;
var CRrx = /^([$])?([a-z]+)([$])?([1-9][0-9]*)$/i;

var slideFormula = function slideFormula(formula, fromCell, toCell) {
  var offset = colCache.decode(fromCell);
  var to = colCache.decode(toCell);
  return formula.replace(replacementCandidateRx, function (refMatch, sheet, sheetMaybe, addrPart, trailingParen) {
    if (trailingParen) {
      return refMatch;
    }
    var match = CRrx.exec(addrPart);
    if (match) {
      var colDollar = match[1];
      var colStr = match[2].toUpperCase();
      var rowDollar = match[3];
      var rowStr = match[4];
      if (colStr.length > 3 || colStr.length === 3 && colStr > 'XFD') {
        // > XFD is the highest col number in excel 2007 and beyond, so this is a named range
        return refMatch;
      }
      var col = colCache.l2n(colStr);
      var row = parseInt(rowStr, 10);
      if (!colDollar) {
        col += to.col - offset.col;
      }
      if (!rowDollar) {
        row += to.row - offset.row;
      }
      var res = (sheet || '') + (colDollar || '') + colCache.n2l(col) + (rowDollar || '') + row;
      return res;
    }
    return refMatch;
  });
};

module.exports = {
  slideFormula: slideFormula
};
//# sourceMappingURL=shared-formula.js.map
