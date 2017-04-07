/**
 * Copyright (c) 2017 Morten Ulrik Sørensen
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 *
 */
'use strict';

var colCache = require('./col-cache');

var cellRefRegex = /(([a-z\_\-0-9]*)\!)?\$?([a-z]+)\$?([1-9][0-9]*)/i;
var replacementCandidateRx = /(([a-z\_\-0-9]*)\!)?([a-z0-9_$]{2,})(\()?/ig;
var CRrx = /^(\$)?([a-z]+)(\$)?([1-9][0-9]*)$/i;

var slideFormula = function(formula, fromCell, toCell){
  var offset = colCache.decode(fromCell);
  var to = colCache.decode(toCell);
  return formula.replace(replacementCandidateRx, function(refMatch, sheet, sheetMaybe, addrPart, trailingParen) {
    if (trailingParen) {
        return refMatch;
    }
    var match = CRrx.exec(addrPart)
    if (match) {
      var colDollar = match[1];
      var colStr = match[2].toUpperCase();
      var rowDollar = match[3];
      var rowStr = match[4];
      if (colStr.length > 3 || (colStr.length == 3 && colStr > "XFD")) { // > XFD is the highest col number in excel 2007 and beyond, so this is a named range
        return refMatch;
      }
      var col = colCache.l2n(colStr);
      var row = parseInt(rowStr);
      if (!colDollar) {
        col += to.col - offset.col;
      }
      if (!rowDollar) {
        row += to.row - offset.row;
      }
      var res = (sheet ? sheet : '') + (colDollar||'') + colCache.n2l(col) + (rowDollar||'') + row;
      return res;
    } else {
      return refMatch;
    }
  });
}

module.exports = {
  slideFormula: slideFormula
};
