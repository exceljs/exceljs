/**
 * Copyright (c) 2016 Guyon Roche
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
var _ = require('underscore');

var WorksheetProperties = module.exports = function(worksheet, model) {
  this._worksheet = worksheet;
  this.model = _.extend({}, {
    defaultRowHeight: 15,
    dyDescent: 55,
    outlineLevelCol: 0,
    outlineLevelRow: 0
  }, model);
};

WorksheetProperties.prototype = {
  // plain properties
  get tabColor() {
    return this.model.tabColor;
  },
  set tabColor(value) {
    return this.model.tabColor = value;
  },
  
  // format properties
  get defaultRowHeight() {
    return this.model.defaultRowHeight;
  },
  set defaultRowHeight(value) {
    return this.model.defaultRowHeight = value;
  },
  get dyDescent() {
    return this.model.dyDescent;
  },
  set dyDescent(value) {
    return this.model.dyDescent = value;
  },
  get outlineLevelCol() {
    return this.model.outlineLevelCol;
  },
  set outlineLevelCol(value) {
    return this.model.outlineLevelCol = value;
  },
  get outlineLevelRow() {
    return this.model.outlineLevelRow;
  },
  set outlineLevelRow(value) {
    return this.model.outlineLevelRow = value;
  }
};

