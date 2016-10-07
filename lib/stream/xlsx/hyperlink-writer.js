/**
 * Copyright (c) 2015 Guyon Roche
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

var utils = require('../../utils/utils');

var HyperlinkWriter = module.exports = function(options) {
  // in a workbook, each sheet will have a number
  this.id = options.id;

  // keep record of all hyperlinks
  this._hyperlinks = [];

  this._workbook = options.workbook;
};
HyperlinkWriter.prototype = {

  get stream() {
    if (!this._stream) {
      this._stream = this._workbook._openStream('/xl/worksheets/_rels/sheet' + this.id + '.xml.rels');
    }
    return this._stream;
  },

  get length() {
    return this._hyperlinks.length;
  },

  each: function(fn) {
    return this._hyperlinks.forEach(fn);
  },

  push: function(hyperlink) {
    // if first hyperlink, must open stream and write xml intro
    if (!this._hyperlinks.length) {
      this._writeOpen();
    }

    // store sheet stuff for later
    this._hyperlinks.push({
      address: hyperlink.address,
      rId: hyperlink.rId
    });

    // and write to stream
    this._writeRelationship(hyperlink);
  },

  commit: function() {
    if (this._hyperlinks.length) {
      // write xml utro
      this._writeClose();
      // and close stream
      this.stream.end();
    }
  },

  // ================================================================================
  _writeOpen: function() {
    this.stream.write(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">');
  },
  _writeRelationship: function(hyperlink) {
    this.stream.write(
      '<Relationship' +
      ' Id="' + hyperlink.rId + '"' +
      ' Type="' + hyperlink.type + '"' +
      ' Target="' + hyperlink.target + '"' +
      ' TargetMode="' + hyperlink.targetMode + '"' +
      '/>');
  },
  _writeClose: function() {
    this.stream.write('</Relationships>');
  }
};