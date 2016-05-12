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

var _ = require('underscore');

var Alignment = require('./alignment');

// Style assists translation from style model to/from xlsx
var Style = module.exports = function() {
};

Style.prototype = {

  get numFmtId() { return this._numFmtId || 0; },
  set numFmtId(value) { return this._numFmtId = value; },

  get fontId() { return this._fontId || 0; },
  set fontId(value) { return this._fontId = value; },

  get fillId() { return this._fillId || 0; },
  set fillId(value) { return this._fillId = value; },

  get borderId() { return this._borderId || 0; },
  set borderId(value) { return this._borderId = value; },

  get xfId() { return this._xfId || 0; },
  set xfId(value) { return this._xfId = value; },

  get xml() {
    // return string containing the <xf> definition
    // <xf numFmtId="[numFmtId]" fontId="[fontId]" fillId="[fillId]" borderId="[xf.borderId]" xfId="[xfId]">
    //   Optional <alignment>
    // </xf>
    if (!this._xml) {
      var xml = [];
      var innerXml = null;
      // build xf node
      xml.push('<xf');
      xml.push(' numFmtId="' + this.numFmtId + '"');
      xml.push(' fontId="' + this.fontId + '"');
      xml.push(' fillId="' + this.fillId + '"');
      xml.push(' borderId="' + this.borderId + '"');
      xml.push(' xfId="' + this.xfId + '"');

      if (this.numFmtId) {
        xml.push(' applyNumberFormat="1"');
      }
      if (this.fontId) {
        xml.push(' applyFont="1"');
      }

      if (this.borderId) {
        xml.push(' applyBorder="1"');
      }

      if (this.alignment && this.alignment.xml) {
        xml.push(' applyAlignment="1"');
        innerXml = innerXml || [];
        innerXml.push(this.alignment.xml);
      }

      if (innerXml) {
        xml.push('>');
        innerXml.forEach(function(item) { xml.push(item); });
        xml.push('</xf>');
      } else {
        xml.push('/>');
      }

      this._xml = xml.join('');
    }
    return this._xml;
  },

  parse: function(node) {
    // used during sax parsing of xml to build font object
    switch (node.name) {
      case 'xf':
        this.numFmtId = parseInt(node.attributes.numFmtId);
        this.fontId = parseInt(node.attributes.fontId);
        this.fillId = parseInt(node.attributes.fillId);
        this.borderId = parseInt(node.attributes.borderId);
        this.xfId = parseInt(node.attributes.xfId);
        break;
      case 'alignment':
        this.alignment = new Alignment();
        this.alignment.parse(node);
        break;
    }
  }
};
