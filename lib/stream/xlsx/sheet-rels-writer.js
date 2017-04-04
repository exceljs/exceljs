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
var RelType = require('../../xlsx/rel-type');

var HyperlinksProxy = function(sheetRelsWriter) {
  this.writer = sheetRelsWriter;
};
HyperlinksProxy.prototype = {
  push: function(hyperlink) {
    this.writer.addHyperlink(hyperlink);
  }
};

var SheetRelsWriter = module.exports = function(options) {
  // in a workbook, each sheet will have a number
  this.id = options.id;

  // count of all relationships
  this.count = 0;

  // keep record of all hyperlinks
  this._hyperlinks = [];

  this._workbook = options.workbook;
};
SheetRelsWriter.prototype = {

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

  get hyperlinksProxy() {
    return this._hyperlinksProxy ||
      (this._hyperlinksProxy = new HyperlinksProxy(this));
  },
  addHyperlink: function(hyperlink) {
    // Write to stream
    var relationship = {
      Target: hyperlink.target,
      Type: RelType.Hyperlink,
      TargetMode: 'External'
    };
    var rId = this._writeRelationship(relationship);

    // store sheet stuff for later
    this._hyperlinks.push({
      rId: rId,
      address: hyperlink.address
    });
  },

  addMedia: function(media) {
    return this._writeRelationship(media);
  },

  commit: function() {
    if (this.count) {
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
  _writeRelationship: function(relationship) {
    if (!this.count) {
      this._writeOpen();
    }

    var rId = 'rId' + ++this.count;

    if (relationship.TargetMode) {
      this.stream.write(
        '<Relationship' +
        ' Id="' + rId + '"' +
        ' Type="' + relationship.Type + '"' +
        ' Target="' + utils.xmlEncode(relationship.Target) + '"' +
        ' TargetMode="' + relationship.TargetMode + '"' +
        '/>');
    } else {
      this.stream.write(
        '<Relationship' +
        ' Id="' + rId + '"' +
        ' Type="' + relationship.Type + '"' +
        ' Target="' + relationship.Target + '"' +
        '/>');
    }


    return rId;
  },
  _writeClose: function() {
    this.stream.write('</Relationships>');
  }
};