/**
 * Copyright (c) 2015-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
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
      ' Target="' + utils.xmlEncode(hyperlink.target) + '"' +
      ' TargetMode="' + hyperlink.targetMode + '"' +
      '/>');
  },
  _writeClose: function() {
    this.stream.write('</Relationships>');
  }
};