/**
 * Copyright (c) 2015-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../utils/utils');
var RelType = require('../../xlsx/rel-type');

var HyperlinksProxy = function HyperlinksProxy(sheetRelsWriter) {
  this.writer = sheetRelsWriter;
};
HyperlinksProxy.prototype = {
  push: function push(hyperlink) {
    this.writer.addHyperlink(hyperlink);
  }
};

var SheetRelsWriter = module.exports = function (options) {
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
      // eslint-disable-next-line no-underscore-dangle
      this._stream = this._workbook._openStream('/xl/worksheets/_rels/sheet' + this.id + '.xml.rels');
    }
    return this._stream;
  },

  get length() {
    return this._hyperlinks.length;
  },

  each: function each(fn) {
    return this._hyperlinks.forEach(fn);
  },

  get hyperlinksProxy() {
    return this._hyperlinksProxy || (this._hyperlinksProxy = new HyperlinksProxy(this));
  },
  addHyperlink: function addHyperlink(hyperlink) {
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

  addMedia: function addMedia(media) {
    return this._writeRelationship(media);
  },

  commit: function commit() {
    if (this.count) {
      // write xml utro
      this._writeClose();
      // and close stream
      this.stream.end();
    }
  },

  // ================================================================================
  _writeOpen: function _writeOpen() {
    this.stream.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">');
  },
  _writeRelationship: function _writeRelationship(relationship) {
    if (!this.count) {
      this._writeOpen();
    }

    var rId = 'rId' + ++this.count;

    if (relationship.TargetMode) {
      this.stream.write('<Relationship' + ' Id="' + rId + '"' + ' Type="' + relationship.Type + '"' + ' Target="' + utils.xmlEncode(relationship.Target) + '"' + ' TargetMode="' + relationship.TargetMode + '"' + '/>');
    } else {
      this.stream.write('<Relationship' + ' Id="' + rId + '"' + ' Type="' + relationship.Type + '"' + ' Target="' + relationship.Target + '"' + '/>');
    }

    return rId;
  },
  _writeClose: function _writeClose() {
    this.stream.write('</Relationships>');
  }
};
//# sourceMappingURL=sheet-rels-writer.js.map
