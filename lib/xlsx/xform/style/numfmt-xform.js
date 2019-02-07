/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var _ = require('../../../utils/under-dash');
var utils = require('../../../utils/utils');
var defaultNumFormats = require('../../defaultnumformats');

var BaseXform = require('../base-xform');


function hashDefaultFormats() {
  var hash = {};
  _.each(defaultNumFormats, function(dnf, id) {
    if (dnf.f) {
      hash[dnf.f] = parseInt(id, 10);
    }
    // at some point, add the other cultures here...
  });
  return hash;
}
var defaultFmtHash = hashDefaultFormats();


// NumFmt encapsulates translation between number format and xlsx
var NumFmtXform = module.exports = function(id, formatCode) {
  this.id = id;
  this.formatCode = formatCode;
};


utils.inherits(NumFmtXform, BaseXform, {

  get tag() { return 'numFmt'; },

  getDefaultFmtId: function(formatCode) {
    return defaultFmtHash[formatCode];
  },
  getDefaultFmtCode: function(numFmtId) {
    return defaultNumFormats[numFmtId] && defaultNumFormats[numFmtId].f;
  }
}, {
  render: function(xmlStream, model) {
    xmlStream.leafNode('numFmt', {numFmtId: model.id, formatCode: model.formatCode});
  },
  parseOpen: function(node) {
    switch (node.name) {
      case 'numFmt':
        this.model = {
          id: parseInt(node.attributes.numFmtId, 10),
          formatCode: node.attributes.formatCode.replace(/[\\](.)/g, '$1')
        };
        return true;
      default:
        return false;
    }
  },
  parseText: function() {
  },
  parseClose: function() {
    return false;
  }
});
