'use strict';

const _ = require('../../../utils/under-dash');
const utils = require('../../../utils/utils');
const defaultNumFormats = require('../../defaultnumformats');

const BaseXform = require('../base-xform');

function hashDefaultFormats() {
  const hash = {};
  _.each(defaultNumFormats, (dnf, id) => {
    if (dnf.f) {
      hash[dnf.f] = parseInt(id, 10);
    }
    // at some point, add the other cultures here...
  });
  return hash;
}
const defaultFmtHash = hashDefaultFormats();

// NumFmt encapsulates translation between number format and xlsx
const NumFmtXform = (module.exports = function(id, formatCode) {
  this.id = id;
  this.formatCode = formatCode;
});

utils.inherits(
  NumFmtXform,
  BaseXform,
  {
    get tag() {
      return 'numFmt';
    },

    getDefaultFmtId(formatCode) {
      return defaultFmtHash[formatCode];
    },
    getDefaultFmtCode(numFmtId) {
      return defaultNumFormats[numFmtId] && defaultNumFormats[numFmtId].f;
    },
  },
  {
    render(xmlStream, model) {
      xmlStream.leafNode('numFmt', { numFmtId: model.id, formatCode: model.formatCode });
    },
    parseOpen(node) {
      switch (node.name) {
        case 'numFmt':
          this.model = {
            id: parseInt(node.attributes.numFmtId, 10),
            formatCode: node.attributes.formatCode.replace(/[\\](.)/g, '$1'),
          };
          return true;
        default:
          return false;
      }
    },
    parseText() {},
    parseClose() {
      return false;
    },
  }
);
