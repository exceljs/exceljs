'use strict';

const utils = require('../../../utils/utils');
const BaseXform = require('../base-xform');

const AlignmentXform = require('./alignment-xform');

// <xf numFmtId="[numFmtId]" fontId="[fontId]" fillId="[fillId]" borderId="[xf.borderId]" xfId="[xfId]">
//   Optional <alignment>
// </xf>

// Style assists translation from style model to/from xlsx
const StyleXform = (module.exports = function(options) {
  this.xfId = !!(options && options.xfId);
  this.map = {
    alignment: new AlignmentXform(),
  };
});

utils.inherits(StyleXform, BaseXform, {
  get tag() {
    return 'xf';
  },

  render(xmlStream, model) {
    xmlStream.openNode('xf', {
      numFmtId: model.numFmtId || 0,
      fontId: model.fontId || 0,
      fillId: model.fillId || 0,
      borderId: model.borderId || 0,
    });
    if (this.xfId) {
      xmlStream.addAttribute('xfId', model.xfId || 0);
    }

    if (model.numFmtId) {
      xmlStream.addAttribute('applyNumberFormat', '1');
    }
    if (model.fontId) {
      xmlStream.addAttribute('applyFont', '1');
    }
    if (model.fillId) {
      xmlStream.addAttribute('applyFill', '1');
    }
    if (model.borderId) {
      xmlStream.addAttribute('applyBorder', '1');
    }

    if (model.alignment) {
      xmlStream.addAttribute('applyAlignment', '1');
      this.map.alignment.render(xmlStream, model.alignment);
    }

    xmlStream.closeNode();
  },

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    // used during sax parsing of xml to build font object
    switch (node.name) {
      case 'xf':
        this.model = {
          numFmtId: parseInt(node.attributes.numFmtId, 10),
          fontId: parseInt(node.attributes.fontId, 10),
          fillId: parseInt(node.attributes.fillId, 10),
          borderId: parseInt(node.attributes.borderId, 10),
        };
        if (this.xfId) {
          this.model.xfId = parseInt(node.attributes.xfId, 10);
        }
        return true;
      case 'alignment':
        this.parser = this.map.alignment;
        this.parser.parseOpen(node);
        return true;
      default:
        return false;
    }
  },
  parseText(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  },
  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model.alignment = this.parser.model;
        this.parser = undefined;
      }
      return true;
    }
    return name !== 'xf';
  },
});
