'use strict';

const utils = require('../../../utils/utils');
const BaseXform = require('../base-xform');

const FloatXform = (module.exports = function(options) {
  this.tag = options.tag;
  this.attr = options.attr;
  this.attrs = options.attrs;
});

utils.inherits(FloatXform, BaseXform, {
  render(xmlStream, model) {
    if (model !== undefined) {
      xmlStream.openNode(this.tag);
      if (this.attrs) {
        xmlStream.addAttributes(this.attrs);
      }
      if (this.attr) {
        xmlStream.addAttribute(this.attr, model);
      } else {
        xmlStream.writeText(model);
      }
      xmlStream.closeNode();
    }
  },

  parseOpen(node) {
    if (node.name === this.tag) {
      if (this.attr) {
        this.model = parseFloat(node.attributes[this.attr]);
      } else {
        this.text = [];
      }
    }
  },
  parseText(text) {
    if (!this.attr) {
      this.text.push(text);
    }
  },
  parseClose() {
    if (!this.attr) {
      this.model = parseFloat(this.text.join(''));
    }
    return false;
  },
});
