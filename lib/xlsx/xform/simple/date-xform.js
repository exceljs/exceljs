'use strict';

const utils = require('../../../utils/utils');
const BaseXform = require('../base-xform');

const DateXform = (module.exports = function(options) {
  this.tag = options.tag;
  this.attr = options.attr;
  this.attrs = options.attrs;
  this._format =
    options.format ||
    function(dt) {
      try {
        if (Number.isNaN(dt.getTime())) return '';
        return dt.toISOString();
      } catch (e) {
        return '';
      }
    };
  this._parse =
    options.parse ||
    function(str) {
      return new Date(str);
    };
});

utils.inherits(DateXform, BaseXform, {
  render(xmlStream, model) {
    if (model) {
      xmlStream.openNode(this.tag);
      if (this.attrs) {
        xmlStream.addAttributes(this.attrs);
      }
      if (this.attr) {
        xmlStream.addAttribute(this.attr, this._format(model));
      } else {
        xmlStream.writeText(this._format(model));
      }
      xmlStream.closeNode();
    }
  },

  parseOpen(node) {
    if (node.name === this.tag) {
      if (this.attr) {
        this.model = this._parse(node.attributes[this.attr]);
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
      this.model = this._parse(this.text.join(''));
    }
    return false;
  },
});
