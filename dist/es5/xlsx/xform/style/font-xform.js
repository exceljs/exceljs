/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var ColorXform = require('./color-xform');
var BooleanXform = require('../simple/boolean-xform');
var IntegerXform = require('../simple/integer-xform');
var StringXform = require('../simple/string-xform');
var UnderlineXform = require('./underline-xform');

var _ = require('../../../utils/under-dash');
var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

// Font encapsulates translation from font model to xlsx
var FontXform = module.exports = function (options) {
  this.options = options || FontXform.OPTIONS;

  this.map = {
    b: { prop: 'bold', xform: new BooleanXform({ tag: 'b', attr: 'val' }) },
    i: { prop: 'italic', xform: new BooleanXform({ tag: 'i', attr: 'val' }) },
    u: { prop: 'underline', xform: new UnderlineXform() },
    charset: { prop: 'charset', xform: new IntegerXform({ tag: 'charset', attr: 'val' }) },
    color: { prop: 'color', xform: new ColorXform() },
    condense: { prop: 'condense', xform: new BooleanXform({ tag: 'condense', attr: 'val' }) },
    extend: { prop: 'extend', xform: new BooleanXform({ tag: 'extend', attr: 'val' }) },
    family: { prop: 'family', xform: new IntegerXform({ tag: 'family', attr: 'val' }) },
    outline: { prop: 'outline', xform: new BooleanXform({ tag: 'outline', attr: 'val' }) },
    scheme: { prop: 'scheme', xform: new StringXform({ tag: 'scheme', attr: 'val' }) },
    shadow: { prop: 'shadow', xform: new BooleanXform({ tag: 'shadow', attr: 'val' }) },
    strike: { prop: 'strike', xform: new BooleanXform({ tag: 'strike', attr: 'val' }) },
    sz: { prop: 'size', xform: new IntegerXform({ tag: 'sz', attr: 'val' }) }
  };
  this.map[this.options.fontNameTag] = { prop: 'name', xform: new StringXform({ tag: this.options.fontNameTag, attr: 'val' }) };
};

FontXform.OPTIONS = {
  tagName: 'font',
  fontNameTag: 'name'
};

utils.inherits(FontXform, BaseXform, {

  get tag() {
    return this.options.tagName;
  },

  render: function render(xmlStream, model) {
    var map = this.map;

    xmlStream.openNode(this.options.tagName);
    _.each(this.map, function (defn, tag) {
      map[tag].xform.render(xmlStream, model[defn.prop]);
    });
    xmlStream.closeNode();
  },

  parseOpen: function parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    if (this.map[node.name]) {
      this.parser = this.map[node.name].xform;
      return this.parser.parseOpen(node);
    }
    switch (node.name) {
      case this.options.tagName:
        this.model = {};
        return true;
      default:
        return false;
    }
  },
  parseText: function parseText(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  },
  parseClose: function parseClose(name) {
    if (this.parser && !this.parser.parseClose(name)) {
      var item = this.map[name];
      if (this.parser.model) {
        this.model[item.prop] = this.parser.model;
      }
      this.parser = undefined;
      return true;
    }
    switch (name) {
      case this.options.tagName:
        return false;
      default:
        return true;
    }
  }
});
//# sourceMappingURL=font-xform.js.map
