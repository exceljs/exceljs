/**
 * Copyright (c) 2016-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');
var BlipXform = require('./blip-xform');

var BlipFillXform = module.exports = function() {
  this.map = {
    'blip': new BlipXform()
  };
};

utils.inherits(BlipFillXform, BaseXform, {

  get tag() { return 'blipFill'; },

  render: function(xmlStream, model) {
    xmlStream.openNode('xdr:' + this.tag);

    this.map['blip'].render(xmlStream, model);

    // TODO: options for this + parsing
    xmlStream.openNode('a:stretch');
    xmlStream.leafNode('a:fillRect');
    xmlStream.closeNode();

    xmlStream.closeNode();
  },

  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    switch (node.name) {
      case this.tag:
        this.reset();
        break;

      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  },

  parseText: function() {
  },

  parseClose: function(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        this.model = this.map['blip'].model;
        return false;

      default:
        return true;
    }
  }
});
