/**
 * Copyright (c) 2016-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var _ = require('./under-dash');

var utils = require('./utils');

// constants
var OPEN_ANGLE = '<';
var CLOSE_ANGLE = '>';
var OPEN_ANGLE_SLASH = '</';
var CLOSE_SLASH_ANGLE = '/>';
var EQUALS_QUOTE = '="';
var QUOTE = '"';
var SPACE = ' ';

function pushAttribute(xml, name, value) {
  xml.push(SPACE);
  xml.push(name);
  xml.push(EQUALS_QUOTE);
  xml.push(utils.xmlEncode(value.toString()));
  xml.push(QUOTE);
}
function pushAttributes(xml, attributes) {
  if (attributes) {
    _.each(attributes, function (value, name) {
      if (value !== undefined) {
        pushAttribute(xml, name, value);
      }
    });
  }
}

var XmlStream = module.exports = function () {
  this._xml = [];
  this._stack = [];
  this._rollbacks = [];
};

XmlStream.StdDocAttributes = {
  version: '1.0',
  encoding: 'UTF-8',
  standalone: 'yes'
};

XmlStream.prototype = {
  get tos() {
    return this._stack.length ? this._stack[this._stack.length - 1] : undefined;
  },

  openXml: function openXml(docAttributes) {
    var xml = this._xml;
    // <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    xml.push('<?xml');
    pushAttributes(xml, docAttributes);
    xml.push('?>\n');
  },

  openNode: function openNode(name, attributes) {
    var parent = this.tos;
    var xml = this._xml;
    if (parent && this.open) {
      xml.push(CLOSE_ANGLE);
    }

    this._stack.push(name);

    // start streaming node
    xml.push(OPEN_ANGLE);
    xml.push(name);
    pushAttributes(xml, attributes);
    this.leaf = true;
    this.open = true;
  },
  addAttribute: function addAttribute(name, value) {
    if (!this.open) {
      throw new Error('Cannot write attributes to node if it is not open');
    }
    pushAttribute(this._xml, name, value);
  },
  addAttributes: function addAttributes(attrs) {
    if (!this.open) {
      throw new Error('Cannot write attributes to node if it is not open');
    }
    pushAttributes(this._xml, attrs);
  },
  writeText: function writeText(text) {
    var xml = this._xml;
    if (this.open) {
      xml.push(CLOSE_ANGLE);
      this.open = false;
    }
    this.leaf = false;
    xml.push(utils.xmlEncode(text.toString()));
  },
  writeXml: function writeXml(xml) {
    if (this.open) {
      this._xml.push(CLOSE_ANGLE);
      this.open = false;
    }
    this.leaf = false;
    this._xml.push(xml);
  },
  closeNode: function closeNode() {
    var node = this._stack.pop();
    var xml = this._xml;
    if (this.leaf) {
      xml.push(CLOSE_SLASH_ANGLE);
    } else {
      xml.push(OPEN_ANGLE_SLASH);
      xml.push(node);
      xml.push(CLOSE_ANGLE);
    }
    this.open = false;
    this.leaf = false;
  },
  leafNode: function leafNode(name, attributes, text) {
    this.openNode(name, attributes);
    if (text !== undefined) {
      // zeros need to be written
      this.writeText(text);
    }
    this.closeNode();
  },

  closeAll: function closeAll() {
    while (this._stack.length) {
      this.closeNode();
    }
  },

  addRollback: function addRollback() {
    this._rollbacks.push({
      xml: this._xml.length,
      stack: this._stack.length,
      leaf: this.leaf,
      open: this.open
    });
  },
  commit: function commit() {
    this._rollbacks.pop();
  },
  rollback: function rollback() {
    var r = this._rollbacks.pop();
    if (this._xml.length > r.xml) {
      this._xml.splice(r.xml, this._xml.length - r.xml);
    }
    if (this._stack.length > r.stack) {
      this._stack.splice(r.stack, this._stack.length - r.stack);
    }
    this.leaf = r.leaf;
    this.open = r.open;
  },

  get xml() {
    this.closeAll();
    return this._xml.join('');
  }
};
//# sourceMappingURL=xml-stream.js.map
