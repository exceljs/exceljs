/**
 * Copyright (c) 2016 Guyon Roche
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the 'Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 *
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
    _.each(attributes, function(value, name) {
      if (value !== undefined) {
        pushAttribute(xml, name, value);
      }
    });
  }
}

var XmlStream = module.exports = function() {
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
    return this._stack.length ? this._stack[this._stack.length-1] : undefined;
  },
  
  openXml: function(docAttributes) {
    var xml = this._xml;
    // <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    xml.push('<?xml');
    pushAttributes(xml, docAttributes);
    xml.push('?>\n');
  },
  
  openNode: function(name, attributes) {
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
  addAttribute: function(name, value) {
    if (!this.open) {
      throw new Error('Cannot write attributes to node if it is not open');
    }
    pushAttribute(this._xml, name, value);
  },
  addAttributes: function(attrs) {
    if (!this.open) {
      throw new Error('Cannot write attributes to node if it is not open');
    }
    pushAttributes(this._xml, attrs);
  },
  writeText: function(text) {
    var xml = this._xml;
    if (this.open) {
      xml.push(CLOSE_ANGLE);
      this.open = false;
    }
    this.leaf = false;
    xml.push(utils.xmlEncode(text.toString()));
  },
  writeXml: function(xml) {
    if (this.open) {
      this._xml.push(CLOSE_ANGLE);
      this.open = false;
    }
    this.leaf = false;
    this._xml.push(xml);
  },
  closeNode: function() {
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
  leafNode: function(name, attributes, text) {
    this.openNode(name, attributes);
    if (text !== undefined) { // zeros need to be written
      this.writeText(text);
    }
    this.closeNode();
  },

  closeAll: function() {
    while (this._stack.length) {
      this.closeNode();
    }
  },

  addRollback: function() {
    this._rollbacks.push({
      xml: this._xml.length,
      stack: this._stack.length,
      leaf: this.leaf,
      open: this.open
    });
  },
  commit: function() {
    this._rollbacks.pop();
  },
  rollback: function() {
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