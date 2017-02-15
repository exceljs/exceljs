/**
 * Copyright (c) 2015 Guyon Roche
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
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

var Sax = require('sax');

var XmlStream = require('../../utils/xml-stream');

// Base class for Xforms
var BaseXform = module.exports = function(model, name) {
};

BaseXform.prototype = {
  // ============================================================
  // Virtual Interface
  prepare:  function(model, options) {
    // optional preparation (mutation) of model so it is ready for write
  },
  render: function(xmlStream, model) {
    // convert model to xml
  },
  parseOpen:  function(node) {
    // Sax Open Node event
  },
  parseText: function(node) {
    // Sax Text event
  },
  parseClose: function(name) {
    // Sax Close Node event
  },
  reconcile: function(model, options) {
    // optional post-parse step (opposite to prepare)
  },
  
  // ============================================================
  reset: function(model) {
    // to make sure parses don't bleed to next iteration
    this.model = model;
  },
  parse: function(parser) {
    var self = this;
    return new Promise(function(resolve, reject){
      parser.on('opentag', function(node) {
        self.parseOpen(node);
      });
      parser.on('text', function(text) {
        self.parseText(text);
      });
      parser.on('closetag', function(name) {
        if (!self.parseClose(name)) {
          resolve(self.model);
        }
      });
      parser.on('end', function() {
        resolve(self.model);
      });
      parser.on('error', function(error) {
        reject(error);
      });
    });
  },
  parseStream: function(stream) {
    var parser = Sax.createStream(true, {});
    var promise = this.parse(parser);
    stream.pipe(parser);
    return promise;
  },
  
  get xml() {
    // convenience function to get the xml of this.model
    // useful for manager types that are built during the prepare phase
    return this.toXml(this.model);
  },
  
  toXml: function(model) {
    var xmlStream = new XmlStream();
    this.render(xmlStream, model);
    return xmlStream.xml;
  }
};