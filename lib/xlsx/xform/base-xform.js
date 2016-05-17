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

var Bluebird = require('bluebird');
var Sax = require('sax');

var XmlStream = require('../../utils/xml-stream');

// Base class for Xforms
var BaseXform = module.exports = function(model, name) {
};

BaseXform.prototype = {
  parse: function(parser) {
    var self = this;
    return new Bluebird(function(resolve, reject){
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
    // convenience function to get the xml of this
    return this.toXml(this.model);
  },
  
  toXml: function(model) {
    var xmlStream = new XmlStream();
    this.write(xmlStream, model);
    return xmlStream.xml;
  }
};