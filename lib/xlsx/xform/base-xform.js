/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var Sax = require('sax');
var PromishLib = require('../../utils/promish');

var XmlStream = require('../../utils/xml-stream');

// Base class for Xforms
var BaseXform = module.exports = function(/* model, name */) {
};

BaseXform.prototype = {
  // ============================================================
  // Virtual Interface
  prepare: function(/* model, options */) {
    // optional preparation (mutation) of model so it is ready for write
  },
  render: function(/* xmlStream, model */) {
    // convert model to xml
  },
  parseOpen: function(/* node */) {
    // Sax Open Node event
  },
  parseText: function(/* node */) {
    // Sax Text event
  },
  parseClose: function(/* name */) {
    // Sax Close Node event
  },
  reconcile: function(/* model, options */) {
    // optional post-parse step (opposite to prepare)
  },
  
  // ============================================================
  reset: function() {
    // to make sure parses don't bleed to next iteration
    this.model = null;

    // if we have a map - reset them too
    if (this.map) {
      var keys = Object.keys(this.map);
      for (var i = 0; i < keys.length; i++) {
        this.map[keys[i]].reset();
      }
    }
  },
  mergeModel: function(obj) {
    // set obj's props to this.model
    this.model = Object.assign(
      this.model || {},
      obj
    );
  },

  parse: function(parser, stream) {
    var self = this;
    return new PromishLib.Promish(function(resolve, reject) {
      parser.on('opentag', function(node) {
        // Watch out for "The number of rows exceeds the limit" errors:
        try {
          self.parseOpen(node);
        } catch (error) {
          // Abandon ship! Prevent the parser from consuming any more resources
          // and reject the rest of the stream.
          // Unfortunately this cleanup code cannot be moved to the parseStream
          // function -- since that doesn't kick in until the next tick,
          // it's 10 seconds later in the example I'm looking at right now
          parser.removeAllListeners();
          stream.unpipe(parser);
          stream.resume(); // Discard the rest of the XML stream
          reject(error);
        }
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
    var promise = this.parse(parser, stream);
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
