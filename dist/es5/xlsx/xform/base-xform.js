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
var BaseXform = module.exports = function () /* model, name */{};

BaseXform.prototype = {
  // ============================================================
  // Virtual Interface
  prepare: function prepare() /* model, options */{
    // optional preparation (mutation) of model so it is ready for write
  },
  render: function render() /* xmlStream, model */{
    // convert model to xml
  },
  parseOpen: function parseOpen() /* node */{
    // Sax Open Node event
  },
  parseText: function parseText() /* node */{
    // Sax Text event
  },
  parseClose: function parseClose() /* name */{
    // Sax Close Node event
  },
  reconcile: function reconcile() /* model, options */{
    // optional post-parse step (opposite to prepare)
  },

  // ============================================================
  reset: function reset() {
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
  mergeModel: function mergeModel(obj) {
    // set obj's props to this.model
    this.model = Object.assign(this.model || {}, obj);
  },

  parse: function parse(parser, stream) {
    var self = this;
    return new PromishLib.Promish(function (resolve, reject) {
      function abort(error) {
        // Abandon ship! Prevent the parser from consuming any more resources
        parser.removeAllListeners();
        stream.unpipe(parser);
        reject(error);
      }

      parser.on('opentag', function (node) {
        try {
          self.parseOpen(node);
        } catch (error) {
          abort(error);
        }
      });
      parser.on('text', function (text) {
        try {
          self.parseText(text);
        } catch (error) {
          abort(error);
        }
      });
      parser.on('closetag', function (name) {
        try {
          if (!self.parseClose(name)) {
            resolve(self.model);
          }
        } catch (error) {
          // console.log('BaseXform closetag', error.stack)
          abort(error);
        }
      });
      parser.on('end', function () {
        resolve(self.model);
      });
      parser.on('error', function (error) {
        abort(error);
      });
    });
  },
  parseStream: function parseStream(stream) {
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

  toXml: function toXml(model) {
    var xmlStream = new XmlStream();
    this.render(xmlStream, model);
    return xmlStream.xml;
  }
};
//# sourceMappingURL=base-xform.js.map
