'use strict';

const Sax = require('sax');
const PromishLib = require('../../utils/promish');

const XmlStream = require('../../utils/xml-stream');

// Base class for Xforms
const BaseXform = (module.exports = function(/* model, name */) {});

BaseXform.prototype = {
  // ============================================================
  // Virtual Interface
  prepare(/* model, options */) {
    // optional preparation (mutation) of model so it is ready for write
  },
  render(/* xmlStream, model */) {
    // convert model to xml
  },
  parseOpen(/* node */) {
    // Sax Open Node event
  },
  parseText(/* node */) {
    // Sax Text event
  },
  parseClose(/* name */) {
    // Sax Close Node event
  },
  reconcile(/* model, options */) {
    // optional post-parse step (opposite to prepare)
  },

  // ============================================================
  reset() {
    // to make sure parses don't bleed to next iteration
    this.model = null;

    // if we have a map - reset them too
    if (this.map) {
      const keys = Object.keys(this.map);
      for (let i = 0; i < keys.length; i++) {
        this.map[keys[i]].reset();
      }
    }
  },
  mergeModel(obj) {
    // set obj's props to this.model
    this.model = Object.assign(this.model || {}, obj);
  },

  parse(parser, stream) {
    const self = this;
    return new PromishLib.Promish((resolve, reject) => {
      function abort(error) {
        // Abandon ship! Prevent the parser from consuming any more resources
        parser.removeAllListeners();
        stream.unpipe(parser);
        reject(error);
      }

      parser.on('opentag', node => {
        try {
          self.parseOpen(node);
        } catch (error) {
          abort(error);
        }
      });
      parser.on('text', text => {
        try {
          self.parseText(text);
        } catch (error) {
          abort(error);
        }
      });
      parser.on('closetag', name => {
        try {
          if (!self.parseClose(name)) {
            resolve(self.model);
          }
        } catch (error) {
          // console.log('BaseXform closetag', error.stack)
          abort(error);
        }
      });
      parser.on('end', () => {
        resolve(self.model);
      });
      parser.on('error', error => {
        abort(error);
      });
    });
  },
  parseStream(stream) {
    const parser = Sax.createStream(true, {});
    const promise = this.parse(parser, stream);
    stream.pipe(parser);

    return promise;
  },

  get xml() {
    // convenience function to get the xml of this.model
    // useful for manager types that are built during the prepare phase
    return this.toXml(this.model);
  },

  toXml(model) {
    const xmlStream = new XmlStream();
    this.render(xmlStream, model);
    return xmlStream.xml;
  },
};
