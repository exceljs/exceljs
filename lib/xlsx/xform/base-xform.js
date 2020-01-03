const Sax = require('sax');

const XmlStream = require('../../utils/xml-stream');

/* 'virtual' methods used as a form of documentation */
/* eslint-disable class-methods-use-this */

// Base class for Xforms
class BaseXform {
  // constructor(/* model, name */) {}

  // ============================================================
  // Virtual Interface
  prepare(/* model, options */) {
    // optional preparation (mutation) of model so it is ready for write
  }

  render(/* xmlStream, model */) {
    // convert model to xml
  }

  parseOpen(/* node */) {
    // Sax Open Node event
  }

  parseText(/* node */) {
    // Sax Text event
  }

  parseClose(/* name */) {
    // Sax Close Node event
  }

  reconcile(/* model, options */) {
    // optional post-parse step (opposite to prepare)
  }

  // ============================================================
  reset() {
    // to make sure parses don't bleed to next iteration
    this.model = null;

    // if we have a map - reset them too
    if (this.map) {
      Object.values(this.map).forEach(xform => {
        if (xform instanceof BaseXform) {
          xform.reset();
        } else if (xform.xform) {
          xform.xform.reset();
        }
      });
    }
  }

  mergeModel(obj) {
    // set obj's props to this.model
    this.model = Object.assign(this.model || {}, obj);
  }

  parse(parser, stream) {
    return new Promise((resolve, reject) => {
      const abort = error => {
        // Abandon ship! Prevent the parser from consuming any more resources
        parser.removeAllListeners();
        parser.on('error', () => {}); // Ignore any parse errors from the chunk being processed
        stream.unpipe(parser);
        reject(error);
      };

      parser.on('opentag', node => {
        try {
          this.parseOpen(node);
        } catch (error) {
          abort(error);
        }
      });
      parser.on('text', text => {
        try {
          this.parseText(text);
        } catch (error) {
          abort(error);
        }
      });
      parser.on('closetag', name => {
        try {
          if (!this.parseClose(name)) {
            resolve(this.model);
          }
        } catch (error) {
          abort(error);
        }
      });
      parser.on('end', () => {
        resolve(this.model);
      });
      parser.on('error', error => {
        abort(error);
      });
    });
  }

  parseStream(stream) {
    const parser = Sax.createStream(true, {});
    const promise = this.parse(parser, stream);
    stream.pipe(parser);

    return promise;
  }

  get xml() {
    // convenience function to get the xml of this.model
    // useful for manager types that are built during the prepare phase
    return this.toXml(this.model);
  }

  toXml(model) {
    const xmlStream = new XmlStream();
    this.render(xmlStream, model);
    return xmlStream.xml;
  }

  // ============================================================
  // Useful Utilities
  static toAttribute(value, dflt) {
    if ((value !== undefined) && (value !== dflt)) {
      return value.toString();
    }
    return undefined;
  }


  static toStringAttribute(value, dflt) {
    if (value !== dflt) {
      return value;
    }
    return undefined;
  }

  static toStringValue(attr, dflt) {
    return (attr === undefined) ? dflt : attr;
  }

  static toBoolAttribute(value, dflt) {
    if ((value !== undefined) && (value !== dflt)) {
      return value ? '1' : '0';
    }
    return undefined;
  }

  static toBoolValue(attr, dflt) {
    return (attr === undefined) ? dflt : (attr === '1');
  }

  static toIntAttribute(value, dflt) {
    return BaseXform.toAttribute(value, dflt);
  }

  static toIntValue(attr, dflt) {
    return (attr === undefined) ? dflt : parseInt(attr, 10);
  }

  static toFloatAttribute(value, dflt) {
    return BaseXform.toAttribute(value, dflt);
  }

  static toFloatValue(attr, dflt) {
    return (attr === undefined) ? dflt : parseFloat(attr);
  }


}

module.exports = BaseXform;
