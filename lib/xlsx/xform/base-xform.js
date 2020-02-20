const SAXStream = require('../../utils/sax-stream');
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

  parseOpen(node) {
    // XML node opened
  }

  parseText(text) {
    // chunk of text encountered for current node
  }

  parseClose(name) {
    // XML node closed
  }

  reconcile(model, options) {
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

  parse(saxStream, stream) {
    return new Promise((resolve, reject) => {
      const abort = error => {
        // Abandon ship! Prevent the parser from consuming any more resources
        saxStream.sax.off('opentag');
        saxStream.sax.off('text');
        saxStream.sax.off('closetag');
        saxStream.sax.off('error');
        saxStream.sax.off('end');
        saxStream.sax.on('error', () => {}); // Ignore any parse errors from the chunk being processed
        if (stream) {
          stream.unpipe(saxStream);
        }
        reject(error);
      };

      saxStream.sax.on('opentag', node => {
        try {
          this.parseOpen(node);
        } catch (error) {
          abort(error);
        }
      });
      saxStream.sax.on('text', text => {
        try {
          this.parseText(text);
        } catch (error) {
          abort(error);
        }
      });
      saxStream.sax.on('closetag', node => {
        try {
          if (!this.parseClose(node.name)) {
            resolve(this.model);
          }
        } catch (error) {
          abort(error);
        }
      });
      saxStream.sax.on('end', () => {
        resolve(this.model);
      });
      saxStream.sax.on('error', error => {
        abort(error);
      });
    });
  }

  parseStream(stream) {
    const saxStream = new SAXStream();
    const promise = this.parse(saxStream, stream);
    stream.pipe(saxStream);

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
  static toAttribute(value, dflt, always = false) {
    if (value === undefined) {
      if (always) {
        return dflt;
      }
    } else if (always || (value !== dflt)) {
      return value.toString();
    }
    return undefined;
  }

  static toStringAttribute(value, dflt, always = false) {
    return BaseXform.toAttribute(value, dflt, always);
  }

  static toStringValue(attr, dflt) {
    return attr === undefined ? dflt : attr;
  }

  static toBoolAttribute(value, dflt, always = false) {
    if (value === undefined) {
      if (always) {
        return dflt;
      }
    } else if (always || (value !== dflt)) {
      return value ? '1' : '0';
    }
    return undefined;
  }

  static toBoolValue(attr, dflt) {
    return attr === undefined ? dflt : attr === '1';
  }

  static toIntAttribute(value, dflt, always = false) {
    return BaseXform.toAttribute(value, dflt, always);
  }

  static toIntValue(attr, dflt) {
    return attr === undefined ? dflt : parseInt(attr, 10);
  }

  static toFloatAttribute(value, dflt, always = false) {
    return BaseXform.toAttribute(value, dflt, always);
  }

  static toFloatValue(attr, dflt) {
    return attr === undefined ? dflt : parseFloat(attr);
  }
}

module.exports = BaseXform;
