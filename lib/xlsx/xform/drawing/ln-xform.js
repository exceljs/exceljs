const BaseXform = require('../base-xform');
const SolidFillXform = require('./solid-fill-xform');
const NoFillXform = require('./no-fill-xform');

class LnXform extends BaseXform {

  constructor() {
    super();
    this.model = null;
    this.map = {
      'a:solidFill': new SolidFillXform(),
      'a:noFill': new NoFillXform(),
    };
  }

  get tag() {
    return 'a:ln';
  }

  render(xmlStream, model) {
    const m = (typeof model !== 'object' || model === null) ? {weight: 1, opacity: 0} : model;

    xmlStream.openNode(this.tag, {w: (m.weight || 1) * 12700});
    if(m.opacity === 0) {
      this.map['a:noFill'].render(xmlStream, model);
    } else {
      this.map['a:solidFill'].render(xmlStream, model);
    }
    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.model = {
          weight: parseInt(node.attributes.w || '12700', 10) / 12700
        };
        break;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  }

  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        const model = Object.assign({}, this.map['a:solidFill'].model, this.map['a:noFill'].model)
        this.model = Object.keys(model).length ? Object.assign(this.model, model) : null;
        return false;
      default:
        return true;
    }
  }
}

module.exports = LnXform;
