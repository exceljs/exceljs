const BaseXform = require('../base-xform');

const LnRefXform = require('./ln-ref-xform');
const FillRefXform = require('./fill-ref-xform');

class StyleXform extends BaseXform {
  constructor() {
    super();
    this.model = null;
    this.map = {
      'a:lnRef': new LnRefXform(),
      'a:fillRef': new FillRefXform(),
    };
  }

  get tag() {
    return 'xdr:style';
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag);
    this.map['a:lnRef'].render(xmlStream, model);
    this.map['a:fillRef'].render(xmlStream, model);
    xmlStream.openNode('a:effectRef', {idx: 0});
    xmlStream.leafNode('a:schemeClr', {val: 'accent1'});
    xmlStream.closeNode();
    xmlStream.openNode('a:fontRef', {idx: 'minor'});
    xmlStream.leafNode('a:schemeClr', {val: 'lt1'});
    xmlStream.closeNode();
    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.reset();
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
        this.model = {
          stroke: this.map['a:lnRef'].model,
          fill: this.map['a:fillRef'].model,
        };
        return false;
      default:
        return true;
    }
  }
}

module.exports = StyleXform;
