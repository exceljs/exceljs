const BaseXform = require('../base-xform')
const XfrmXform = require('./xfrm-xform')
const PrstGeomXform = require('./prst-geom-xform')
const SolidFillXform = require('./solid-fill-xform')
const NoFillXform = require('./no-fill-xform')
const LnXform = require('./ln-xform')

class SpPrXform extends BaseXform {
  constructor () {
    super();

    this.map = {
      'a:xfrm': new XfrmXform(),
      'a:prstGeom': new PrstGeomXform(),
      'a:solidFill': new SolidFillXform(),
      'a:noFill': new NoFillXform(),
      'a:ln': new LnXform(),
    };
  }

  get tag () {
    return 'xdr:spPr';
  }

  render (xmlStream, model) {
    xmlStream.openNode(this.tag);

    this.map['a:xfrm'].render(xmlStream, model);
    this.map['a:prstGeom'].render(xmlStream, model);
    if(typeof model.fill !== 'object' || model.fill === null || model.fill.opacity === 0) {
      this.map['a:noFill'].render(xmlStream, model.fill);
    } else {
      this.map['a:solidFill'].render(xmlStream, model.fill);
    }
    this.map['a:ln'].render(xmlStream, model.stroke);

    xmlStream.closeNode();
  }

  parseOpen (node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.model = {
          shape: 'rect',
          rotation: 0,
          fill: null,
          stroke: null
        };
        break
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break
    }
    return true
  }

  parseClose (name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true
    }
    switch (name) {
      case this.tag:
        Object.assign(
          this.model,
          this.map['a:xfrm'].model ? this.map['a:xfrm'].model : null,
          this.map['a:prstGeom'].model ? this.map['a:prstGeom'].model : null,
          this.map['a:solidFill'].model ? {fill: this.map['a:solidFill'].model} : null,
          this.map['a:noFill'].model ? {fill: this.map['a:noFill'].model} : null,
          this.map['a:ln'].model ? {stroke: this.map['a:ln'].model} : null,
        );
        return false
      default:
        return true
    }
  }
}

module.exports = SpPrXform
