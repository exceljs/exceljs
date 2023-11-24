const BaseXform = require('../../base-xform');
const StaticXform = require('../../static-xform');
const StyleMatrixReferenceTypeXform = require('./style-matrix-reference-type-xform');

// DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeStyle
class StyleXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      'a:lnRef': new StyleMatrixReferenceTypeXform('a:lnRef'),
      'a:fillRef': new StyleMatrixReferenceTypeXform('a:fillRef'),
      'a:effectRef': new StaticXform(effectRefJSON),
      'a:fontRef': new StaticXform(fontRefJSON),
    };
  }

  get tag() {
    return 'xdr:style';
  }

  render(xmlStream, shape) {
    xmlStream.openNode(this.tag);
    // Must care about the order
    this.map['a:lnRef'].render(xmlStream);
    this.map['a:fillRef'].render(xmlStream);
    this.map['a:effectRef'].render(xmlStream, shape.effectRef);
    this.map['a:fontRef'].render(xmlStream, shape.fontRef);
    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    switch (node.name) {
      case this.tag:
        this.model = {};
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

  parseText() {}

  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        if (this.map['a:lnRef'].model) {
          this.model.outline = this.map['a:lnRef'].model;
        }
        if (this.map['a:fillRef'].model) {
          this.model.fill = {
            type: 'solid',
            color: this.map['a:fillRef'].model,
          };
        }
        return false;
      default:
        return true;
    }
  }
}

const effectRefJSON = {
  tag: 'a:effectRef',
  $: {idx: '0'},
  c: [
    {
      tag: 'a:schemeClr',
      $: {val: 'accent1'},
    },
  ],
};

const fontRefJSON = {
  tag: 'a:fontRef',
  $: {idx: 'minor'},
  c: [
    {
      tag: 'a:schemeClr',
      $: {val: 'lt1'},
    },
  ],
};

module.exports = StyleXform;
