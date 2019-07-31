const BaseXform = require('../base-xform');

const AlignmentXform = require('./alignment-xform');
const BorderXform = require('./border-xform');
const FillXform = require('./fill-xform');
const FontXform = require('./font-xform');
const NumFmtXform = require('./numfmt-xform');

// <xf numFmtId="[numFmtId]" fontId="[fontId]" fillId="[fillId]" borderId="[xf.borderId]" xfId="[xfId]">
//   Optional <alignment>
// </xf>

// Style assists translation from style model to/from xlsx
class DxfXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      alignment: new AlignmentXform(),
      border: new BorderXform(),
      fill: new FillXform(),
      font: new FontXform(),
      numFmt: new NumFmtXform(),
    };
  }

  get tag() {
    return 'dxf';
  }

  // how do we generate dxfid?

  render(xmlStream, model) {
    xmlStream.openNode(this.tag);

    if (model.alignment) {
      this.map.alignment.render(xmlStream, model.alignment);
    }
    if (model.border) {
      this.map.border.render(xmlStream, model.border);
    }
    if (model.fill) {
      this.map.fill.render(xmlStream, model.fill);
    }
    if (model.font) {
      this.map.font.render(xmlStream, model.font);
    }
    if (model.numFmt) {
      this.map.numFmt.render(xmlStream, model.numFmt);
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
        return true;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        return true;
    }
  }

  parseText(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    if (name === this.tag) {
      this.model = {
        alignment: this.map.alignment.model,
        border: this.map.border.model,
        fill: this.map.fill.model,
        font: this.map.font.model,
        numFmt: this.map.numFmt.model,
      };
      return false;
    }

    return true;
  }
}

module.exports = DxfXform;
