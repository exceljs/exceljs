const BaseXform = require('../../base-xform');
const SolidFillXform = require('./solid-fill-xform');

// DocumentFormat.OpenXml.Drawing.Run
class RunXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      'a:solidFill': new SolidFillXform(),
    };
  }

  get tag() {
    return 'a:r';
  }

  render(xmlStream, run) {
    xmlStream.openNode(this.tag);
    xmlStream.openNode('a:rPr');
    if (run.font) {
      xmlStream.addAttributes({
        sz: run.font.size ? run.font.size * 100 : undefined,
        b: run.font.bold ? 1 : undefined,
        i: run.font.italic ? 1 : undefined,
        u: run.font.underline || undefined,
      });
    }
    if (run.font && run.font.color) {
      this.map['a:solidFill'].render(xmlStream, run.font.color);
    }
    xmlStream.closeNode();
    xmlStream.leafNode('a:t', undefined, run.text);
    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    switch (node.name) {
      case this.tag:
        this.model = {text: '', font: {}};
        this.parsingText = false;
        break;
      case 'a:rPr':
        if (node.attributes.sz) {
          this.model.font.size = parseInt(node.attributes.sz, 10) / 100;
        }
        if (node.attributes.b) {
          this.model.font.bold = node.attributes.b === '1';
        }
        if (node.attributes.i) {
          this.model.font.italic = node.attributes.i === '1';
        }
        if (node.attributes.u) {
          this.model.font.underline = node.attributes.u;
        }
        break;
      case 'a:t':
        this.parsingText = true;
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

  parseText(text) {
    if (this.parsingText) {
      this.model.text = text.replace(/_x([0-9A-F]{4})_/g, ($0, $1) => String.fromCharCode(parseInt($1, 16)));
    }
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
        if (this.map['a:solidFill'].model) {
          this.model.font.color = this.map['a:solidFill'].model;
        }
        return false;
      case 'a:t':
        this.parsingText = false;
        return true;
      default:
        return true;
    }
  }
}

module.exports = RunXform;
