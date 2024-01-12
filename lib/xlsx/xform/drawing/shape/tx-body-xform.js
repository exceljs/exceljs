const BaseXform = require('../../base-xform');
const ParagraphXform = require('./p-xform');

// DocumentFormat.OpenXml.Drawing.Spreadsheet.TextBody
class TxBodyXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      'a:p': new ParagraphXform(),
    };
  }

  get tag() {
    return 'xdr:txBody';
  }

  render(xmlStream, textBody) {
    xmlStream.openNode(this.tag);
    xmlStream.openNode('a:bodyPr', {
      vertOverflow: 'clip',
      horzOverflow: 'clip',
      rtlCol: 0,
    });
    if (textBody.vertAlign) {
      xmlStream.addAttribute('anchor', textBody.vertAlign);
    }
    xmlStream.closeNode();
    xmlStream.leafNode('a:lstStyle');
    textBody.paragraphs.forEach(p => {
      this.map['a:p'].render(xmlStream, p);
    });
    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    switch (node.name) {
      case this.tag:
        this.model = {paragraphs: []};
        break;
      case 'a:bodyPr':
        if (node.attributes.anchor) {
          this.model.vertAlign = node.attributes.anchor;
        }
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
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        if (name === 'a:p') {
          this.model.paragraphs.push(this.parser.model);
        }
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        return false;
      default:
        return true;
    }
  }
}

module.exports = TxBodyXform;
