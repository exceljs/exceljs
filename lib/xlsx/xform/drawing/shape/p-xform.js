const BaseXform = require('../../base-xform');
const RunXform = require('./r-xform');

// DocumentFormat.OpenXml.Drawing.Paragraph
class ParagraphXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      'a:r': new RunXform(),
    };
  }

  get tag() {
    return 'a:p';
  }

  render(xmlStream, paragraph) {
    xmlStream.openNode('a:p');
    xmlStream.openNode('a:pPr');
    if (paragraph.alignment) {
      xmlStream.addAttribute('algn', paragraph.alignment);
    }
    xmlStream.closeNode();
    paragraph.runs.forEach(r => {
      this.map['a:r'].render(xmlStream, r);
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
        this.model = {runs: []};
        break;
      case 'a:pPr':
        if (node.attributes.algn) {
          this.model.alignment = node.attributes.algn;
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
        if (name === 'a:r') {
          this.model.runs.push(this.parser.model);
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

module.exports = ParagraphXform;
