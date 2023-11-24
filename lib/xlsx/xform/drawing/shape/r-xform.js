const BaseXform = require('../../base-xform');

// DocumentFormat.OpenXml.Drawing.Run
class RunXform extends BaseXform {
  get tag() {
    return 'a:r';
  }

  render(xmlStream, run) {
    xmlStream.openNode(this.tag);
    xmlStream.leafNode('a:rPr');
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
        this.model = {text: ''};
        break;
      default:
        break;
    }
    return true;
  }

  parseText(text) {
    this.model.text = text.replace(/_x([0-9A-F]{4})_/g, ($0, $1) => String.fromCharCode(parseInt($1, 16)));
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
        return false;
      default:
        return true;
    }
  }
}

module.exports = RunXform;
