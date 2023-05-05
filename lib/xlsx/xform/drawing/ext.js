const BaseXform = require('exceljs/lib/xlsx/xform/base-xform');

class ExtXform extends BaseXform {
  get tag() {
    return 'a:ext';
  }

  render(xmlStream, model) {

    model.ext&&xmlStream.leafNode(this.tag, {
      'cx': model.ext.extentX,
      'cy': model.ext.extentY
    });
  }

  parseOpen(node) {
    switch (node.name) {
      case this.tag:
        this.model = {
          extentX: node.attributes['cx'],
          extentY: node.attributes['cy'],

        };
        return true;
      default:
        return true;
    }
  }

  parseText() {}

  parseClose(name) {
    switch (name) {
      case this.tag:
        return false;
      default:
        // unprocessed internal nodes
        return true;
    }
  }
}

module.exports = ExtXform;