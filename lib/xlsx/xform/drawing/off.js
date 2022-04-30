const BaseXform = require('exceljs/lib/xlsx/xform/base-xform');

class OffXform extends BaseXform {
  get tag() {
    return 'a:off';
  }

  render(xmlStream, model) {

    model.off&&xmlStream.leafNode(this.tag, {
      'x': model.off.offsetX,
      'y': model.off.offsetY
    });
  }

  parseOpen(node) {
    switch (node.name) {
      case this.tag:
        this.model = {
          offsetX: node.attributes['x'],
          offsetY: node.attributes['y'],
        };
        return true;
      default:
        return true;
    }
  }

  parseText() {}

  parseClose(name) {
    return false;
  }
}

module.exports = OffXform;