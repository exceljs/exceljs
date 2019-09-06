const BaseXform = require('../base-xform');

class FilterColumnXform extends BaseXform {
  get tag() {
    return 'filterColumn';
  }

  prepare(model, options) {
    model.colId = options.index.toString();
  }

  render(xmlStream, model) {
    xmlStream.leafNode(
      this.tag,
      {
        colId: model.colId,
        hiddenButton: model.filterButton ? '0' : '1',
      }
    );
    return true;
  }

  parseOpen(node) {
    if (node.name === this.tag) {
      const {attributes} = node;
      this.model = {
        filterButton: attributes.hiddenButton === '0',
      };
      return true;
    }
    return false;
  }

  parseText() {}

  parseClose() {
    return false;
  }
}

module.exports = FilterColumnXform;
