const _ = require('../../../utils/under-dash');
const BaseXform = require('../base-xform');

class SheetProtectionXform extends BaseXform {
  get tag() {
    return 'sheetProtection';
  }

  render(xmlStream, model) {
    if (model) {
      const attributes = {
        algorithmName: model.algorithmName,
        hashValue: model.hashValue,
        saltValue: model.saltValue,
        spinCount: model.spinCount,
        sheet: model.sheet,
        objects: model.objects,
        scenarios: model.scenarios,
        selectLockedCells: model.selectLockedCells,
        selectUnlockedCells: model.selectUnlockedCells,
        formatCells: model.formatCells,
        formatColumns: model.formatColumns,
        formatRows: model.formatRows,
        insertColumns: model.insertColumns,
        insertRows: model.insertRows,
        insertHyperlinks: model.insertHyperlinks,
        deleteColumns: model.deleteColumns,
        deleteRows: model.deleteRows,
        sort: model.sort,
        autoFilter: model.autoFilter,
        pivotTables: model.pivotTables,
      };
      if (_.some(attributes, value => value !== undefined)) {
        xmlStream.leafNode(this.tag, attributes);
      }
    }
  }

  parseOpen(node) {
    switch (node.name) {
      case this.tag:
        this.model = {
          algorithmName: node.attributes.algorithmName,
          hashValue: node.attributes.hashValue,
          saltValue: node.attributes.saltValue,
          spinCount: node.attributes.spinCount,
          sheet: parseInt(node.attributes.sheet || '1', 10),
          objects: parseInt(node.attributes.objects || '1', 10),
          scenarios: parseInt(node.attributes.scenarios || '1', 10),
          selectLockedCells: parseInt(node.attributes.selectLockedCells || '0', 10),
          selectUnlockedCells: parseInt(node.attributes.selectUnlockedCells || '0', 10),
          formatCells: parseInt(node.attributes.formatCells || '1', 10),
          formatColumns: parseInt(node.attributes.formatColumns || '1', 10),
          formatRows: parseInt(node.attributes.formatRows || '1', 10),
          insertColumns: parseInt(node.attributes.insertColumns || '1', 10),
          insertRows: parseInt(node.attributes.insertRows || '1', 10),
          insertHyperlinks: parseInt(node.attributes.insertHyperlinks || '1', 10),
          deleteColumns: parseInt(node.attributes.deleteColumns || '1', 10),
          deleteRows: parseInt(node.attributes.deleteRows || '1', 10),
          sort: parseInt(node.attributes.sort || '1', 10),
          autoFilter: parseInt(node.attributes.autoFilter || '1', 10),
          pivotTables: parseInt(node.attributes.pivotTables || '1', 10),
        };
        return true;
      default:
        return false;
    }
  }

  parseText() {}

  parseClose() {
    return false;
  }
}

module.exports = SheetProtectionXform;
