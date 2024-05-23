const BaseXform = require('../base-xform');
const ListXform = require('../list-xform');

const CustomFilterXform = require('./custom-filter-xform');
const FilterXform = require('./filter-xform');

class FilterColumnXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      customFilters: new ListXform({
        tag: 'customFilters',
        count: false,
        empty: true,
        childXform: new CustomFilterXform(),
      }),
      filters: new ListXform({
        tag: 'filters',
        count: false,
        empty: true,
        childXform: new FilterXform(),
      }),
    };
  }

  get tag() {
    return 'filterColumn';
  }

  prepare(model, options) {
    model.colId = options.index.toString();
  }

  render(xmlStream, model) {
    if (model.customFilters) {
      xmlStream.openNode(this.tag, {
        colId: model.colId,
        hiddenButton: model.filterButton ? '0' : '1',
      });

      this.map.customFilters.render(xmlStream, model.customFilters);

      xmlStream.closeNode();
      return true;
    }
    if (model.filters) {
      xmlStream.leafNode(this.tag, {
        colId: model.colId,
        hiddenButton: model.filterButton ? '0' : '1',
      });

      this.map.filters.render(xmlStream, model.filters);
    }
    
    return true;
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    const {attributes} = node;
    switch (node.name) { // ensure we are looking at the filterColumn name
      case this.tag: // if the node name matches this classes tag, set the hiddenButton attribute.
        this.model = {
          filterButton: attributes.hiddenButton === '0',
        };
        return true; // finish
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parseOpen(node);
          return true;
        }
        throw new Error(`Unexpected xml node in parseOpen: ${JSON.stringify(node)}`);
    }
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
        this.model.customFilters = this.map.customFilters.model;
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  }
}

module.exports = FilterColumnXform;
