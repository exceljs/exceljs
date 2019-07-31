const FilterColumnXform = require('../../../../../lib/xlsx/xform/table/filter-column-xform');
const testXformHelper = require('./../test-xform-helper');

const expectations = [
  {
    title: 'showing filter',
    create() {
      return new FilterColumnXform();
    },
    initialModel: {filterButton: true},
    preparedModel: {colId: '0', filterButton: true},
    xml: '<filterColumn colId="0" hiddenButton="0" />',
    get parsedModel() { return this.initialModel; },
    tests: ['prepare', 'render', 'renderIn', 'parse'],
    options: {index: 0},
  },
  {
    title: 'hidden filter',
    create() {
      return new FilterColumnXform();
    },
    initialModel: {filterButton: false},
    preparedModel: {colId: '1', filterButton: false},
    xml: '<filterColumn colId="1" hiddenButton="1" />',
    get parsedModel() { return this.initialModel; },
    tests: ['prepare', 'render', 'renderIn', 'parse'],
    options: {index: 1},
  },
];

describe('FilterColumnXform', () => {
  testXformHelper(expectations);
});
