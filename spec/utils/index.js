'use strict';

var expect = require('chai').expect;
var verquire = require('./verquire');

var _ = require('./under-dash');
var tools = require('./tools');

var testWorkbookReader = require('./test-workbook-reader');

var Row = verquire('doc/row');
var Column = verquire('doc/column');

var testSheets = {
  dataValidations: require('./test-data-validation-sheet'),
  values: require('./test-values-sheet'),
  splice: require('./test-spliced-sheet')
};

function getOptions(docType, options) {
  var result;
  switch (docType) {
    case 'xlsx':
      result = {
        sheetName: 'values',
        checkFormulas: true,
        checkMerges: true,
        checkStyles: true,
        checkBadAlignments: true,
        checkSheetProperties: true,
        dateAccuracy: 3,
        checkViews: true
      };
      break;
    case 'csv':
      result = {
        sheetName: 'sheet1',
        checkFormulas: false,
        checkMerges: false,
        checkStyles: false,
        checkBadAlignments: false,
        checkSheetProperties: false,
        dateAccuracy: 1000,
        checkViews: false
      };
      break;
    default:
      throw new Error('Bad doc-type: ' + docType);
  }
  return Object.assign(result, options);
}


module.exports = {
  views: tools.fix(require('./data/views.json')),
  testValues: tools.fix(require('./data/sheet-values.json')),
  styles: tools.fix(require('./data/styles.json')),
  properties: tools.fix(require('./data/sheet-properties.json')),
  pageSetup: tools.fix(require('./data/page-setup.json')),

  createTestBook: function(workbook, docType, sheets) {
    var options = getOptions(docType);
    sheets = sheets || ['values'];
    
    workbook.views = [
      {x: 1, y: 2, width: 10000, height: 20000, firstSheet: 0, activeTab: 0}
    ];
    
    sheets.forEach(function(sheet) {
      var testSheet = _.get(testSheets, sheet);
      testSheet.addSheet(workbook, options);
    });

    return workbook;
  },

  checkTestBook: function(workbook, docType, sheets, options) {
    options = getOptions(docType, options);
    sheets = sheets || ['values'];

    expect(workbook).to.not.be.undefined();

    if (options.checkViews) {
      expect(workbook.views).to.deep.equal([{x: 1, y: 2, width: 10000, height: 20000, firstSheet: 0, activeTab: 0, visibility: 'visible'}]);
    }

    sheets.forEach(function(sheet) {
      var testSheet = _.get(testSheets, sheet);
      testSheet.checkSheet(workbook, options);
    });
  },

  checkTestBookReader: testWorkbookReader.checkBook,

  createSheetMock: function() {
    return {
      _keys: {},
      _cells: {},
      rows: [],
      columns: [],
      properties: {
        outlineLevelCol: 0,
        outlineLevelRow: 0
      },
      
      addColumn: function(colNumber, defn) {
        return (this.columns[colNumber - 1] = new Column(this, colNumber, defn));
      },
      getColumn: function(colNumber) {
        var column = this.columns[colNumber - 1] || this._keys[colNumber];
        if (!column) {
          column = this.columns[colNumber - 1] = new Column(this, colNumber);
        }
        return column;
      },
      getRow: function(rowNumber) {
        var row = this.rows[rowNumber - 1];
        if (!row) {
          row = this.rows[rowNumber - 1] = new Row(this, rowNumber);
        }
        return row;
      },
      getCell: function(rowNumber, colNumber) {
        return this.getRow(rowNumber).getCell(colNumber);
      },
      getColumnKey(key) {
        return this._keys[key];
      },
      setColumnKey(key, value) {
        this._keys[key] = value;
      },
      deleteColumnKey(key) {
        delete this._keys[key];
      },
      eachColumnKey(f) {
        _.each(this._keys, f);
      },
      eachRow(opt, f) {
        if (!f) {
          f = opt;
          opt = {};
        }
        if (opt && opt.includeEmpty) {
          var n = this.rows.length;
          for (var i = 1; i <= n; i++) {
            f(this.getRow(i), i);
          }
        } else {
          this.rows.forEach(function (r, i) {
            if (r || includeEmpty) {
              f(r, i + 1);
            }
          });
        }
      },
    };
  }
};
