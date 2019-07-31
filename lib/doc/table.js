const colCache = require('./../utils/col-cache');

class Table {
  constructor(worksheet, table) {
    this.worksheet = worksheet;
    if (table) {
      this.table = table;
      this.validate();
      this.store();
    }
  }

  getFormula(column) {
    // get the correct formula to apply to the totals row
    switch (column.totalsRowFunction) {
      case 'none': return null;
      case 'average': return `SUBTOTAL(101,${this.table.name}[${column.name}])`;
      case 'countNums': return `SUBTOTAL(102,${this.table.name}[${column.name}])`;
      case 'count': return `SUBTOTAL(103,${this.table.name}[${column.name}])`;
      case 'max': return `SUBTOTAL(104,${this.table.name}[${column.name}])`;
      case 'min': return `SUBTOTAL(105,${this.table.name}[${column.name}])`;
      case 'stdDev': return `SUBTOTAL(106,${this.table.name}[${column.name}])`;
      case 'var': return `SUBTOTAL(107,${this.table.name}[${column.name}])`;
      case 'custom': return column.totalsRowFormula;
      default:
        throw new Error(`Invalid Totals Row Function: ${column.totalsRowFunction}`);
    }
  }

  validate() {
    const {table} = this;
    // set defaults and check is valid
    const assign = (o, name, dflt) => {
      if (o[name] === undefined) {
        o[name] = dflt;
      }
    };
    assign(table, 'headerRow', true);
    assign(table, 'totalsRow', false);

    assign(table, 'style', {});
    assign(table.style, 'theme', 'TableStyleMedium2');
    assign(table.style, 'showFirstColumn', false);
    assign(table.style, 'showLastColumn', false);
    assign(table.style, 'showRowStripes', false);
    assign(table.style, 'showColumnStripes', false);

    const assert = (test, message) => {
      if (!test) {
        throw new Error(message);
      }
    };
    assert(table.ref, 'Table must have ref');
    assert(table.columns, 'Table must have column definitions');
    assert(table.rows, 'Table must have row definitions');

    table.tl = colCache.decodeAddress(table.ref);
    const {row, col} = table.tl;
    assert(row > 0, 'Table must be on valid row');
    assert(col > 0, 'Table must be on valid col');

    const width = table.columns.length;
    const height = table.rows.length;
    const filterHeight = height + (table.headerRow ? 1 : 0);
    const tableHeight = filterHeight + (table.totalsRow ? 1 : 0);

    // autoFilterRef is a range that includes optional headers only
    table.autoFilterRef = colCache.encode(
      row,
      col,
      row + filterHeight - 1,
      col + width - 1
    );

    // tableRef is a range that includes optional headers and totals
    table.tableRef = colCache.encode(
      row,
      col,
      row + tableHeight - 1,
      col + width - 1
    );

    table.columns.forEach((column, i) => {
      assert(column.name, `Column ${i} must have a name`);
      if (i === 0) {
        assign(column, 'totalsRowLabel', 'Total');
      } else {
        assign(column, 'totalsRowFunction', 'none');
        column.totalsRowFormula = this.getFormula(column);
      }
    });
  }

  store() {
    // where the table needs to store table data, headers, footers in
    // the sheet...
    const assignStyle = (cell, style) => {
      if (style) {
        Object.keys(style).forEach(key => {
          cell[key] = style[key];
        });
      }
    };

    const {worksheet, table} = this;
    const {row, col} = table.tl;
    let count = 0;
    if (table.headerRow) {
      const r = worksheet.getRow(row + count++);
      table.columns.forEach((column, j) => {
        const {style, name} = column;
        const cell = r.getCell(col + j);
        cell.value = name;
        assignStyle(cell, style);
      });
    }
    table.rows.forEach(data => {
      const r = worksheet.getRow(row + count++);
      data.forEach((value, j) => {
        const cell = r.getCell(col + j);
        cell.value = value;

        assignStyle(cell, table.columns[j].style);
      });
    });

    if (table.totalsRow) {
      const r = worksheet.getRow(row + count++);
      table.columns.forEach((column, j) => {
        const cell = r.getCell(col + j);
        if (j === 0) {
          cell.value = column.totalsRowLabel;
        } else {
          const formula = this.getFormula(column);
          if (formula) {
            cell.value = {
              formula: column.totalsRowFormula,
              result: column.totalsRowResult,
            };
          }
        }

        assignStyle(cell, column.style);
      });
    }
  }

  load(worksheet) {
    // where the table will read necessary features from a loaded sheet
    const {table} = this;
    const {row, col} = table.tl;
    let count = 0;
    if (table.headerRow) {
      const r = worksheet.getRow(row + count++);
      table.columns.forEach((column, j) => {
        const cell = r.getCell(col + j);
        cell.value = column.name;
      });
    }
    table.rows.forEach(data => {
      const r = worksheet.getRow(row + count++);
      data.forEach((value, j) => {
        const cell = r.getCell(col + j);
        cell.value = value;
      });
    });

    if (table.totalsRow) {
      const r = worksheet.getRow(row + count++);
      table.columns.forEach((column, j) => {
        const cell = r.getCell(col + j);
        if (j === 0) {
          cell.value = column.totalsRowLabel;
        } else {
          const formula = this.getFormula(column);
          if (formula) {
            cell.value = {
              formula: column.totalsRowFormula,
              result: column.totalsRowResult,
            };
          }
        }
      });
    }
  }

  get model() {
    return this.table;
  }

  set model(value) {
    this.table = value;
  }

  // ================================================================
  // TODO: Mutating methods
  cacheState() {
    if (!this._cache) {
      this._cache = {
        ref: this.ref,
      };
    }
  }

  /* eslint-disable no-unused-vars */
  addRow(values, rowNumber) {
    // Add a row of data, either insert at rowNumber or append
    this.cacheState();
  }

  removeRow(rowNumber) {
    // Remove a row of data
    this.cacheState();
  }

  addColumn(column, values, colNumber) {
    // Add a new column, including column defn and values
    // Inserts at colNumber or adds to the right
    this.cacheState();
  }

  removeColumn(colNumber) {
    // Remove a column with data
    this.cacheState();
  }

  assign(target, prop, value) {
    this.cacheState();
    target[prop] = value;
  }

  /* eslint-disable lines-between-class-members */
  get ref() { return this.table.ref; }
  set ref(value) { this.assign(this.table, 'ref', value); }

  get name() { return this.table.name; }
  set name(value) { this.table.name = value; }

  get displayName() { return this.table.displyName || this.table.name; }
  set displayNamename(value) { this.table.displayName = value; }

  get headerRow() { return this.table.headerRow; }
  set headerRow(value) { this.assign(this.table, 'headerRow', value); }

  get totalsRow() { return this.table.totalsRow; }
  set totalsRow(value) { this.assign(this.table, 'totalsRow', value); }

  get theme() { return this.table.style.name; }
  set theme(value) { this.table.style.name = value; }

  get showFirstColumn() { return this.table.style.showFirstColumn; }
  set showFirstColumn(value) { this.table.style.showFirstColumn = value; }

  get showLastColumn() { return this.table.style.showLastColumn; }
  set showLastColumn(value) { this.table.style.showLastColumn = value; }

  get showRowStripes() { return this.table.style.showRowStripes; }
  set showRowStripes(value) { this.table.style.showRowStripes = value; }

  get showColumnStripes() { return this.table.style.showColumnStripes; }
  set showColumnStripes(value) { this.table.style.showColumnStripes = value; }

}

module.exports = Table;
