
const colCache = require('./col-cache');

/**
 * ColumnFlatter is a helper class to create sheets with nested columns.
 * 
 * Based on following concepts
 *  - Walk throught nested input structure to build flat list and tree meta information
 *  - Use "leaf" columns as phisical cols and "branch" as merge-slots
 *  - Generate cell matrix and merge rules
 */
class ColumnFlatter {
  constructor(input, params) {
    this._params = params;
    // key-value storage for item aggregate sizes
    this._sizes = {};
  
    // flat columns list
    this._list = [];

    // cells matrix storage
    this._rows = [];

    this._getFlatList(input);
    this._alignRows(this._alignCells());

    // merge rules storage
    this._merges = [
      ...this._calcVerticalMerges(),
      ...this._calcHorizontalMerges(),
    ];
  }

  /**
   * Append null placeholders for entity list alignement
   */
  _pad(arr, num) {
    if (num > 0) {
      for(let i=0; i<num; i++) {
        arr.push(null);
      }
    }
  }

  /**
   * Filters off invalid columns entries
   */
  _check(item) {
    return item && item.id;
  }

  /**
   * Walk throught tree input.
   * Build flat columns list and aggregate column size (recursive child length sum)
   */
  _getFlatList(input) {
    const trace = (item, meta) => {
      if (!this._check(item)) return;

      const path = [...meta.path, item.id];
      const children = (item.child || []).filter(this._check);
      
      if (children.length && children.length > 1) {
        for (const id of path) {

          if (!this._sizes[id]) {
            this._sizes[id] = 0;
          }
      
          this._sizes[id] += children.length - 1;
        } 

        for(const child of children) {
          trace(child, { path });
        }
      }

      this._list.push({
        meta, 
        ...(children.length === 1 ? children[0] : item)
      });
    }
    
    for (const item of input) {
      trace(item, { path: [] })
    }
  }

  /**
   * Align with cells with null-ish appending
   * by aggregated size num
   */
  _alignCells() {
    const res = [];

    for (const item of this._list) {
      const index = item.meta.path.length;
  
      if (!res[index]) {
        res[index] = [];
      }
  
      res[index].push(item);
  
      if (item.child) {
        this._pad(res[index], this._sizes[item.id]);
      }
    }

    return res;
  }

  /**
   * Align cell groups in rows according 
   * parent cell position
   */
  _alignRows(cells) {
    const width = cells.reduce((acc, row) => Math.max(acc, row.length), 0);

    for(let i = 0; i < cells.length; i++) {
      const row = cells[i];
  
      if (!i) {
        this._rows.push(row);
      } else {
        const items = [];
        const handled = {};
        let added = 0;
  
        for(let k = 0; k<width; k++) {
          const item = row[k];

          if (k + added >= width) {
            break;
          }
  
          if (item) {
            const path = item.meta.path;
            const parent = path[path.length - 1];
  
            if (parent) {
              const parentPos = this._rows[i-1].findIndex(el => el?.id === parent);
              const offset = parentPos - (k + added);
              
              if (offset > 0 && !handled[parent]) {
                added += offset;
  
                this._pad(items, offset);

                handled[parent] = true;
              }
            } 
          }
  
          items.push(item || null);
        }
  
        this._rows.push(items)
      }
    }
  }

  /**
   * Calculates horizontal merge rules
   * 
   * Walks width-throught rows collecting ranges with cell index and its recursive size
   */
  _calcHorizontalMerges() {
    const res = [];

    for (let i=0; i< this._rows.length; i++) {
      const cells = this._rows[i];

      for(let k=0; k < cells.length; k++) {
        const cell = cells[k];
        const span = cell && this._sizes[cell.id];
        const row = i + 1;
  
        span && res.push(colCache.encode(row, k + 1, row, k + span + 1));
      }
    }
  
    return res;
  }

  /**
   * Calculates vertical merge rules
   * 
   * Walks deep-throught rows looking for non-empty cell in row
   */
  _calcVerticalMerges() {
    const depth = this._rows.length - 1;
    const width = this._rows[0].length;
    const res = [];
  
    for(let i = 0; i < width; i++) {
      for(let k = depth; k >= 0; k--) {
        if (this._rows[k][i]) {
          const col = i + 1;

          if (k !== depth) { 
            res.push(colCache.encode(k + 1, col, depth + 1, col));
          }
          
          break;
        }  
      }
    }
  
    return res;
  } 

  /**
   * Collect "leaf" columns
   * 
   * Filter off all cells with "child" property
   */
  getCoumns() {
    const res = [];
  
    for (const item of this._list) {
      if (!item.child) {
        res.push({
          key: item.id,
          ...item
        })
      }
    }
  
    return res;
  }

  /**
   * Cells matrix getter
   */
   getRows() {
    return this._rows;
  }

  /**
   * Merge rules getter
   */
  getMerges() {
    return this._merges;
  }
}

module.exports = ColumnFlatter;