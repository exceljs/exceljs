var _ = require('underscore');

var ColumnSum = module.exports = function(columns) {
    this.columns = columns;
    this.sums = [];
    this.count = 0;
    _.each(this.columns, column => {
        this.sums[column] = 0;
    });
};

ColumnSum.prototype = {
    add: function(row) {
        _.each(this.columns, column => {
            this.sums[column] += row.getCell(column).value;
        });
        this.count++;
    },
    
    toString: function() {
        return this.sums.join(', ');
    },
    toAverages: function() {
        return this.sum.map(value => {
            return value ? value / this.count : value;
        }).join(', ');
    }
}