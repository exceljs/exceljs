var _ = require('underscore');

var ColumnSum = module.exports = function(columns) {
    var self = this;
    this.columns = columns;
    this.sums = [];
    this.count = 0;
    _.each(this.columns, function(column) {
        self.sums[column] = 0;
    });
};

ColumnSum.prototype = {
    add: function(row) {
        var self = this;
        _.each(this.columns, function(column) {
            self.sums[column] += row.getCell(column).value;
        });
        this.count++;
    },
    
    toString: function() {
        return this.sums.join(', ');
    },
    toAverages: function() {
        return this.sum.map(function(value) {
            return value ? value / self.count : value;
        }).join(', ');
    }
}