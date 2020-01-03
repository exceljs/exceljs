var HrStopwatch = module.exports = function() {
    this.total = 0;
};

HrStopwatch.prototype = {
    start: function() {
        this.hrStart = process.hrtime();
    },

    stop: function() {
        if (this.hrStart) {
            this.total += this._span;
            delete this.hrStart;
        }
    },

    reset: function() {
        this.total = 0;
        delete this.hrStart;
    },

    get span() {
        if (this.hrStart) {
            return this.total + this._span;
        } else {
            return this.total;
        }
    },

    get _span() {
        var hrNow = process.hrtime();
        var start = this.hrStart[0] + this.hrStart[1] / 1e9;
        var finish = hrNow[0] + hrNow[1] / 1e9;
        return finish - start;
    },
    
    toString: function(format) {
        switch (format) {
            case "ms":
                return "" + this.ms + "ms";
            case "microseconds":
                return "" + this.microseconds + "\xB5s";
            default:
                return "" + this.span + " seconds";
        }
    },
    get ms() {
        return Math.round(this.span * 1000);
    },
    get microseconds() {
        return Math.round(this.span * 1000000);
    }
    
};
