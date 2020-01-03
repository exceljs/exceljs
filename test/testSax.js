var fs = require('fs');
var events = require('events');
var _ = require('underscore');
var Promise = require('bluebird');
var Sax = require('sax');
var unzip = require('unzip');

var filename = process.argv[2];

var Row = function(r) {
    this.number = r;
    this.cells = {};
}
Row.prototype = {
    add: function(cell) {
        this.cells[cell.address] = cell;
    }
}

var count = 0;
var e = new events.EventEmitter();
e.on('row', function(row) {
    count++;
    
    if (count % 1000 === 0) {
        process.stdout.write('Count:' + count + '\u001b[0G'); // "\033[0G"
    }
});
e.on('finished', function() {
    console.log('Finished worksheet: ' + count);
});

var zip = unzip.Parse();
zip.on('entry',function (entry) {
    if (entry.path.match(/xl\/worksheets\/sheet\d+[.]xml/)) {
        parseSheet(entry,e);
    }
});

function parseSheet(entry, emitter) {
    var parser = Sax.createStream(true, {});
    var row = null;
    var cell = null;
    var current = null;
    parser.on('opentag', function(node) {
        switch(node.name) {
            case 'row':
                var r = parseInt(node.attributes.r);
                row = new Row(r);
                break;
            case 'c':
                cell = {
                    address: node.attributes.r,
                    s: parseInt(node.attributes.s),
                    t: node.attributes.t
                };
                break;
            case 'v':
                current = cell.v = { text: '' };
                break;
            case 'f':
                current = cell.f = { text: '' };
                break;
        }
    });
    parser.on('text', function (text) {
        if (current) {
            current.text += text;
        }
    });
    parser.on('closetag', function(name) {
        switch(name) {
            case 'row':
                emitter.emit('row', row);
                row = null;
                break;
            case 'c':
                row.add(cell);
                break;
        }
    });
    parser.on('end', function() {
        e.emit('finished');
    });
    entry.pipe(parser);
}

var stream = fs.createReadStream(filename);
var eod = false;
stream.on('end', function() {
    eod = true;
});
function schedule() {
    setImmediate(function() {
        if (!eod) {
            var data = stream.read(16384);
            if (data && data.length) {
                zip.write(data);
            }
            schedule();
        }
    });
}

//stream.pipe(zip);
schedule();