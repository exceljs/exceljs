var fs = require('fs');
var events = require('events');
var _ = require('underscore');
var Promise = require('bluebird');
var Sax = require('sax');
var unzip = require('unzip');
var utils = require('./utils/utils');

var Row = function(r) {
    this.number = r;
    this.cells = {};
}
Row.prototype = {
    add: function(cell) {
        this.cells[cell.address] = cell;
    }
}

var parser = Sax.createStream(true, {});

var target = 0;
var count = 0;
var e = new events.EventEmitter();
e.on('drain', function() {
    setImmediate(function() {
        var a = [];
        for (var i = 0; i < 1000; i++) {
            a.push('<row r="' + (target++) + '">');
            for (var j = 0; j < 10; j++) {
                a.push('<cell c="' + j + '">' + utils.randomNum(10000) + '</cell>');
            }
            a.push('</row>');
        }
        var buf = new Buffer(a.join(''));
        parser.write(buf);
    });
});
e.on('row', function(row) {
    if (++count % 1000 === 0) {
        process.stdout.write('Count:' + count + ', ' + target + '\u001b[0G'); // "\033[0G"
    }
    if (target - count < 100) {
        e.emit('drain');
    }
});
e.on('finished', function() {
    console.log('Finished worksheet: ' + count);
});

var row = null;
var cell = null;
parser.on('opentag', function(node) {
    //console.log('opentag ' + node.name);
    switch(node.name) {
        case 'row':
            var r = parseInt(node.attributes.r);
            row = new Row(r);
            break;
        case 'cell':
            var c = parseInt(node.attributes.c);
            cell = {
                c: c,
                value: ''
            };
            break;
    }
});
parser.on('text', function (text) {
    //console.log('text ' + text);
    if (cell) {
        cell.value += text;
    }
});
parser.on('closetag', function(name) {
    //console.log('closetag ' + name);
    switch(name) {
        case 'row':
            e.emit('row', row);
            row = null;
            break;
        case 'cell':
            row.add(cell);
            break;
    }
});
parser.on('end', function() {
    e.emit('finished');
});

parser.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><root>');
e.emit('drain');
