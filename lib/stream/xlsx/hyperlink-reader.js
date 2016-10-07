/**
 * Copyright (c) 2015 Guyon Roche
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 * 
 */
'use strict';

var events = require('events');
var Sax = require('sax');

var utils = require('../../utils/utils');
var Enums = require('../../doc/enums');

var HyperlinkReader = module.exports = function(workbook, id) {
    // in a workbook, each sheet will have a number
    this.id = id;
    
    this._workbook = workbook;    
};

utils.inherits(HyperlinkReader, events.EventEmitter, {
    
    get count() {
        return this.hyperlinks && this.hyperlinks.length || 0;
    },
    
    each: function(fn) {
        return this.hyperlinks.forEach(fn);
    },
    
    read: function(entry, options) {
        var self = this;
        var emitHyperlinks = false;
        var hyperlinks = null;
        switch(options.hyperlinks) {
            case 'emit':
                emitHyperlinks = true;
                break;
            case 'cache':
                this.hyperlinks = hyperlinks = {};
                break;
        }
        if (!emitHyperlinks && !hyperlinks) {
            entry.autodrain();
            return;
        }
        
        var parser = Sax.createStream(true, {});
        parser.on('opentag', function(node) {
            if (node.name === 'Relationship') {
                var rId = node.attributes.Id;
                switch (node.attributes.Type) {
                    case XLSX.RelType.Hyperlink:
                        var relationship = {
                            type: Enums.RelationshipType.Styles,
                            rId: rId,
                            target: node.attributes.Target,
                            targetMode: node.attributes.TargetMode
                        };
                        if (emitHyperlinks) {
                            self.emit('hyperlink', relationship);
                        } else {
                            hyperlinks[relationship.rId] = relationship;
                        }
                        break;
                }
            }
        });
        parser.on('end', function() {
            self.emit('finished');
        });
        
        // create a down-stream flow-control to regulate the stream
        var flowControl = this.workbook.flowControl.createChild();
        flowControl.pipe(parser, {sync: true});
        entry.pipe(flowControl);
    }
});