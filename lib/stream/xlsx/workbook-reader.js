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
 "use strict";

 var fs = require("fs");
 var async = require("async");
 var events = require("events");
 var Stream = require("stream");
 var Promise = require("bluebird");
 var _ = require("underscore");
 var StreamZip = require("node-stream-zip");
 var Sax = require("sax");

 var utils = require("../../utils/utils");
 var Enums = require("../../enums");
 var StreamBuf = require("../../utils/stream-buf");
 var StutteredPipe = require("../../utils/stuttered-pipe");
 var FlowControl = require("../../utils/flow-control");
 var AutoDrain = require("../../utils/auto-drain");

 var RelType = require("../../xlsx/rel-type");
 var StyleManager = require("../../xlsx/stylemanager");
 var SharedStrings = require("../../utils/shared-strings");

 var WorksheetReader = require("./worksheet-reader");
 var HyperlinkReader = require("./hyperlink-reader");

// private var
var sharedStrings = [];

var WorkbookReader = module.exports = function(options) {
	this.options = options = options || {};
	
	// until we actually parse a styles.xml file, just assume we're not handling styles
	// (but we do need to handle dates)
	this.styles = new StyleManager.Mock();

	
	// worksheet readers, indexed by sheetNo
	this.worksheetReaders = {};
	
	// hyperlink readers, indexed by sheetNo
	this.hyperlinkReaders = {};
	
	// count the open readers
	this.readers = 0;

}

utils.inherits(WorkbookReader, events.EventEmitter, {
	_getStream: function(input, autoClose) {
		if (autoClose != false) autoClose = true;
		if (input instanceof Stream.Readable) {
			return input;
		}
		if (_.isString(input)) {
			return fs.createReadStream(input, {autoClose: autoClose});
		}
		throw new Error("Could not recognise input");
	},
	
	get flowControl() {
		if (!this._flowControl) {
			this._flowControl = new FlowControl(this.options);
		}
		return this._flowControl;
	},
	
	options: {
		entries: ["emit"],
		sharedStrings: ["cache", "emit"],
		styles: ["cache"],
		hyperlinks: ["cache", "emit"],
		worksheets: ["emit"]
	},
	read: function(file, options) {
		var self = this;

		var zipEntries = [];
		var zip = new StreamZip({  
			file: file,
			storeEntries: true    
		});
		zip.on('error', function(err) {
			throw new Error("compression error:" + err);
		});

		zip.on('entry', function(entry) {
			zipEntries.push(entry);
		});

		zip.on('ready', function() {
			async.series([
				function(callbackSeries){
					var erred = false;
					zip.stream('xl/sharedStrings.xml', function(err, stm) {
						stm.on('error', function(err) {
							erred = true;
							setImmediate(callbackSeries, err, '_parseSharedStrings')
						})
						stm.on('end', function() {
							if (!erred) setImmediate(callbackSeries, null, '_parseSharedStrings')
						})
						self._parseSharedStrings(stm, options);
					});
				},
				function(callbackSeries){
					var erred = false;
					zip.stream('xl/styles.xml', function(err, stm) {
						stm.on('error', function(err) {
							erred = true;
							setImmediate(callbackSeries, err, '_parseStyles')
						})
						stm.on('end', function() {
							if (!erred) setImmediate(callbackSeries, null, '_parseStyles')
						})
						self._parseStyles(stm, options);
					});
				},
				function(callbackSeries){
					var erred = false;
					async.eachSeries(zipEntries, function iterator(entry, callbackZipEntries) {
						if (entry.name.match(/xl\/worksheets\/_rels\/sheet\d+\.xml.rels/)) {
							var match = entry.name.match(/xl\/worksheets\/_rels\/sheet(\d+)\.xml.rels/)
							var sheetNo = match[1];
							zip.stream(entry.name, function(err, stm) {
								stm.on('error', function(err) {
									erred = true;
									setImmediate(callbackZipEntries, err)
								})
								stm.on('end', function() {
									if (!erred) setImmediate(callbackZipEntries)
								})
								self._parseHyperlinks(stm, sheetNo, options);
							});
						}else{
							if (!erred) setImmediate(callbackZipEntries)
						}
					}, function(err){
						setImmediate(callbackSeries, err, '_parseHyperlinks')
					})
				},
				function(callbackSeries){
					var erred = false;
					async.eachSeries(zipEntries, function iterator(entry, callbackZipEntries) {
						if (entry.name.match(/xl\/worksheets\/sheet\d+\.xml/)) {
							var match = entry.name.match(/xl\/worksheets\/sheet(\d+)\.xml/)
							var sheetNo = match[1];
							zip.stream(entry.name, function(err, stm) {
								stm.on('error', function(err) {
									erred = true;
									setImmediate(callbackZipEntries, err)
								})
								stm.on('end', function() {
									if (!erred) setImmediate(callbackZipEntries)
								})
								self._parseWorksheet(stm, sheetNo, options);
							});
						}else{
							setImmediate(callbackZipEntries)
						}
					}, function(err){
						setImmediate(callbackSeries, err, '_parseWorksheet')
					})
				},
			],
			function(err, results){
				console.log("queue finished.");
			    if (self.listenerCount("finished")) self.emit("finished");
			});
		});
	},
	_parseSharedStrings: function(entry, options) {
		var self = this;

		var parser = Sax.createStream(true, {});
		var inT = false;
		var t = null;
		var index = 0;
		parser.on('opentag', function(node) {
			if (node.name == "t") {
				t = null;
				inT = true;
			}
		});
		parser.on('closetag', function (name) {
			if (inT && (name == "t")) {
				sharedStrings.push(t);
				t = null;
			}
		});
		parser.on('text', function (text) {
			t = t ? t + text : text;
		});
		parser.on('error', function (error) {
			self.emit('error', error);
		});
		entry.pipe(parser);
	},
	_getSharedString: function(id) {
		if (id in sharedStrings){
			return sharedStrings[id];
		}else{
			throw new Error("Shared string not found: " + id);
		}
	},
	_parseStyles: function(entry, options) {
		var self = this;
		if (this.listenerCount("entry")) this.emit("entry", {type: "styles"});

		this.styles = new StyleManager();
		this.styles.parse(entry);
	},
	_getReader: function(Type, collection, sheetNo) {
		var self = this;
		var reader = collection[sheetNo];
		if (!reader) {
			reader = new Type(this, sheetNo);
			self.readers++;
			reader.on("finished", function() {
				console.log("reader finished!");
				if (!--self.readers) {
					if (self.atEnd) {
						if (self.listenerCount("finished")) self.emit("finished");
					}
				}
			});
			collection[sheetNo] = reader;
		}
		return reader;
	},
	_parseWorksheet: function(entry, sheetNo, options) {
		if (this.listenerCount("entry")) this.emit("entry", {type: "worksheet", id: sheetNo});
		var worksheetReader = this._getReader(WorksheetReader, this.worksheetReaders, sheetNo);
		if (this.listenerCount("worksheet")) this.emit("worksheet", worksheetReader);
		worksheetReader.read(entry, this.hyperlinkReaders[sheetNo]);
	},
	_parseHyperlinks: function(entry, sheetNo, options) {
		if (this.listenerCount("entry")) this.emit("entry", {type: "hyerlinks", id: sheetNo});
		var hyperlinksReader = this._getReader(HyperlinkReader, this.hyperlinkReaders, sheetNo);
		if (this.listenerCount("hyperlinks")) this.emit("hyperlinks", hyperlinksReader);
		hyperlinksReader.read(entry, options);
	}
});
