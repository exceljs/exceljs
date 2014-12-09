var fs = require("fs");
var unzip = require("unzip");
var _ = require("underscore");

var ss = null;

var filename = process.argv[2];
console.log("Reading " + filename);
fs.createReadStream(filename)
    .pipe(unzip.Parse())
    .on('entry',function (entry) {
        console.log(entry.path);
        if (entry.path == "xl/sharedStrings.xml") {
            ss = entry;
        }
    }).on('close', function () {
        var parser = require("sax").createStream(true, {});
        parser.on('opentag', function(node) {
            var line = [];
            line.push("<" + node.name);
            //console.log(JSON.stringify(node.attributes));
            _.each(node.attributes, function(value, name) {
                line.push(name + '="' + value + '"');
            });
            console.log(line.join(" ") + ">");
        });
        parser.on('closetag', function (name) {
            console.log("</" + name + ">");
        });
        parser.on('text', function (text) {
            console.log(text);
        });
        parser.on('error', function (err) {
            console.error(err);
        });
        ss.pipe(parser);
    });
