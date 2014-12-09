var colCache = require("../lib/colcache");

var arg = process.argv[2];

var match = arg.match(/^[A-Z]+$/);
if (match) {
    var n = colCache.l2n(match[0]);
    var l = colCache.n2l(n);
    console.log(arg + " --> " + n + " --> " + l);
} else {
    var l = colCache.n2l(parseInt(arg));
    var n = colCache.l2n(l);
    console.log(arg + " --> " + l + " --> " + n);
}