var colCache = require('../lib/colcache');

var arg = process.argv[2];

var match = arg.match(/^[A-Z]+$/);
var n,l;
if (match) {
  n = colCache.l2n(match[0]);
  l = colCache.n2l(n);
  console.log(arg + ' --> ' + n + ' --> ' + l);
} else {
  l = colCache.n2l(parseInt(arg, 10));
  n = colCache.l2n(l);
  console.log(arg + ' --> ' + l + ' --> ' + n);
}