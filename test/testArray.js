var _ = require('underscore');

var a = [];
a[3] = 'three';
//a[2] = 'two';
a[4] = 'four';
a[1] = 'one';
a[5] = 'five';
a[0] = 'zero';

_.each(a, function(i) {
    console.log(i);
})