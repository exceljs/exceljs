var fs = require('fs');
var _ = require('underscore');
var SharedStrings = require('../lib/utils/shared-strings');

var filename = process.argv[2];

var st = new SharedStrings(null);

var lst = ['Hello', 'Hello', 'World', 'Hello\nWorld!', 'Hello, "World!"'];
_.each(lst, function(item) {
  st.add(item);
});

console.log('Writing sharedstrings to ' + filename);
var stream = fs.createWriteStream(filename);
st.write(stream)
  .then(function() {
    stream.close();
    console.log('Done.');
  });
