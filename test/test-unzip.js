var fs = require('fs');
var unzip = require('unzip2');
var Sax = require('sax');

var filename = process.argv[2];

class Writable {
  constructor(name) {
    this.name = name;
    this.buffs = [];
  }
  write(buf) {
    this.buffs.push(buf);
  }
  end(buf) {
    if (buf) {
      this.buffs.push(buf);
    }
    if (this.name[this.name.length - 1] !== '/') {
      var length = this.buffs.reduce((t, b) => t + b.length, 0);
      console.log(this.name, length);
    }
  }
  on() {}
  once() {}
  emit() {}
}

var zipEntries = {};
var zipStream = unzip.Parse();
zipStream.on('entry',function (entry) {
  if (entry.path === 'xl/media/image1.png') {
    console.log('entry', entry);
  }
  var writable = new Writable(entry.path);
  entry.pipe(writable);
});
zipStream.on('close', function () {
  console.log('zip close')
});

var readStream = fs.createReadStream(filename);
readStream.pipe(zipStream);