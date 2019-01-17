const fs = require('fs');
const unzip = require('unzip2');
const Sax = require('sax');

const filename = process.argv[2];

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
      const length = this.buffs.reduce((t, b) => t + b.length, 0);
      console.log(this.name, length);
    }
  }

  on() {}

  once() {}

  emit() {}
}

const zipEntries = {};
const zipStream = unzip.Parse();
zipStream.on('entry', entry => {
  if (entry.path === 'xl/media/image1.png') {
    console.log('entry', entry);
  }
  const writable = new Writable(entry.path);
  entry.pipe(writable);
});
zipStream.on('close', () => {
  console.log('zip close');
});

const readStream = fs.createReadStream(filename);
readStream.pipe(zipStream);
