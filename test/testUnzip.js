const fs = require('fs');
const unzip = require('node-unzip-2');

const filename = process.argv[2];
console.log(`Reading ${filename}`);
fs.createReadStream(filename)
  .pipe(unzip.Parse())
  .on('entry', entry => {
    const buf = entry.read();
    console.log(entry.path, buf.length);
  })
  .on('close', () => {
    console.log('Finished');
  })
  .on('error', error => {
    console.log(`Error: ${error.message}`);
  });
