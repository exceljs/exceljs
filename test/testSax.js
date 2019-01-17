const fs = require('fs');
const events = require('events');
const Sax = require('sax');
const unzip = require('node-unzip-2');

const filename = process.argv[2];

const Row = function(r) {
  this.number = r;
  this.cells = {};
};
Row.prototype = {
  add(cell) {
    this.cells[cell.address] = cell;
  },
};

let count = 0;
const e = new events.EventEmitter();
e.on('row', row => {
  count++;

  if (count % 1000 === 0) {
    process.stdout.write(`Count:${count}\u001b[0G`); // "\033[0G"
  }
});
e.on('finished', () => {
  console.log(`Finished worksheet: ${count}`);
});

const zip = unzip.Parse();
zip.on('entry', entry => {
  if (entry.path.match(/xl\/worksheets\/sheet\d+[.]xml/)) {
    parseSheet(entry, e);
  }
});

function parseSheet(entry, emitter) {
  const parser = Sax.createStream(true, {});
  let row = null;
  let cell = null;
  let current = null;
  parser.on('opentag', node => {
    switch (node.name) {
      case 'row': {
        const r = parseInt(node.attributes.r, 10);
        row = new Row(r);
        break;
      }
      case 'c':
        cell = {
          address: node.attributes.r,
          s: parseInt(node.attributes.s, 10),
          t: node.attributes.t,
        };
        break;
      case 'v':
        current = cell.v = { text: '' };
        break;
      case 'f':
        current = cell.f = { text: '' };
        break;
      default:
    }
  });
  parser.on('text', text => {
    if (current) {
      current.text += text;
    }
  });
  parser.on('closetag', name => {
    switch (name) {
      case 'row':
        emitter.emit('row', row);
        row = null;
        break;
      case 'c':
        row.add(cell);
        break;
      default:
    }
  });
  parser.on('end', () => {
    e.emit('finished');
  });
  entry.pipe(parser);
}

const stream = fs.createReadStream(filename);
let eod = false;
stream.on('end', () => {
  eod = true;
});
function schedule() {
  setImmediate(() => {
    if (!eod) {
      const data = stream.read(16384);
      if (data && data.length) {
        zip.write(data);
      }
      schedule();
    }
  });
}

// stream.pipe(zip);
schedule();
