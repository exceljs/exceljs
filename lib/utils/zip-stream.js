/* eslint-disable max-classes-per-file */
const events = require('events');
const JSZip = require('jszip');

const StreamBuf = require('./stream-buf');

// The purpose of this module is to wrap the js-zip library into a streaming zip library
// since most of the exceljs code uses streams.
// One day I might find (or build) a properly streaming browser safe zip lib


// =============================================================================
// The ZipReader class
// Unpacks an incoming zip stream
class ZipReader extends events.EventEmitter {
  constructor(options) {
    super();

    this.count = 0;
    this.jsZip = new JSZip();
    this.stream = new StreamBuf();
    this.stream.on('finish', () => {
      this._process();
    });
    this.getEntryType = options.getEntryType || (() => 'string');
  };

  _finished() {
    if (!--this.count) {
      process.nextTick(() => {
        this.emit('finished');
      });
    }
  }

  async _process() {
    try {
      const content = this.stream.read();
      const zip = await this.jsZip.loadAsync(content);
      zip.forEach(async (path, entry) => {
        if (!entry.dir) {
          this.count++;
          try {
            const data = await entry.async(this.getEntryType(path));
            const entryStream = new StreamBuf();
            entryStream.path = path;
            entryStream.write(data);
            entryStream.autodrain = () => {
              this._finished();
            };
            entryStream.on('finish', () => {
              this._finished();
            });

            this.emit('entry', entryStream);
          } catch (error) {
            this.emit('error', error);
          }
        }
      });
    } catch (error) {
      this.emit('error', error);
    }
  }

  // ==========================================================================
  // Stream.Writable interface
  write(data, encoding, callback) {
    if (this.error) {
      if (callback) {
        callback(this.error);
      }
      throw this.error;
    } else {
      return this.stream.write(data, encoding, callback);
    }
  }

  cork() {
    return this.stream.cork();
  }

  uncork() {
    return this.stream.uncork();
  }

  end() {
    return this.stream.end();
  }

  destroy(error) {
    this.emit('finished');
    this.error = error;
  }
}

// =============================================================================
// The ZipWriter class
// Packs streamed data into an output zip stream
class ZipWriter extends events.EventEmitter {
  constructor(options) {
    super();
    this.options = Object.assign({
      type: 'nodebuffer',
      compression: 'DEFLATE',
    }, options);

    this.zip = new JSZip();
    this.stream = new StreamBuf();
  };

  append(data, options) {
    if (options.hasOwnProperty('base64') && options.base64) {
      this.zip.file(options.name, data, {base64: true});
    } else {
      this.zip.file(options.name, data);
    }
  }

  async finalize() {
    const content = await this.zip.generateAsync(this.options);
    this.stream.end(content);
    this.emit('finish');
  }

  // ==========================================================================
  // Stream.Readable interface
  read(size) {
    return this.stream.read(size);
  }

  setEncoding(encoding) {
    return this.stream.setEncoding(encoding);
  }

  pause() {
    return this.stream.pause();
  }

  resume() {
    return this.stream.resume();
  }

  isPaused() {
    return this.stream.isPaused();
  }

  pipe(destination, options) {
    return this.stream.pipe(
      destination,
      options
    );
  }

  unpipe(destination) {
    return this.stream.unpipe(destination);
  }

  unshift(chunk) {
    return this.stream.unshift(chunk);
  }

  wrap(stream) {
    return this.stream.wrap(stream);
  }
}

// =============================================================================

module.exports = {
  ZipReader,
  ZipWriter,
};
