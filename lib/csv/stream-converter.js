'use strict';

// =======================================================================================================
// StreamConverter
//
// convert between encoding schemes in a stream
// Work in Progress - Will complete this at some point
let jconv;

const StreamConverter = (module.exports = function(inner, options) {
  this.inner = inner;

  options = options || {};
  this.innerEncoding = (options.innerEncoding || 'UTF8').toUpperCase();
  this.outerEncoding = (options.outerEncoding || 'UTF8').toUpperCase();

  this.innerBOM = options.innerBOM || null;
  this.outerBOM = options.outerBOM || null;

  this.writeStarted = false;
});

StreamConverter.prototype.convertInwards = function(data) {
  if (data) {
    if (typeof data === 'string') {
      data = new Buffer(data, this.outerEncoding);
    }

    if (this.innerEncoding !== this.outerEncoding) {
      data = jconv.convert(data, this.outerEncoding, this.innerEncoding);
    }
  }

  return data;
};
StreamConverter.prototype.convertOutwards = function(data) {
  if (typeof data === 'string') {
    data = new Buffer(data, this.innerEncoding);
  }

  if (this.innerEncoding !== this.outerEncoding) {
    data = jconv.convert(data, this.innerEncoding, this.outerEncoding);
  }
  return data;
};

StreamConverter.prototype.addListener = function(event, handler) {
  this.inner.addListener(event, handler);
};

StreamConverter.prototype.removeListener = function(event, handler) {
  this.inner.removeListener(event, handler);
};

StreamConverter.prototype.write = function(data, encoding, callback) {
  if (encoding instanceof Function) {
    callback = encoding;
    encoding = undefined;
  }

  if (!this.writeStarted) {
    // if inner encoding has BOM, write it now
    if (this.innerBOM) {
      this.inner.write(this.innerBOM);
    }

    // if outer encoding has BOM, delete it now
    if (this.outerBOM) {
      if (data.length <= this.outerBOM.length) {
        if (callback) {
          callback();
        }
        return;
      }
      const bomless = new Buffer(data.length - this.outerBOM.length);
      data.copy(bomless, 0, this.outerBOM.length, data.length);
      data = bomless;
    }

    this.writeStarted = true;
  }

  this.inner.write(this.convertInwards(data), encoding ? this.innerEncoding : undefined, callback);
};

StreamConverter.prototype.read = function() {
  // TBD
};

StreamConverter.prototype.pipe = function(destination, options) {
  const reverseConverter = new StreamConverter(destination, {
    innerEncoding: this.outerEncoding,
    outerEncoding: this.innerEncoding,
    innerBOM: this.outerBOM,
    outerBOM: this.innerBOM,
  });

  this.inner.pipe(
    reverseConverter,
    options
  );
};

StreamConverter.prototype.close = function() {
  this.inner.close();
};

StreamConverter.prototype.on = function(type, callback) {
  switch (type) {
    case 'data':
      this.inner.on('data', chunk => {
        callback(this.convertOutwards(chunk));
      });
      return this;
    default:
      this.inner.on(type, callback);
      return this;
  }
};

StreamConverter.prototype.once = function(type, callback) {
  this.inner.once(type, callback);
};

StreamConverter.prototype.end = function(chunk, encoding, callback) {
  this.inner.end(this.convertInwards(chunk), this.innerEncoding, callback);
};

StreamConverter.prototype.emit = function(type, value) {
  this.inner.emit(type, value);
};
