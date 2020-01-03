'use strict';

var expect = require('chai').expect;
var Promish = require('promish');
var fs = require('fs');
var path = require('path');

var StreamBuf = require('../../../lib/utils/stream-buf');
var StringBuf = require('../../../lib/utils/string-buf');

describe('StreamBuf', function() {
  // StreamBuf is designed as a general-purpose writable-readable stream
  // However its use in ExcelJS is primarily as a memory buffer between
  // the streaming writers and the archive, hence the tests here will
  // focus just on that.
  it('writes strings as UTF8', function() {
    var stream = new StreamBuf();
    stream.write('Hello, World!');
    var chunk = stream.read();
    expect(chunk instanceof Buffer).to.be.ok();
    expect(chunk.toString('UTF8')).to.equal('Hello, World!');
  });

  it('writes StringBuf chunks', function() {
    var stream = new StreamBuf();
    var strBuf = new StringBuf({size: 64});
    strBuf.addText('Hello, World!');
    stream.write(strBuf);
    var chunk = stream.read();
    expect(chunk instanceof Buffer).to.be.ok();
    expect(chunk.toString('UTF8')).to.equal('Hello, World!');
  });

  it('signals end', function(done) {
    var stream = new StreamBuf();
    stream.on('finish', function() {
      done();
    });
    stream.write('Hello, World!');
    stream.end();
  });

  it('handles buffers', function() {
    return new Promish((resolve, reject) => {
      var s = fs.createReadStream(path.join(__dirname, 'data/image1.png'));
      var sb = new StreamBuf();
      sb.on('finish', () => {
        var buf = sb.toBuffer();
        expect(buf.length).to.equal(1672);
        resolve();
      });
      sb.on('error', reject);
      s.pipe(sb);
    });
  });
});
