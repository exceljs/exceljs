/* eslint-disable consistent-return,no-mixed-operators,no-bitwise,max-classes-per-file */
/**
 * @license node-stream-zip | (c) 2020 Antelle | https://github.com/antelle/node-stream-zip/blob/master/LICENSE
 * Portions copyright https://github.com/cthackers/adm-zip | https://raw.githubusercontent.com/cthackers/adm-zip/master/LICENSE
 */

// region Deps

const events = require('events');
const zlib = require('zlib');
const stream = require('readable-stream');

// endregion

// region Constants
const consts = {
    /* The local file header */
    LOCHDR: 30, // LOC header size
    LOCSIG: 0x04034b50, // "PK\003\004"
    LOCVER: 4, // version needed to extract
    LOCFLG: 6, // general purpose bit flag
    LOCHOW: 8, // compression method
    LOCTIM: 10, // modification time (2 bytes time, 2 bytes date)
    LOCCRC: 14, // uncompressed file crc-32 value
    LOCSIZ: 18, // compressed size
    LOCLEN: 22, // uncompressed size
    LOCNAM: 26, // filename length
    LOCEXT: 28, // extra field length

    /* The Data descriptor */
    EXTSIG: 0x08074b50, // "PK\007\008"
    EXTHDR: 16, // EXT header size
    EXTCRC: 4, // uncompressed file crc-32 value
    EXTSIZ: 8, // compressed size
    EXTLEN: 12, // uncompressed size

    /* The central directory file header */
    CENHDR: 46, // CEN header size
    CENSIG: 0x02014b50, // "PK\001\002"
    CENVEM: 4, // version made by
    CENVER: 6, // version needed to extract
    CENFLG: 8, // encrypt, decrypt flags
    CENHOW: 10, // compression method
    CENTIM: 12, // modification time (2 bytes time, 2 bytes date)
    CENCRC: 16, // uncompressed file crc-32 value
    CENSIZ: 20, // compressed size
    CENLEN: 24, // uncompressed size
    CENNAM: 28, // filename length
    CENEXT: 30, // extra field length
    CENCOM: 32, // file comment length
    CENDSK: 34, // volume number start
    CENATT: 36, // internal file attributes
    CENATX: 38, // external file attributes (host system dependent)
    CENOFF: 42, // LOC header offset

    /* The entries in the end of central directory */
    ENDHDR: 22, // END header size
    ENDSIG: 0x06054b50, // "PK\005\006"
    ENDSIGFIRST: 0x50,
    ENDSUB: 8, // number of entries on this disk
    ENDTOT: 10, // total number of entries
    ENDSIZ: 12, // central directory size in bytes
    ENDOFF: 16, // offset of first CEN header
    ENDCOM: 20, // zip file comment length
    MAXFILECOMMENT: 0xffff,

    /* The entries in the end of ZIP64 central directory locator */
    ENDL64HDR: 20, // ZIP64 end of central directory locator header size
    ENDL64SIG: 0x07064b50, // ZIP64 end of central directory locator signature
    ENDL64SIGFIRST: 0x50,
    ENDL64OFS: 8, // ZIP64 end of central directory offset

    /* The entries in the end of ZIP64 central directory */
    END64HDR: 56, // ZIP64 end of central directory header size
    END64SIG: 0x06064b50, // ZIP64 end of central directory signature
    END64SIGFIRST: 0x50,
    END64SUB: 24, // number of entries on this disk
    END64TOT: 32, // total number of entries
    END64SIZ: 40,
    END64OFF: 48,

    /* Compression methods */
    STORED: 0, // no compression
    SHRUNK: 1, // shrunk
    REDUCED1: 2, // reduced with compression factor 1
    REDUCED2: 3, // reduced with compression factor 2
    REDUCED3: 4, // reduced with compression factor 3
    REDUCED4: 5, // reduced with compression factor 4
    IMPLODED: 6, // imploded
    // 7 reserved
    DEFLATED: 8, // deflated
    ENHANCED_DEFLATED: 9, // deflate64
    PKWARE: 10, // PKWare DCL imploded
    // 11 reserved
    BZIP2: 12, //  compressed using BZIP2
    // 13 reserved
    LZMA: 14, // LZMA
    // 15-17 reserved
    IBM_TERSE: 18, // compressed using IBM TERSE
    IBM_LZ77: 19, // IBM LZ77 z

    /* General purpose bit flag */
    FLG_ENC: 0, // encrypted file
    FLG_COMP1: 1, // compression option
    FLG_COMP2: 2, // compression option
    FLG_DESC: 4, // data descriptor
    FLG_ENH: 8, // enhanced deflation
    FLG_STR: 16, // strong encryption
    FLG_LNG: 1024, // language encoding
    FLG_MSK: 4096, // mask header values
    FLG_ENTRY_ENC: 1,

    /* 4.5 Extensible data fields */
    EF_ID: 0,
    EF_SIZE: 2,

    /* Header IDs */
    ID_ZIP64: 0x0001,
    ID_AVINFO: 0x0007,
    ID_PFS: 0x0008,
    ID_OS2: 0x0009,
    ID_NTFS: 0x000a,
    ID_OPENVMS: 0x000c,
    ID_UNIX: 0x000d,
    ID_FORK: 0x000e,
    ID_PATCH: 0x000f,
    ID_X509_PKCS7: 0x0014,
    ID_X509_CERTID_F: 0x0015,
    ID_X509_CERTID_C: 0x0016,
    ID_STRONGENC: 0x0017,
    ID_RECORD_MGT: 0x0018,
    ID_X509_PKCS7_RL: 0x0019,
    ID_IBM1: 0x0065,
    ID_IBM2: 0x0066,
    ID_POSZIP: 0x4690,

    EF_ZIP64_OR_32: 0xffffffff,
    EF_ZIP64_OR_16: 0xffff,
};

// endregion

// region BlobZipStream
class BlobZipStream extends events.EventEmitter {
    constructor(config) {
        super();

        /**
         * @type Blob
         */
        let fd = config.file;
        let chunkSize;
        const ready = false;
        const that = this;
        let op;
        let centralDirectory;
        let closed;

        const entries = config.storeEntries !== false ? {} : null;

        const fileSize = fd.size;
        chunkSize = config.chunkSize || Math.round(fileSize / 1000);
        chunkSize = Math.max(
          Math.min(chunkSize, Math.min(128 * 1024, fileSize)),
          Math.min(1024, fileSize)
        );
        readCentralDirectory();

        function readUntilFoundCallback(err, bytesRead) {
            if (err || !bytesRead) {
                return that.emit('error', err || new Error('Archive read error'));
            }
            let pos = op.lastPos;
            let bufferPosition = pos - op.win.position;
            const {buffer} = op.win;
            const {minPos} = op;
            while (--pos >= minPos && --bufferPosition >= 0) {
                if (buffer.length - bufferPosition >= 4 && buffer[bufferPosition] === op.firstByte) {
                    // quick check first signature byte
                    if (buffer.readUInt32LE(bufferPosition) === op.sig) {
                        op.lastBufferPosition = bufferPosition;
                        op.lastBytesRead = bytesRead;
                        op.complete();
                        return;
                    }
                }
            }
            if (pos === minPos) {
                return that.emit('error', new Error('Bad archive'));
            }
            op.lastPos = pos + 1;
            op.chunkSize *= 2;
            if (pos <= minPos) {
                return that.emit('error', new Error('Bad archive'));
            }
            const expandLength = Math.min(op.chunkSize, pos - minPos);
            op.win.expandLeft(expandLength, readUntilFoundCallback);
        }

        function readCentralDirectory() {
            const totalReadLength = Math.min(consts.ENDHDR + consts.MAXFILECOMMENT, fileSize);
            op = {
                win: new FileWindowBuffer(fd),
                totalReadLength,
                minPos: fileSize - totalReadLength,
                lastPos: fileSize,
                chunkSize: Math.min(1024, chunkSize),
                firstByte: consts.ENDSIGFIRST,
                sig: consts.ENDSIG,
                complete: readCentralDirectoryComplete,
            };
            op.win.read(fileSize - op.chunkSize, op.chunkSize, readUntilFoundCallback);
        }

        function readCentralDirectoryComplete() {
            const {buffer} = op.win;
            const pos = op.lastBufferPosition;
            try {
                centralDirectory = new CentralDirectoryHeader();
                centralDirectory.read(buffer.slice(pos, pos + consts.ENDHDR));
                centralDirectory.headerOffset = op.win.position + pos;
                if (centralDirectory.commentLength) {
                    that.comment = buffer
                            .slice(
                                    pos + consts.ENDHDR,
                                    pos + consts.ENDHDR + centralDirectory.commentLength
                            )
                            .toString();
                } else {
                    that.comment = null;
                }
                that.entriesCount = centralDirectory.volumeEntries;
                that.centralDirectory = centralDirectory;
                if (
                        (centralDirectory.volumeEntries === consts.EF_ZIP64_OR_16 &&
                                centralDirectory.totalEntries === consts.EF_ZIP64_OR_16) ||
                        centralDirectory.size === consts.EF_ZIP64_OR_32 ||
                        centralDirectory.offset === consts.EF_ZIP64_OR_32
                ) {
                    readZip64CentralDirectoryLocator();
                } else {
                    op = {};
                    readEntries();
                }
            } catch (err) {
                that.emit('error', err);
            }
        }

        function readZip64CentralDirectoryLocator() {
            const length = consts.ENDL64HDR;
            if (op.lastBufferPosition > length) {
                op.lastBufferPosition -= length;
                readZip64CentralDirectoryLocatorComplete();
            } else {
                op = {
                    win: op.win,
                    totalReadLength: length,
                    minPos: op.win.position - length,
                    lastPos: op.win.position,
                    chunkSize: op.chunkSize,
                    firstByte: consts.ENDL64SIGFIRST,
                    sig: consts.ENDL64SIG,
                    complete: readZip64CentralDirectoryLocatorComplete,
                };
                op.win.read(op.lastPos - op.chunkSize, op.chunkSize, readUntilFoundCallback);
            }
        }

        function readZip64CentralDirectoryLocatorComplete() {
            const {buffer} = op.win;
            const locHeader = new CentralDirectoryLoc64Header();
            locHeader.read(
                    buffer.slice(op.lastBufferPosition, op.lastBufferPosition + consts.ENDL64HDR)
            );
            const readLength = fileSize - locHeader.headerOffset;
            op = {
                win: op.win,
                totalReadLength: readLength,
                minPos: locHeader.headerOffset,
                lastPos: op.lastPos,
                chunkSize: op.chunkSize,
                firstByte: consts.END64SIGFIRST,
                sig: consts.END64SIG,
                complete: readZip64CentralDirectoryComplete,
            };
            op.win.read(fileSize - op.chunkSize, op.chunkSize, readUntilFoundCallback);
        }

        function readZip64CentralDirectoryComplete() {
            const {buffer} = op.win;
            const zip64cd = new CentralDirectoryZip64Header();
            zip64cd.read(buffer.slice(op.lastBufferPosition, op.lastBufferPosition + consts.END64HDR));
            that.centralDirectory.volumeEntries = zip64cd.volumeEntries;
            that.centralDirectory.totalEntries = zip64cd.totalEntries;
            that.centralDirectory.size = zip64cd.size;
            that.centralDirectory.offset = zip64cd.offset;
            that.entriesCount = zip64cd.volumeEntries;
            op = {};
            readEntries();
        }

        function readEntries() {
            op = {
                win: new FileWindowBuffer(fd),
                pos: centralDirectory.offset,
                chunkSize,
                entriesLeft: centralDirectory.volumeEntries,
            };
            op.win.read(op.pos, Math.min(chunkSize, fileSize - op.pos), readEntriesCallback);
        }

        function readEntriesCallback(err, bytesRead) {
            if (err || !bytesRead) {
                return that.emit('error', err || new Error('Entries read error'));
            }
            let bufferPos = op.pos - op.win.position;
            let {entry} = op;
            const {buffer} = op.win;
            const bufferLength = buffer.length;
            try {
                while (op.entriesLeft > 0) {
                    if (!entry) {
                        entry = new ZipEntry();
                        entry.readHeader(buffer, bufferPos);
                        entry.headerOffset = op.win.position + bufferPos;
                        op.entry = entry;
                        op.pos += consts.CENHDR;
                        bufferPos += consts.CENHDR;
                    }
                    const entryHeaderSize = entry.fnameLen + entry.extraLen + entry.comLen;
                    const advanceBytes = entryHeaderSize + (op.entriesLeft > 1 ? consts.CENHDR : 0);
                    if (bufferLength - bufferPos < advanceBytes) {
                        op.win.moveRight(chunkSize, readEntriesCallback, bufferPos);
                        op.move = true;
                        return;
                    }
                    entry.read(buffer, bufferPos);
                    if (!config.skipEntryNameValidation) {
                        entry.validateName();
                    }
                    if (entries) {
                        entries[entry.name] = entry;
                    }
                    that.emit('entry', entry);
                    op.entry = entry = null;
                    op.entriesLeft--;
                    op.pos += entryHeaderSize;
                    bufferPos += entryHeaderSize;
                }
                that.emit('ready');
            } catch (err0) {
                that.emit('error', err0);
            }
        }

        function checkEntriesExist() {
            if (!entries) {
                throw new Error('storeEntries disabled');
            }
        }

        Object.defineProperty(this, 'ready', {
            get() {
                return ready;
            },
        });

        this.entry = function(name) {
            checkEntriesExist();
            return entries[name];
        };

        this.entries = function() {
            checkEntriesExist();
            return entries;
        };

        this.stream = function(entry0, callback) {
            return this.openEntry(
                    entry0,
                    (err, entry) => {
                        if (err) {
                            return callback(err);
                        }
                        const offset = dataOffset(entry);
                        let entryStream = new EntryDataReaderStream(fd, offset, entry.compressedSize);
                        if (entry.method === consts.STORED) {
                            // nothing to do
                        } else if (entry.method === consts.DEFLATED) {
                            entryStream = entryStream.pipe(zlib.createInflateRaw());
                        } else {
                            return callback(new Error(`Unknown compression method: ${  entry.method}`));
                        }
                        if (canVerifyCrc(entry)) {
                            entryStream = entryStream.pipe(
                                    new EntryVerifyStream(entryStream, entry.crc, entry.size)
                            );
                        }
                        callback(null, entryStream);
                    },
                    false
            );
        };


        this.openEntry = function(entry, callback, sync) {
            if (typeof entry === 'string') {
                checkEntriesExist();
                entry = entries[entry];
                if (!entry) {
                    return callback(new Error('Entry not found'));
                }
            }
            if (!entry.isFile) {
                return callback(new Error('Entry is not file'));
            }
            if (!fd) {
                return callback(new Error('Archive closed'));
            }
            const buffer = Buffer.alloc(consts.LOCHDR);
            new FsRead(fd, buffer, 0, buffer.length, entry.offset, err => {
                if (err) {
                    return callback(err);
                }
                let readEx;
                try {
                    entry.readDataHeader(buffer);
                    if (entry.encrypted) {
                        readEx = new Error('Entry encrypted');
                    }
                } catch (ex) {
                    readEx = ex;
                }
                callback(readEx, entry);
            }).read(sync);
        };

        function dataOffset(entry) {
            return entry.offset + consts.LOCHDR + entry.fnameLen + entry.extraLen;
        }

        function canVerifyCrc(entry) {
            // if bit 3 (0x08) of the general-purpose flags field is set,
            // then the CRC-32 and file sizes are not known when the header is written
            return (entry.flags & 0x8) !== 0x8;
        }

        this.close = function(callback) {
            if (closed || !fd) {
                closed = true;
                if (callback)
                    callback();
            } else {
                closed = true;
                fd = null;
                if (callback)
                    callback();
            }
        };

        const originalEmit = events.EventEmitter.prototype.emit;
        this.emit = function(...args) {
            if (!closed) {
                return originalEmit.call(this, ...args);
            }
        };
    }

}


// endregion

class CentralDirectoryHeader {
    read(data) {
        if (data.length !== consts.ENDHDR || data.readUInt32LE(0) !== consts.ENDSIG) {
            throw new Error('Invalid central directory');
        }
        // number of entries on this volume
        this.volumeEntries = data.readUInt16LE(consts.ENDSUB);
        // total number of entries
        this.totalEntries = data.readUInt16LE(consts.ENDTOT);
        // central directory size in bytes
        this.size = data.readUInt32LE(consts.ENDSIZ);
        // offset of first CEN header
        this.offset = data.readUInt32LE(consts.ENDOFF);
        // zip file comment length
        this.commentLength = data.readUInt16LE(consts.ENDCOM);
    }
}

class CentralDirectoryLoc64Header {
    read(data) {
        if (data.length !== consts.ENDL64HDR || data.readUInt32LE(0) !== consts.ENDL64SIG) {
            throw new Error('Invalid zip64 central directory locator');
        }
        // ZIP64 EOCD header offset
        this.headerOffset = readUInt64LE(data, consts.ENDSUB);
    }
}

class CentralDirectoryZip64Header {
    read(data) {
        if (data.length !== consts.END64HDR || data.readUInt32LE(0) !== consts.END64SIG) {
            throw new Error('Invalid central directory');
        }
        // number of entries on this volume
        this.volumeEntries = readUInt64LE(data, consts.END64SUB);
        // total number of entries
        this.totalEntries = readUInt64LE(data, consts.END64TOT);
        // central directory size in bytes
        this.size = readUInt64LE(data, consts.END64SIZ);
        // offset of first CEN header
        this.offset = readUInt64LE(data, consts.END64OFF);
    }
}

class ZipEntry {
    readHeader(data, offset) {
        // data should be 46 bytes and start with "PK 01 02"
        if (data.length < offset + consts.CENHDR || data.readUInt32LE(offset) !== consts.CENSIG) {
            throw new Error('Invalid entry header');
        }
        // version made by
        this.verMade = data.readUInt16LE(offset + consts.CENVEM);
        // version needed to extract
        this.version = data.readUInt16LE(offset + consts.CENVER);
        // encrypt, decrypt flags
        this.flags = data.readUInt16LE(offset + consts.CENFLG);
        // compression method
        this.method = data.readUInt16LE(offset + consts.CENHOW);
        // modification time (2 bytes time, 2 bytes date)
        const timebytes = data.readUInt16LE(offset + consts.CENTIM);
        const datebytes = data.readUInt16LE(offset + consts.CENTIM + 2);
        this.time = parseZipTime(timebytes, datebytes);

        // uncompressed file crc-32 value
        this.crc = data.readUInt32LE(offset + consts.CENCRC);
        // compressed size
        this.compressedSize = data.readUInt32LE(offset + consts.CENSIZ);
        // uncompressed size
        this.size = data.readUInt32LE(offset + consts.CENLEN);
        // filename length
        this.fnameLen = data.readUInt16LE(offset + consts.CENNAM);
        // extra field length
        this.extraLen = data.readUInt16LE(offset + consts.CENEXT);
        // file comment length
        this.comLen = data.readUInt16LE(offset + consts.CENCOM);
        // volume number start
        this.diskStart = data.readUInt16LE(offset + consts.CENDSK);
        // internal file attributes
        this.inattr = data.readUInt16LE(offset + consts.CENATT);
        // external file attributes
        this.attr = data.readUInt32LE(offset + consts.CENATX);
        // LOC header offset
        this.offset = data.readUInt32LE(offset + consts.CENOFF);
    }

    readDataHeader(data) {
        // 30 bytes and should start with "PK\003\004"
        if (data.readUInt32LE(0) !== consts.LOCSIG) {
            throw new Error('Invalid local header');
        }
        // version needed to extract
        this.version = data.readUInt16LE(consts.LOCVER);
        // general purpose bit flag
        this.flags = data.readUInt16LE(consts.LOCFLG);
        // compression method
        this.method = data.readUInt16LE(consts.LOCHOW);
        // modification time (2 bytes time ; 2 bytes date)
        const timebytes = data.readUInt16LE(consts.LOCTIM);
        const datebytes = data.readUInt16LE(consts.LOCTIM + 2);
        this.time = parseZipTime(timebytes, datebytes);

        // uncompressed file crc-32 value
        this.crc = data.readUInt32LE(consts.LOCCRC) || this.crc;
        // compressed size
        const compressedSize = data.readUInt32LE(consts.LOCSIZ);
        if (compressedSize && compressedSize !== consts.EF_ZIP64_OR_32) {
            this.compressedSize = compressedSize;
        }
        // uncompressed size
        const size = data.readUInt32LE(consts.LOCLEN);
        if (size && size !== consts.EF_ZIP64_OR_32) {
            this.size = size;
        }
        // filename length
        this.fnameLen = data.readUInt16LE(consts.LOCNAM);
        // extra field length
        this.extraLen = data.readUInt16LE(consts.LOCEXT);
    }

    read(data, offset) {
        this.name = data.slice(offset, (offset += this.fnameLen)).toString();
        const lastChar = data[offset - 1];
        this.isDirectory = lastChar === 47 || lastChar === 92;

        if (this.extraLen) {
            this.readExtra(data, offset);
            offset += this.extraLen;
        }
        this.comment = this.comLen ? data.slice(offset, offset + this.comLen).toString() : null;
    }

    validateName() {
        if (/\\|^\w+:|^\/|(^|\/)\.\.(\/|$)/.test(this.name)) {
            throw new Error(`Malicious entry: ${  this.name}`);
        }
    }

    readExtra(data, offset) {
        let signature;
        let size;
        const maxPos = offset + this.extraLen;
        while (offset < maxPos) {
            signature = data.readUInt16LE(offset);
            offset += 2;
            size = data.readUInt16LE(offset);
            offset += 2;
            if (consts.ID_ZIP64 === signature) {
                this.parseZip64Extra(data, offset, size);
            }
            offset += size;
        }
    }

    parseZip64Extra(data, offset, length) {
        if (length >= 8 && this.size === consts.EF_ZIP64_OR_32) {
            this.size = readUInt64LE(data, offset);
            offset += 8;
            length -= 8;
        }
        if (length >= 8 && this.compressedSize === consts.EF_ZIP64_OR_32) {
            this.compressedSize = readUInt64LE(data, offset);
            offset += 8;
            length -= 8;
        }
        if (length >= 8 && this.offset === consts.EF_ZIP64_OR_32) {
            this.offset = readUInt64LE(data, offset);
            offset += 8;
            length -= 8;
        }
        if (length >= 4 && this.diskStart === consts.EF_ZIP64_OR_16) {
            this.diskStart = data.readUInt32LE(offset);
            // offset += 4; length -= 4;
        }
    }

    get encrypted() {
        return (this.flags & consts.FLG_ENTRY_ENC) === consts.FLG_ENTRY_ENC;
    }

    get isFile() {
        return !this.isDirectory;
    }
}

// region FsRead
class FsRead {
    constructor(fd, buffer, offset, length, position, callback) {
        /**
         * @type Blob
         */
        this.fd = fd;
        /**
         * @type Buffer
         */
        this.buffer = buffer;
        this.offset = offset;
        this.length = length;
        this.position = position;
        this.callback = callback;
        this.bytesRead = 0;
        this.waiting = false;
    }

    read() {
        if (BlobZipStream.debug) {
            // eslint-disable-next-line no-console
            console.log('read', this.position, this.bytesRead, this.length, this.offset);
        }
        this.waiting = true;
        const blob = this.fd.slice(this.position + this.bytesRead,
                this.position + this.bytesRead + this.length - this.bytesRead);
        // eslint-disable-next-line no-undef
        const fr = new FileReader();
        fr.onerror = err => {
            this.readCallback(err);
        };
        fr.onload = () => {
            this.buffer.set(new Uint8Array(fr.result), this.offset + this.bytesRead);
            this.readCallback(null, fr.result.byteLength);
        };
        fr.readAsArrayBuffer(blob);
    }

    readCallback(err, bytesRead) {
        if (typeof bytesRead === 'number')
            this.bytesRead += bytesRead;
        if (err || !bytesRead || this.bytesRead === this.length) {
            this.waiting = false;
            return this.callback(err, this.bytesRead);
        }
        this.read();
    }
}

// endregion

class FileWindowBuffer {
    constructor(fd) {
        this.position = 0;
        this.buffer = Buffer.alloc(0);
        this.fd = fd;
        this.fsOp = null;
    }

    checkOp() {
        if (this.fsOp && this.fsOp.waiting) {
            throw new Error('Operation in progress');
        }
    }

    read(pos, length, callback) {
        this.checkOp();
        if (this.buffer.length < length) {
            this.buffer = Buffer.alloc(length);
        }
        this.position = pos;
        this.fsOp = new FsRead(this.fd, this.buffer, 0, length, this.position, callback).read();
    }

    expandLeft(length, callback) {
        this.checkOp();
        this.buffer = Buffer.concat([Buffer.alloc(length), this.buffer]);
        this.position -= length;
        if (this.position < 0) {
            this.position = 0;
        }
        this.fsOp = new FsRead(this.fd, this.buffer, 0, length, this.position, callback).read();
    }

    expandRight(length, callback) {
        this.checkOp();
        const offset = this.buffer.length;
        this.buffer = Buffer.concat([this.buffer, Buffer.alloc(length)]);
        this.fsOp = new FsRead(
                this.fd,
                this.buffer,
                offset,
                length,
                this.position + offset,
                callback
        ).read();
    }

    moveRight(length, callback, shift) {
        this.checkOp();
        if (shift) {
            this.buffer.copy(this.buffer, 0, shift);
        } else {
            shift = 0;
        }
        this.position += shift;
        this.fsOp = new FsRead(
                this.fd,
                this.buffer,
                this.buffer.length - shift,
                shift,
                this.position + this.buffer.length - shift,
                callback
        ).read();
    }
}

class EntryDataReaderStream extends stream.Readable {
    constructor(fd, offset, length) {
        super();
        /**
         * @type Blob
         */
        this.fd = fd;
        this.offset = offset;
        this.length = length;
        this.pos = 0;
    }

    _read(n) {
        const length = Math.min(n, this.length - this.pos);
        if (length > 0) {

            const blob = this.fd.slice(this.offset + this.pos,
                    this.offset + this.pos + Math.min(n, this.length - this.pos));
            // eslint-disable-next-line no-undef
            const fr = new FileReader();
            fr.onload = () => {
                this.readCallback(null, fr.result.byteLength, Buffer.from(fr.result));
            };
            fr.onerror = err => {
                this.readCallback(err);
            };
            fr.readAsArrayBuffer(blob);
        } else {
            this.push(null);
        }
    }

    readCallback(err, bytesRead, buffer) {
        this.pos += bytesRead;
        if (err) {
            this.emit('error', err);
            this.push(null);
        } else if (!bytesRead) {
            this.push(null);
        } else {
            if (bytesRead !== buffer.length) {
                buffer = buffer.slice(0, bytesRead);
            }
            this.push(buffer);
        }
    }
}

class EntryVerifyStream extends stream.Transform {
    constructor(baseStm, crc, size) {
        super();
        this.verify = new CrcVerify(crc, size);
        baseStm.on('error', e => {
            this.emit('error', e);
        });
    }

    _transform(data, encoding, callback) {
        let err;
        try {
            this.verify.data(data);
        } catch (e) {
            err = e;
        }
        callback(err, data);
    }
}

class CrcVerify {
    constructor(crc, size) {
        this.crc = crc;
        this.size = size;
        this.state = {
            crc: ~0,
            size: 0,
        };
    }

    data(data) {
        const crcTable = CrcVerify.getCrcTable();
        let {crc} = this.state;
        let off = 0;
        let len = data.length;
        while (--len >= 0) {
            crc = crcTable[(crc ^ data[off++]) & 0xff] ^ (crc >>> 8);
        }
        this.state.crc = crc;
        this.state.size += data.length;
        if (this.state.size >= this.size) {
            const buf = Buffer.alloc(4);
            buf.writeInt32LE(~this.state.crc & 0xffffffff, 0);
            crc = buf.readUInt32LE(0);
            if (crc !== this.crc) {
                throw new Error('Invalid CRC');
            }
            if (this.state.size !== this.size) {
                throw new Error('Invalid size');
            }
        }
    }

    static getCrcTable() {
        let {crcTable} = CrcVerify;
        if (!crcTable) {
            CrcVerify.crcTable = crcTable = [];
            const b = Buffer.alloc(4);
            for (let n = 0; n < 256; n++) {
                let c = n;
                for (let k = 8; --k >= 0; ) {
                    if ((c & 1) !== 0) {
                        c = 0xedb88320 ^ (c >>> 1);
                    } else {
                        c >>>= 1;
                    }
                }
                if (c < 0) {
                    b.writeInt32LE(c, 0);
                    c = b.readUInt32LE(0);
                }
                crcTable[n] = c;
            }
        }
        return crcTable;
    }
}

// region Util

function parseZipTime(timebytes, datebytes) {
    const timebits = toBits(timebytes, 16);
    const datebits = toBits(datebytes, 16);

    const mt = {
        h: parseInt(timebits.slice(0, 5).join(''), 2),
        m: parseInt(timebits.slice(5, 11).join(''), 2),
        s: parseInt(timebits.slice(11, 16).join(''), 2) * 2,
        Y: parseInt(datebits.slice(0, 7).join(''), 2) + 1980,
        M: parseInt(datebits.slice(7, 11).join(''), 2),
        D: parseInt(datebits.slice(11, 16).join(''), 2),
    };
    const dtStr = `${[mt.Y, mt.M, mt.D].join('-')  } ${  [mt.h, mt.m, mt.s].join(':')  } GMT+0`;
    return new Date(dtStr).getTime();
}

function toBits(dec, size) {
    let b = (dec >>> 0).toString(2);
    while (b.length < size) {
        b = `0${b}`;
    }
    return b.split('');
}

function readUInt64LE(buffer, offset) {
    return buffer.readUInt32LE(offset + 4) * 0x0000000100000000 + buffer.readUInt32LE(offset);
}

// endregion

// region exports

module.exports = BlobZipStream;

// endregion
