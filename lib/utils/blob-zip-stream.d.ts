import {EventEmitter} from "events";

interface StreamZipOptions {
    /**
     * Blob to read
     */
    file: Blob

    /**
     * You will be able to work with entries inside zip archive,
     * otherwise the only way to access them is entry event
     * @default true
     */
    storeEntries?: boolean

    /**
     * By default, entry name is checked for malicious characters, like ../ or c:\123,
     * pass this flag to disable validation error
     * @default false
     */
    skipEntryNameValidation?: boolean

    /**
     * Filesystem read chunk size
     * @default automatic based on file size
     */
    chunkSize?: number
}

interface ZipEntry {
    /**
     * file name
     */
    name: string;

    /**
     * true if it's a directory entry
     */
    isDirectory: boolean;

    /**
     * true if it's a file entry, see also isDirectory
     */
    isFile: boolean;

    /**
     * file comment
     */
    comment: string;

    /**
     * if the file is encrypted
     */
    encrypted: boolean;

    /**
     * version made by
     */
    verMade: number;

    /**
     * version needed to extract
     */
    version: number;

    /**
     * encrypt, decrypt flags
     */
    flags: number;

    /**
     * compression method
     */
    method: number;

    /**
     * modification time
     */
    time: number;

    /**
     * uncompressed file crc-32 value
     */
    crc: number;

    /**
     * compressed size
     */
    compressedSize: number;

    /**
     * uncompressed size
     */
    size: number;

    /**
     * volume number start
     */
    diskStart: number;

    /**
     * internal file attributes
     */
    inattr: number;

    /**
     * external file attributes
     */
    attr: number;

    /**
     * LOC header offset
     */
    offset: number;
}

declare class BlobZipStream extends EventEmitter {
    constructor(config: StreamZipOptions);

    /**
     * number of entries in the archive
     */
    entriesCount: number;

    /**
     * archive comment
     */
    comment: string;

    readonly ready: boolean;

    on(event: 'error', handler: (error: any) => void): this;
    on(event: 'entry', handler: (entry: ZipEntry) => void): this;
    on(event: 'ready', handler: () => void): this;

    entry(name: string): ZipEntry | undefined;

    entries(): { [name: string]: ZipEntry };

    stream(entry: string, callback: (err: any | null, stream?: NodeJS.ReadableStream) => void): void;

    openEntry(entry: string, callback: (err: any | null, entry?: ZipEntry) => void): void;

    close(callback?: (err?: any) => void): void;
}

export = BlobZipStream;
