'use strict';

const crypto = require('crypto');

const Encryptor = {
  /**
   * Calculate a hash of the concatenated buffers with the given algorithm.
   * @param {string} algorithm - The hash algorithm.
   * @returns {Buffer} The hash
   */
  hash(algorithm, ...buffers) {
    const hashes = crypto.getHashes();
    if (hashes.indexOf(algorithm) < 0) {
      throw new Error(`Hash algorithm '${  algorithm  }' not supported!`);
    }
    const hash = crypto.createHash(algorithm);
    hash.update(Buffer.concat(buffers));
    return hash.digest();
  },
  /**
   * Convert a password into an encryption key
   * @param {string} password - The password
   * @param {string} hashAlgorithm - The hash algoritm
   * @param {string} saltValue - The salt value
   * @param {number} spinCount - The spin count
   * @param {number} keyBits - The length of the key in bits
   * @param {Buffer} blockKey - The block key
   * @returns {Buffer} The encryption key
   */
  convertPasswordToHash(password, hashAlgorithm, saltValue, spinCount) {
    hashAlgorithm = hashAlgorithm.toLowerCase();
    // Password must be in unicode buffer
    const passwordBuffer = Buffer.from(password, 'utf16le');
    // Generate the initial hash
    let key = this.hash(
      hashAlgorithm,
      Buffer.from(saltValue, 'base64'),
      passwordBuffer
    );
    // Now regenerate until spin count
    for (let i = 0; i < spinCount; i++) {
      const iterator = Buffer.alloc(4);
      iterator.writeUInt32LE(i, 0);
      key = this.hash(hashAlgorithm, key, iterator);
    }
    return key.toString('base64');
  },
  /**
   * Generates cryptographically strong pseudo-random data.
   * @param size The size argument is a number indicating the number of bytes to generate.
   */
  randomBytes(size) {
    return crypto.randomBytes(size);
  },
};
module.exports = Encryptor;
