const { expect } = require('chai');

const Encryptor = require('../../../lib/utils/encryptor');

describe('Encryptor', () => {
  it('Generates SHA-512 hash for given password, salt value and spin count', () => {
    const password = '123';
    const saltValue = '6tC6yotbNa8JaMaDvbUgxw==';
    const spinCount = 100000;
    const hash = Encryptor.convertPasswordToHash(password, 'SHA512', saltValue, spinCount);
    expect(hash).to.equal('RHtx1KpAYT7nBzGCTInkHrbf2wTZxP3BT4Eo8PBHPTM4KfKArJTluFvizDvo6GnBCOO6JJu7qwKvMqnKHs7dcw==');
  });

});
