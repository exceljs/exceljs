const crypto = require('crypto');

function uuidv4() {
  const bytes = crypto.randomBytes(16);
  bytes[6] = (bytes[6] % 16) + 0x40;
  bytes[8] = (bytes[8] % 64) + 0x80;
  const hex = bytes.toString('hex');
  return `${hex.substr(0, 8)}-${hex.substr(8, 4)}-${hex.substr(12, 4)}-${hex.substr(
    16,
    4
  )}-${hex.substr(20, 12)}`;
}

module.exports = uuidv4;
