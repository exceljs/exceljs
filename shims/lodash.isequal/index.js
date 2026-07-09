const { isDeepStrictEqual } = require("util");

module.exports = function isEqual(a, b) {
  return isDeepStrictEqual(a, b);
};
