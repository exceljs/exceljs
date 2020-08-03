module.exports = function(value) {
  if (value.name) {
    // Remove prefixes
    value.name = value.name.replace(/^\w+:/, '');
  }
  return value;
};
