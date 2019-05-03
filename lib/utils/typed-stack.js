'use strict';

var TypedStack = module.exports = function(type) {
  this._type = type;
  this._stack = [];
};

TypedStack.prototype = {
  get size() {
    return this._stack.length;
  },

  pop: function() {
    var tos = this._stack.pop();
    return tos || new (this._type)();
  },

  push: function(instance) {
    if (!(instance instanceof this._type)) {
      throw new Error('Invalid type pushed to TypedStack');
    }
    this._stack.push(instance);
  }
};
