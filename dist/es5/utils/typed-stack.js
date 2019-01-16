/**
 * Copyright (c) 2015-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var TypedStack = module.exports = function (type) {
  this._type = type;
  this._stack = [];
};

TypedStack.prototype = {
  get size() {
    return this._stack.length;
  },

  pop: function pop() {
    var tos = this._stack.pop();
    return tos || new this._type();
  },

  push: function push(instance) {
    if (!(instance instanceof this._type)) {
      throw new Error('Invalid type pushed to TypedStack');
    }
    this._stack.push(instance);
  }
};
//# sourceMappingURL=typed-stack.js.map
