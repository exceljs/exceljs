'use strict';

var Promise = require('es6-promise').Promise;

function isErrorClass(type) {
  while (type && (type !== Object)) {
    if ((type === Error) || (type instanceof Error)) {
      return true;
    }
    type = type.prototype;
  }
  return false;
}

class Promish extends Promise {
  constructor(f) {
    if (f instanceof Promish) {
      return f;
    } else if ((f instanceof Promise) || (f.then instanceof Function)) {
      super((resolve, reject) => f.then(resolve, reject));
    } else if (f instanceof Error) {
      // sugar for 'rethrow'
      super((resolve, reject) => reject(f));
    } else if (f instanceof Function) {
      super(f);
    } else {
      // anything else, resolve with value
      super(resolve => resolve(f));
    }
  }

  finally(h) {
    return this.then(
        value => Promish.resolve(h()).then(() => value),
    error => Promish.resolve(h()).then(() => Promish.reject(error))
  );
  }

  catch() {
    // extend catch with type-aware or matcher handling
    var args = Array.from(arguments);
    var h = args.pop();
    return this.then(undefined, function(error) {
      // default catch - no matchers. Just return handler result
      if (!args.length) {
        return h(error);
      }

      //console.log('catch matcher', error)
      // search for a match in argument order and return handler result if found
      for (var i = 0; i < args.length; i++) {
        var matcher = args[i];
        if (isErrorClass(matcher)) {
          if (error instanceof matcher) {
            return h(error);
          }
        } else if (matcher instanceof Function) {
          //console.log('matcher function')
          if (matcher(error)) {
            //console.log('matched!!')
            return h(error);
          }
        }
      }

      // no match was found send this error to the next promise handler in the chain
      return new Promish((resolve, reject) => reject(error));
    });
  }

  delay(timeout) {
    return this.then(function(value) {
      return new Promish(function(resolve) {
        setTimeout(function() {
          resolve(value);
        }, timeout);
      });
    });
  }

  spread(f) {
    return this.then(function(values) {
      return Promish.all(values);
    }).then(function(values) {
      return f.apply(undefined, values);
    });
  }

  map(f) {
    return this.then(function(values) {
      return Promish.map(values, f);
    });
  }

  static map(values, f) {
    return Promish.all(
      values.map((v,i) => f(v, i, values.length))
    )
  }

  // Wrap a synchronous method and resolve with its return value
  static method(f) {
    return function() {
      var self = this; // is this necessary?
      var args = Array.from(arguments);
      return new Promish(resolve => resolve(f.apply(self, args)));
    };
  }

  //
  static apply(f, args) {
    // take a copy of args because a) might not be Array and b) no side-effects
    args = Array.from(args);
    return new Promish(function(resolve, reject) {
      args.push(function () {
        var error = Array.prototype.shift.apply(arguments);
        if (error) {
          reject(error);
        } else {
          if (arguments.length === 1) {
            resolve(arguments[0]);
          } else {
            resolve(arguments);
          }
        }
      });
      f.apply(undefined, args);
    });
  }
  static nfapply(f, args) {
    return Promish.apply(f, args);
  }

  static call() {
    var f = Array.prototype.shift.apply(arguments);
    return Promish.apply(f, arguments);
  }
  static nfcall() {
    return Promish.call.apply(null, arguments);
  }

  static post(o, f, a) {
    return Promish.apply(f.bind(o), a);
  }
  static npost(o,f,a) {
    return Promish.apply(f.bind(o), a);
  }

  static invoke() {
    var o = Array.prototype.shift.apply(arguments);
    var f = Array.prototype.shift.apply(arguments);
    return Promish.apply(f.bind(o), arguments);
  }
  static ninvoke() {
    return Promish.invoke(arguments);
  }

  // create curry function for nfcall
  static promisify(f) {
    return function() {
      return Promish.apply(f, arguments);
    };
  }
  static denodify(f) {
    return Promish.promisify(f);
  }

  // create Q based curry function for ninvoke
  static nbind(f, o) {
    // Why is it function, object and not object, function like the others?
    return function() {
      return Promish.post(o, f, arguments);
    };
  }

  // curry function for ninvoke with arguments in object, method order
  static bind(o, f) {
    return function() {
      return Promish.post(o, f, arguments);
    };
  }

  // Promishify every method in an object
  static promisifyAll(o, options) {
    options = options || {};
    var inPlace = options.inPlace || false;
    var suffix = options.suffix || (inPlace ? 'Async' : '');

    var p = {};
    var oo = o;
    while (oo && (oo !== Object)) {
      for (let i in oo) {
        if (!p[i + suffix] && (oo[i] instanceof Function)) {
          p[i + suffix] = Promish.bind(o, oo[i]);
        }
      }
      oo = Object.getPrototypeOf(oo) || oo.prototype;
    }

    if (inPlace) {
      for (let i in p) {
        if(p[i] instanceof Function) {
          o[i] = p[i];
        }
      }
      p = o;
    }

    return p;
  }

  // some - the first n to resolve, win - else reject with all of the errors
  static some(promises, n) {
    return new Promish(function(resolve, reject) {
      var values = [];
      var rejects = [];
      promises.forEach(function(promise) {
        promise
          .then(function(value) {
            values.push(value);
            if (values.length >= n) {
              resolve(values);
            }
          })
          .catch(function(error) {
            rejects.push(error);
            if (rejects.length > promises.length - n){
              reject(rejects);
            }
          });
      });
    });
  }

  // any - the first to resolve, wins - else reject with all of the errors
  static any(promises) {
    return Promish.some(promises, 1)
      .then(function(values) {
        return values[0];
      });
  }

  // old-style for ease of adoption
  static defer() {
    var deferred = {};
    deferred.promise = new Promish(function(resolve, reject) {
      deferred.resolve = resolve;
      deferred.reject = reject;
    });
    return deferred;
  }

  // spread - apply array of values to function as args
  static spread(value, f) {
    return f.apply(undefined, value);
  }
}

module.exports = {
  Promish
};
