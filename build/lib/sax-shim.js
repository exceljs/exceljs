"use strict";

var _slicedToArray = function () { function sliceIterator(arr, i) { var _arr = []; var _n = true; var _d = false; var _e = undefined; try { for (var _i = arr[Symbol.iterator](), _s; !(_n = (_s = _i.next()).done); _n = true) { _arr.push(_s.value); if (i && _arr.length === i) break; } } catch (err) { _d = true; _e = err; } finally { try { if (!_n && _i["return"]) _i["return"](); } finally { if (_d) throw _e; } } return _arr; } return function (arr, i) { if (Array.isArray(arr)) { return arr; } else if (Symbol.iterator in Object(arr)) { return sliceIterator(arr, i); } else { throw new TypeError("Invalid attempt to destructure non-iterable instance"); } }; }();

var Sax = require("sax");

var oldCreateStream = Sax.createStream;
Sax.createStream = function (strict, opts) {
    var stream = oldCreateStream.call(this, strict, opts);

    var oldOn = stream.on;
    stream.on = function (evt, cb) {
        oldOn.call(this, evt, function (node) {
            if (evt === 'opentag') {
                var normalizedTag = {};
                normalizedTag.name = removeNamespace(node.name);
                var normalizedAttributes = {};
                var _iteratorNormalCompletion = true;
                var _didIteratorError = false;
                var _iteratorError = undefined;

                try {
                    for (var _iterator = Object.entries(node.attributes)[Symbol.iterator](), _step; !(_iteratorNormalCompletion = (_step = _iterator.next()).done); _iteratorNormalCompletion = true) {
                        var _step$value = _slicedToArray(_step.value, 2),
                            attribName = _step$value[0],
                            attrib = _step$value[1];

                        normalizedAttributes[removeNamespace(attribName)] = attrib;
                    }
                } catch (err) {
                    _didIteratorError = true;
                    _iteratorError = err;
                } finally {
                    try {
                        if (!_iteratorNormalCompletion && _iterator.return) {
                            _iterator.return();
                        }
                    } finally {
                        if (_didIteratorError) {
                            throw _iteratorError;
                        }
                    }
                }

                normalizedTag.attributes = normalizedAttributes;
                cb(normalizedTag);
            } else if (evt === "closetag") {
                cb(removeNamespace(node));
            } else {
                cb(node);
            }
        });
    };
    return stream;
};
function removeNamespace(tagName) {
    var split = tagName.split(":");
    if (split.length === 1) {
        return split[0];
    } else {
        return split.splice(1).join(":");
    }
}

module.exports = Sax;
//# sourceMappingURL=sax-shim.js.map
