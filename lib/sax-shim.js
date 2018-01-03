var Sax = require("sax");

const oldCreateStream = Sax.createStream;
Sax.createStream = function(strict, opts){
    const stream = oldCreateStream.call(this, strict, opts);

    const oldOn = stream.on;
    stream.on = function(evt, cb){
        oldOn.call(this, evt, function(node) {
            if (evt === 'opentag') {
                const normalizedTag = {};
                normalizedTag.name = removeNamespace(node.name);
                const normalizedAttributes = {};
                for (const [attribName, attrib] of Object.entries(node.attributes)) {
                    normalizedAttributes[removeNamespace(attribName)] = attrib;
                }
                normalizedTag.attributes = normalizedAttributes;
                cb(normalizedTag);
            }
            else if (evt === "closetag") {
                cb(removeNamespace(node));
            } else {
                cb(node);
            }
        });
    };
    return stream;
};
function removeNamespace(tagName){
    const split = tagName.split(":");
    if(split.length === 1){ return split[0]; }
    else { return split.splice(1).join(":"); }
}


module.exports = Sax;