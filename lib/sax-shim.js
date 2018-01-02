var Sax = require("sax");

const oldCreateStream = Sax.createStream;
Sax.createStream = function(){
    const stream = oldCreateStream.call(this, true, {xmlns: true});

    const oldOn = stream.on;
    stream.on = function(evt, cb){
        oldOn.call(this, evt, function(node){
            if(evt === 'opentag'){
                const normalizedTag = {};
                normalizedTag.name = (node.ns && node.ns.local) || node.name;
                const normalizedAttributes = {};
                for(const [attribName, attrib] of Object.entries(node.attributes)){
                    normalizedAttributes[attribName] = attrib.value;
                }
                normalizedTag.attributes = normalizedAttributes;
                cb(normalizedTag);
            }
            else { cb(node); }
        });
    };
    return stream;
};


module.exports = Sax;