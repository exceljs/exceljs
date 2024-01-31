

const textDecoder = new TextDecoder('utf-8');
const { StringDecoder } = require('string_decoder');
const decoder = new StringDecoder('utf8');
const testStringBuffer = Buffer.from('凡是能够说的，都能够说清楚；凡是不能说的，必须保持沉默');

const bufferArr = [];
let start = 0;
const chunkSize = 16;
while(start<=testStringBuffer.length){
    let end = start+chunkSize;
    end =  end > testStringBuffer.length?testStringBuffer.length:end;
    bufferArr.push(testStringBuffer.subarray(start,end))
    start = start+chunkSize
}

for(let chunk of bufferArr){ // StringDecoder can handle multibyte charcter properly
    console.log(decoder.write(chunk));
}

for(let chunk of bufferArr){ // textDecoder will get � because of the multibyte charcter was incomplete in each buffer chunk
    console.log(textDecoder.decode(chunk));
}

