'use strict';

var expect = require('chai').expect;

var utils = require("../../../lib/utils/utils");

describe("utils", function() {
  it("encodes xml text", function() {

    expect(utils.xmlEncode('<')).to.equal('&lt;');
    expect(utils.xmlEncode('>')).to.equal('&gt;');
    expect(utils.xmlEncode('&')).to.equal('&amp;');
    expect(utils.xmlEncode('"')).to.equal('&quot;');
    expect(utils.xmlEncode("'")).to.equal('&apos;');

    expect(utils.xmlEncode('<a href="www.whatever.com">Talk to the H&</a>')).to.equal('&lt;a href=&quot;www.whatever.com&quot;&gt;Talk to the H&amp;&lt;/a&gt;');
  });
});