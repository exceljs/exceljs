var expect = require('chai').expect;

var SharedStrings = require('../../../lib/utils/shared-strings');

describe('SharedStrings', function() {
  it('Stores and shares string values', function() {
    var ss = new SharedStrings();

    var iHello = ss.add('Hello');
    var iHelloV2 = ss.add('Hello');
    var iGoodbye = ss.add('Goodbye');

    expect(iHello).to.equal(iHelloV2);
    expect(iGoodbye).to.not.equal(iHelloV2);

    expect(ss.count).to.equal(2);
    expect(ss.totalRefs).to.equal(3);
  });

  it('Does not escape values', function() {
    // that's the job of the xml utils
    var ss = new SharedStrings();

    var iXml = ss.add('<tag>value</tag>');
    var iAmpersand = ss.add('&');

    expect(ss.getString(iXml)).to.equal('<tag>value</tag>');
    expect(ss.getString(iAmpersand)).to.equal('&');
  });
});
