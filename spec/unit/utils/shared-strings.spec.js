var expect = require('chai').expect;

var SharedStrings = require('../../../lib/utils/shared-strings');

describe('SharedStrings', function() {

  it('Stores and shares string values', function() {
    var ss = new SharedStrings();

    var sHello = ss.add('Hello');
    var sHello_v2 = ss.add('Hello');
    var sGoodbye = ss.add('Goodbye');

    expect(sHello).to.equal(sHello_v2);
    expect(sGoodbye).to.not.equal(sHello_v2);

    expect(ss.count).to.equal(2);
    expect(ss.totalRefs).to.equal(3);
  });
});
