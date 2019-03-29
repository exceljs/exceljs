var Anchor = require('../../../../../../lib/doc/anchor');

module.exports = {
  "anchors":[
    {
      "range": "A1:C7",
      "picture": {
        "rId": "rId1"
      }
    },
    {
      "range": {
        "tl": new Anchor({"row": 2.5, "col": 5.5}),
        "br": new Anchor({"row": 10.5, "col": 8.5}),
      },
      "picture": {
        "rId": "rId2"
      }
    }
  ]
};