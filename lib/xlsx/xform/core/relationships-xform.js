/**
 * Copyright (c) 2016 Guyon Roche
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 *
 */
'use strict';

var utils = require('../../../utils/utils');
var XmlStream = require('../../../utils/xml-stream');
var BaseXform = require('./../base-xform');

// <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
// <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
//   <%_.each(relationships, function(r) {%>
//     <Relationship
//       Id="<%=r.rId%>"
//       Type="<%=r.type%>"
//       Target="<%=r.target%>" <% if (r.targetMode) {%>
//       TargetMode="<%=r.targetMode%>"<%}%>
//     />
//     <%});%>
// </Relationships>
  
var RelationshipsXform = module.exports = function() {
  this.map = {
  }
};

utils.inherits(RelationshipsXform, BaseXform, {
  
},{
  write: function(xmlStream, model) {
    model = model || this._values;
    xmlStream.openXml(XmlStream.StdDocAttributes);

  },

  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else {
      switch(node.name) {
        case 'cp:coreProperties':
          return true;
        default:
          this.parser = this.map[node.name];
          if (this.parser) {
            this.parser.parseOpen(node);
            return true;
          }
          throw new Error('Unexpected xml node in parseOpen: ' + JSON.stringify(node));
      }
    }
  },
  parseText: function(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  },
  parseClose: function(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    } else {
      switch(name) {
        case 'cp:coreProperties':
          this.model = {
            creator: this.map['dc:creator'].model,
            lastModifiedBy: this.map['cp:lastModifiedBy'].model,
            created: this.map['dcterms:created'].model,
            modified: this.map['dcterms:modified'].model
          };
          return false;
        default:
          throw new Error('Unexpected xml node in parseClose: ' + name);
      }
    }
  }
});
