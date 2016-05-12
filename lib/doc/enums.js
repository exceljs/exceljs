/**
 * Copyright (c) 2014 Guyon Roche
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
"use strict";

module.exports = {
    ValueType: {
        Null: 0,
        Merge: 1,
        Number: 2,
        String: 3,
        Date: 4,
        Hyperlink: 5,
        Formula: 6,
        SharedString: 7,
        RichText: 8
    },
    RelationshipType: {
        None: 0,
        OfficeDocument: 1,
        Worksheet: 2,
        CalcChain: 3,
        SharedStrings: 4,
        Styles: 5,
        Theme: 6,
        Hyperlink: 7
    },
    DocumentType: {
        Xlsx: 1
    },
    ReadingOrder: {
        RightToLeft: 1,
        LeftToRight: 2
    }
}