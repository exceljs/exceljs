var Excel = require("../excel");

describe("Worksheet", function() {
    describe("Values", function() {
        it("stores values properly", function() {
            var wb = new Excel.Workbook()
            var ws = wb.addWorksheet("blort");
            
            var now = new Date();
            
            // plain number
            ws.getCell("A1").value = 7;
            
            // simple string
            ws.getCell("B1").value = "Hello, World!";
            
            // floating point
            ws.getCell("C1").value = 3.14;
            
            // 5 will be overwritten by the current date-time
            ws.getCell("D1").value = 5;
            ws.getCell("D1").value = now;
            
            // constructed string - will share recored with B1
            ws.getCell("E1").value = ["Hello", "World"].join(", ") + "!";
            
            // hyperlink
            ws.getCell("F1").value = {text: "www.google.com", hyperlink:"http://www.google.com"};
            
            // number formula
            ws.getCell("A2").value = {formula: "A1", result: 7};
            
            // string formula
            ws.getCell("B2").value = {formula: 'CONCATENATE("Hello", ", ", "World!")', result: "Hello, World!"};
            
            // date formula
            ws.getCell("C2").value = {formula: "D1", result: now};
            
            expect(ws.getCell("A1").value).toEqual(7);
            expect(ws.getCell("B1").value).toEqual("Hello, World!");
            expect(ws.getCell("C1").value).toEqual(3.14);
            expect(ws.getCell("D1").value).toEqual(now);
            expect(ws.getCell("E1").value).toEqual("Hello, World!");
            expect(ws.getCell("F1").value.text).toEqual("www.google.com");
            expect(ws.getCell("F1").value.hyperlink).toEqual("http://www.google.com");
            
            expect(ws.getCell("A2").value.formula).toEqual("A1");
            expect(ws.getCell("A2").value.result).toEqual(7);
            
            expect(ws.getCell("B2").value.formula).toEqual('CONCATENATE("Hello", ", ", "World!")');
            expect(ws.getCell("B2").value.result).toEqual("Hello, World!");
            
            expect(ws.getCell("C2").value.formula).toEqual("D1");
            expect(ws.getCell("C2").value.result).toEqual(now);
        });
        
        it("stores shared string values properly", function() {
            var wb = new Excel.Workbook()
            var ws = wb.addWorksheet("blort");
            
            ws.getCell("A1").value = "Hello, World!";
            
            ws.getCell("A2").value = "Hello";
            ws.getCell("B2").value = "World";
            ws.getCell("C2").value = {formula: 'CONCATENATE(A2, ", ", B2, "!")', result: "Hello, World!"};
            
            ws.getCell("A3").value = ["Hello", "World"].join(", ") + "!";
            
            // A1 and A3 should reference the same string object
            expect(ws.getCell("A1").value).toBe(ws.getCell("A3").value);
            
            // A1 and C2 should not reference the same object
            expect(ws.getCell("A1").value).toBe(ws.getCell("C2").value.result);
        });
        
        it("assigns cell types properly", function() {
            var wb = new Excel.Workbook()
            var ws = wb.addWorksheet("blort");
            
            // plain number
            ws.getCell("A1").value = 7;
            
            // simple string
            ws.getCell("B1").value = "Hello, World!";
            
            // floating point
            ws.getCell("C1").value = 3.14;
            
            // date-time
            ws.getCell("D1").value = new Date();
            
            // hyperlink
            ws.getCell("E1").value = {text: "www.google.com", hyperlink:"http://www.google.com"};
            
            // number formula
            ws.getCell("A2").value = {formula: "A1", result: 7};
            
            // string formula
            ws.getCell("B2").value = {formula: 'CONCATENATE("Hello", ", ", "World!")', result: "Hello, World!"};
            
            // date formula
            ws.getCell("C2").value = {formula: "D1", result: new Date()};
            
            expect(ws.getCell("A1").type).toEqual(Excel.ValueType.Number);
            expect(ws.getCell("B1").type).toEqual(Excel.ValueType.String);
            expect(ws.getCell("C1").type).toEqual(Excel.ValueType.Number);
            expect(ws.getCell("D1").type).toEqual(Excel.ValueType.Date);
            expect(ws.getCell("E1").type).toEqual(Excel.ValueType.Hyperlink);

            expect(ws.getCell("A2").type).toEqual(Excel.ValueType.Formula);
            expect(ws.getCell("B2").type).toEqual(Excel.ValueType.Formula);
            expect(ws.getCell("C2").type).toEqual(Excel.ValueType.Formula);
        });
        
        it("adds column headers", function() {
            var wb = new Excel.Workbook()
            var ws = wb.addWorksheet("blort");
            
            ws.columns = [
                { header: "Id", width: 10 },
                { header: "Name", width: 32 },
                { header: "D.O.B.", width: 10 }
            ];
            
            expect(ws.getCell("A1").value).toEqual("Id");
            expect(ws.getCell("B1").value).toEqual("Name");
            expect(ws.getCell("C1").value).toEqual("D.O.B.");
        });
        it("adds rows by object", function() {
            var wb = new Excel.Workbook()
            var ws = wb.addWorksheet("blort");
            
            // add columns to define column keys
            ws.columns = [
                { header: "Id", key: "id", width: 10 },
                { header: "Name", key: "name", width: 32 },
                { header: "D.O.B.", key: "dob", width: 10 }
            ];
            
            ws.addRow({id:1, name: "John Doe", dob: new Date(1970,1,1)});
            ws.addRow({id:2, name: "Jane Doe", dob: new Date(1965,1,7)});
            
            expect(ws.getCell("A2").value).toEqual(1);
            expect(ws.getCell("B2").value).toEqual("John Doe");
            expect(ws.getCell("C2").value).toEqual(new Date(1970,1,1));

            expect(ws.getCell("A3").value).toEqual(2);
            expect(ws.getCell("B3").value).toEqual("Jane Doe");
            expect(ws.getCell("C3").value).toEqual(new Date(1965,1,7));
        });
        
        it("adds rows by array", function() {
            var wb = new Excel.Workbook()
            var ws = wb.addWorksheet("blort");
            
            ws.addRow([1, "John Doe", new Date(1970,1,1)]);
            ws.addRow([2, "Jane Doe", new Date(1965,1,7)]);
            
            expect(ws.getCell("A1").value).toEqual(1);
            expect(ws.getCell("B1").value).toEqual("John Doe");
            expect(ws.getCell("C1").value).toEqual(new Date(1970,1,1));

            expect(ws.getCell("A2").value).toEqual(2);
            expect(ws.getCell("B2").value).toEqual("Jane Doe");
            expect(ws.getCell("C2").value).toEqual(new Date(1965,1,7));
        });
    });
    
    describe("Merge Cells", function() {
        it("references the same top-left value", function() {
            var wb = new Excel.Workbook()
            var ws = wb.addWorksheet("blort");
            
            // initial values
            ws.getCell("A1").value = "A1";
            ws.getCell("B1").value = "B1";
            ws.getCell("A2").value = "A2";
            ws.getCell("B2").value = "B2";
            
            ws.mergeCells("A1:B2");
            
            expect(ws.getCell("A1").value).toEqual("A1");
            expect(ws.getCell("B1").value).toEqual("A1");
            expect(ws.getCell("A2").value).toEqual("A1");
            expect(ws.getCell("B2").value).toEqual("A1");
            
            expect(ws.getCell("A1").type).toEqual(Excel.ValueType.String);
            expect(ws.getCell("B1").type).toEqual(Excel.ValueType.Merge);
            expect(ws.getCell("A2").type).toEqual(Excel.ValueType.Merge);
            expect(ws.getCell("B2").type).toEqual(Excel.ValueType.Merge);
        });
        
        it("does not allow overlapping merges", function() {
            var wb = new Excel.Workbook()
            var ws = wb.addWorksheet("blort");
            
            ws.mergeCells("B2:C3");
            
            // intersect four corners
            expect(function() { ws.mergeCells("A1:B2"); }).toThrow();
            expect(function() { ws.mergeCells("C1:D2"); }).toThrow();
            expect(function() { ws.mergeCells("C3:D4"); }).toThrow();
            expect(function() { ws.mergeCells("A3:B4"); }).toThrow();
            
            // enclosing
            expect(function() { ws.mergeCells("A1:D4"); }).toThrow();
        });
        
        it("puts the lotion in the basket", function() {
            var wb = new Excel.Workbook()
            var ws = wb.addWorksheet("basket");
            ws.getCell("A1").value = "lotion";
        });
    });
});
