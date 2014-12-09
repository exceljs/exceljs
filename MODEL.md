# Workbook Model

The workbook and its components define a simple JavaScript Object model which can be accessed and manipulated if desired.

## Workbook Model

```javascript
{
    worksheets: [
        // array of worksheet models
    ]
}
```

## Worksheet Model

```javascript
{
    // worksheet id (integer>=1)
    id: 1,
    
    // worksheet name
    name: "blort",
    
    // rows
    rows: [
        // array of row models
    ],
    
    // merge ranges
    "merges": [
        "A2:B3"
    ],
    "dimensions": {
        "top": 1,
        "left": 1,
        "bottom": 3,
        "right": 6
    }    
}
```

## Row Model

```javascript
{
    // row number
    number: 1,
    cells: [
        // array of cell models
    ],
    
    // min column number
    min: 1,
    
    // maximum column number
    max: 6
}
```

## Cell Models

### Null Cell Model

```javascript
{
    address: "A1",
    type: 0
}
```

### Merge Cell Model

```javascript
{
    address: "B1",
    type: 1,
    master: "A1"
}
```

### Number Cell Model

```javascript
{
    address: "B1",
    type: 2,
    value: 5
}
```

### String Cell Model

```javascript
{
    address: "B1",
    type: 2,
    value: "Hello, World!"
}
```

### Date Cell Model

```javascript
{
    address: "C1",
    type: 3,
    value: new Date()
}
```

### Formula Cell Model

```javascript
{
    address: "D1",
    type: 4,
    formula: "A1+A2",
    result: 7
}
```

### Hyperlink Cell Model

```javascript
{
    "address": "F1",
    "type": 5,
    "text": "www.link.com",
    "hyperlink": "http://www.link.com"
}
```