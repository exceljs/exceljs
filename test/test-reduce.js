var a = [
  { name: 'tom', age: 7 },
  { name: 'dick', age: 8 },
  { name: 'harry', age: 9 },
];

var index = a.reduce((o,v) => ((o[v.name] = v), o), {});
console.log(index)