const a = [
  { name: 'tom', age: 7 },
  { name: 'dick', age: 8 },
  { name: 'harry', age: 9 },
];

const index = a.reduce((o, v) => ((o[v.name] = v), o), {});
console.log(index);
