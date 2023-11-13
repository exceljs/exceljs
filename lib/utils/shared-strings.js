class SharedStrings {
  constructor() {
    this._values = [];
    this._totalRefs = 0;
    this._hash = Object.create(null);
  }

  get count() {
    return this._values.length;
  }

  get values() {
    return this._values;
  }

  get totalRefs() {
    return this._totalRefs;
  }

  getString(index) {
    return this._values[index];
  }

  add(value) {
    const key = this.getKeyFromValue(value);
    let index = this._hash[key];

    if (index === undefined) {
      index = this._hash[key] = this._values.length;
      this._values.push(value);
    }

    this._totalRefs++;
    return index;
  }

  getKeyFromValue(value) {
    if (typeof value !== 'object' || value === null) {
      return value;
    }

    // Convert richText value to string
    return JSON.stringify(value);
  }
}

module.exports = SharedStrings;
