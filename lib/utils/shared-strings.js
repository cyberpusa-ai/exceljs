class SharedStrings {
  constructor() {
    this._values = [];
    this._totalRefs = 0;
    this._hash = Object.create(null);
    this._rich = Object.create(null);
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
    // richText objects are keyed by their serialized runs in a separate map,
    // so they neither collapse into '[object Object]' nor collide with plain strings
    const isRich = value !== null && typeof value === 'object';
    const hash = isRich ? this._rich : this._hash;
    const key = isRich ? JSON.stringify(value) : value;
    let index = hash[key];
    if (index === undefined) {
      index = hash[key] = this._values.length;
      this._values.push(value);
    }
    this._totalRefs++;
    return index;
  }
}

module.exports = SharedStrings;
