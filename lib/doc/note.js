module.exports = class Note {
  constructor(note) {
    this.note = note;
  }

  get model() {
    return {
      type: 'note',
      note: this.note,
    };
  }

  set model(value) {
    this.note = value.note;
  }
};
