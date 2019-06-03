class Note {
  constructor(note) {
    if (typeof note === 'string') {
      this.note = {
        texts: [
          {text: note},
        ],
      };
    } else {
      this.note = note;
    }
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
}

module.exports = Note;
