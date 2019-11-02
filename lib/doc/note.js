class Note {
  constructor(note) {
    this.note = note;
  }

  get model() {
    switch (typeof this.note) {
      case 'string':
        return {
          type: 'note',
          note: {
            texts: [{text: this.note}],
          },
        };

      default:
        return {
          type: 'note',
          note: this.note,
        };
    }
  }

  set model(value) {
    // convenience - simplify unstyled notes
    const {note} = value;
    const {texts} = note;
    if ((texts.length === 1) && (Object.keys(texts[0]).length === 1)) {
      this.note = texts[0].text;
    } else {
      this.note = note;
    }
  }

  static fromModel(model) {
    const note = new Note();
    note.model = model;
    return note;
  }
}

module.exports = Note;
