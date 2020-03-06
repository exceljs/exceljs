const _ = require('../utils/under-dash');

class Note {
  constructor(note) {
    this.note = this.normalizeNote(note);
  }

  get model() {
    return this.note;
  }

  set model(value) {
    this.note = value;
  }

  normalizeNote(value) {
    switch (typeof value) {
      case 'string':
        value = {
          type: 'note',
          note: {
            texts: [
              {
                text: value,
              },
            ],
          },
        };
        break;
      default:
        value = {
          type: 'note',
          note: value,
        };
        break;
    }
    return _.deepMerge({}, Note.DEFAULT_CONFIGS, value);
  }

  static fromModel(model) {
    const note = new Note(model.note);
    note.model = model;
    return note;
  }
}

Note.DEFAULT_CONFIGS = {
  note: {
    insetmode: 'auto',
    inset: [0.13, 0.13, 0.25, 0.25],
  },
};

module.exports = Note;
