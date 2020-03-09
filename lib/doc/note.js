const _ = require('../utils/under-dash');

class Note {
  constructor(note) {
    note = this.normalizeNode(note);

    // Suitable for all cell comments
    this.note = _.deepMerge(
      {},
      {
        margins: {
          insetmode: 'auto',
          inset: [0.13, 0.13, 0.25, 0.25],
        },
        protection: {
          locked: 'True',
          lockText: 'True',
        },
        editAs: 'absolute',
      },
      note.note
    );
  }

  normalizeNode(note) {
    if (typeof note === 'string') {
      note = {
        texts: [
          {
            text: note,
          },
        ],
      };
    }
    return {note};
  }

  static fromModel(model) {
    const note = new Note(model.note);
    note.model = model;
    return note;
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
