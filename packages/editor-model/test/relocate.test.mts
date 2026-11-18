// SPDX-License-Identifier: GPL-3.0-or-later
//
// A place found again in another version of a document (ED-011 section 3.5): where a caret lands when its text is
// there, edited, moved or gone, over lists of strings, and a place read from a document and put back.
import { describe, expect, test } from 'vitest';
import { EditorDocument, compareText, paragraphStrings, positionOfPlace, relocator, stateFor, textPlaceOf } from '../src/index.mjs';
import { allFixtures, fixture } from './helpers.mjs';

const CLAUSE = 'The members may remove a director by ordinary resolution.';
const OTHER = 'A quorum at any general meeting is five members present.';
const THIRD = 'Notice of a meeting is given fourteen clear days before it.';
const FOURTH = 'The financial year of the society ends on the last day of June.';

const land = (before: string[], after: string[], paragraph: number, offset = 0) => relocator(before, after)({ paragraph, offset });

describe('a text that is there', () => {
  test('stays with its paragraph, at its offset, when paragraphs come and go around it', () => {
    expect(land([OTHER, CLAUSE, THIRD], ['New.', 'Newer.', OTHER, CLAUSE, THIRD], 1, 9)).toEqual({ paragraph: 3, offset: 9, found: 'kept' });
    expect(land([OTHER, THIRD, CLAUSE], [CLAUSE], 2, 9)).toEqual({ paragraph: 0, offset: 9, found: 'kept' });
  });

  test('among repeated paragraphs is the one the alignment holds, not the first of its text', () => {
    expect(land(['x', 'Reserved', 'y', 'Reserved', 'z'], ['x', 'Reserved', 'y', 'Reserved', 'added', 'z'], 3)).toMatchObject({ paragraph: 3, found: 'kept' });
  });

  test('a blank paragraph is kept by its neighbours', () => {
    expect(land([OTHER, '', CLAUSE], ['Added.', OTHER, '', CLAUSE], 1)).toEqual({ paragraph: 2, offset: 0, found: 'kept' });
  });
});

describe('a paragraph edited', () => {
  const edited = `${CLAUSE.slice(0, 16)}at any time ${CLAUSE.slice(16)}`;

  test('keeps the caret in the text before the edit, and with the text after it', () => {
    expect(land([OTHER, CLAUSE, THIRD], [OTHER, edited, THIRD], 1, 4)).toEqual({ paragraph: 1, offset: 4, found: 'kept' });
    expect(land([OTHER, CLAUSE, THIRD], [OTHER, edited, THIRD], 1, 30)).toEqual({ paragraph: 1, offset: 30 + 'at any time '.length, found: 'kept' });
    expect(land([OTHER, CLAUSE, THIRD], [OTHER, edited, THIRD], 1, CLAUSE.length)).toEqual({ paragraph: 1, offset: edited.length, found: 'kept' });
  });

  test('puts a caret inside the replaced words at the replacement\'s start', () => {
    const replaced = CLAUSE.replace('remove a director', 'dismiss an officer');
    expect(land([OTHER, CLAUSE, THIRD], [OTHER, replaced, THIRD], 1, 20)).toEqual({ paragraph: 1, offset: CLAUSE.indexOf('remove'), found: 'kept' });
  });

  test('two edits far apart leave it the same paragraph', () => {
    const twice = CLAUSE.replace('The members', 'Those members').replace('ordinary resolution.', 'special resolution.');
    expect(land([OTHER, CLAUSE, THIRD], [OTHER, twice, THIRD], 1, 20).found).toBe('kept');
  });

  test('text added all round it leaves it the paragraph it is in', () => {
    const grown = `${OTHER} ${THIRD} ${CLAUSE} ${FOURTH}`;
    expect(land(['Title of it', CLAUSE], ['Title of it', grown], 1, 5).found).toBe('kept');
  });

  test('a text of a word or two stands if it is still in the paragraph', () => {
    expect(land([OTHER, 'Aims'], [OTHER, 'Aims and objects of the society'], 1, 2)).toEqual({ paragraph: 1, offset: 2, found: 'kept' });
  });

  test('in a document of one paragraph, with nothing else to go by', () => {
    expect(land([CLAUSE], [`${CLAUSE} More.`], 0, 7)).toEqual({ paragraph: 0, offset: 7, found: 'kept' });
  });

  test('never lands inside a character of two code units', () => {
    const a = 'A smile \u{1F600} and the rest of the sentence goes on for a while.';
    const b = 'A smile \u{1F601} and the rest of the sentence goes on for a while.';
    const { offset } = land([OTHER, a], [OTHER, b], 1, 9);
    expect(offset).toBe(8);
  });
});

describe('a paragraph moved', () => {
  test('is followed, the offset with it', () => {
    expect(land([CLAUSE, OTHER, THIRD, FOURTH], [OTHER, THIRD, CLAUSE, FOURTH], 0, 6)).toEqual({ paragraph: 2, offset: 6, found: 'moved' });
  });

  test('is not followed to a copy that had its own place', () => {
    expect(land([OTHER, CLAUSE, THIRD, CLAUSE], [OTHER, THIRD, CLAUSE], 1)).toMatchObject({ found: 'gone', paragraph: 1 });
  });

  test('is not followed when two paragraphs could be it', () => {
    // The four others outweigh either copy, so neither is matched.
    expect(land([OTHER, THIRD, CLAUSE, FOURTH, 'And the last of them.'], [CLAUSE, OTHER, THIRD, FOURTH, 'And the last of them.', CLAUSE], 2).found).toBe('gone');
  });

  test('a short paragraph is not followed: it repeats by chance', () => {
    expect(land(['Notes', OTHER, THIRD], [OTHER, THIRD, 'Notes'], 0).found).toBe('gone');
  });
});

describe('a text that is gone', () => {
  test('leaves the place where it stood: the paragraph after the gap', () => {
    expect(land([OTHER, CLAUSE, THIRD], [OTHER, THIRD], 1, 9)).toEqual({ paragraph: 1, offset: 0, found: 'gone' });
  });

  test('a block around it, the same', () => {
    expect(land([OTHER, 'One of them.', CLAUSE, 'Two of them.', THIRD], [OTHER, THIRD], 2)).toEqual({ paragraph: 1, offset: 0, found: 'gone' });
  });

  test('through to the end of the file: the end of what is left', () => {
    expect(land([OTHER, THIRD, CLAUSE, FOURTH], [OTHER, THIRD], 2, 4)).toEqual({ paragraph: 1, offset: THIRD.length, found: 'gone' });
  });

  test('from the start of the file: the first paragraph left', () => {
    expect(land([CLAUSE, OTHER, THIRD], [OTHER, THIRD], 0)).toEqual({ paragraph: 0, offset: 0, found: 'gone' });
  });

  test('replaced by unrelated text, which the alignment calls changed, is gone all the same', () => {
    expect(compareText([OTHER, CLAUSE, THIRD], [OTHER, FOURTH, THIRD]).changed).toBe(1);
    expect(land([OTHER, CLAUSE, THIRD], [OTHER, FOURTH, THIRD], 1, 9)).toEqual({ paragraph: 1, offset: 0, found: 'gone' });
  });

  test('a blank paragraph has no text to lose: it lands the same, and nothing is said', () => {
    expect(land([OTHER, '', THIRD], [OTHER, THIRD], 1)).toEqual({ paragraph: 1, offset: 0, found: 'kept' });
  });
});

describe('versions with nothing in common', () => {
  test('the place is the start, wherever the caret was', () => {
    expect(land([OTHER, CLAUSE, THIRD], ['Completely.', 'Different things here.', 'Nothing alike.', 'At all.'], 2, 5)).toEqual({ paragraph: 0, offset: 0, found: 'unrelated' });
  });

  test('blank paragraphs in common make nothing in common', () => {
    expect(land([OTHER, '', CLAUSE], ['Completely.', '', 'Different things here.'], 2)).toMatchObject({ paragraph: 0, offset: 0, found: 'unrelated' });
  });

  test('one paragraph edited in common is something in common', () => {
    const edited = `${CLAUSE} And more.`;
    expect(land([CLAUSE, OTHER], [edited, FOURTH], 1, 5)).toEqual({ paragraph: 1, offset: 0, found: 'gone' });
  });

  test('a document of no text had nothing to lose', () => {
    expect(land(['', ''], [OTHER, CLAUSE], 1).found).toBe('kept');
  });
});

describe('a place in a document', () => {
  test('is the paragraph and the text offset, and comes back to its position, in a table cell too', async () => {
    const session = await EditorDocument.open(await fixture('tables.docx'), {});
    const { doc } = stateFor(session);
    const strings = paragraphStrings(doc);
    expect(strings.length).toBeGreaterThan(100);
    let inCell = 0;
    doc.descendants((node, pos, parent) => {
      if (node.type.name !== 'paragraph' || node.content.size === 0) return true;
      if (parent?.type.name === 'table_cell' || parent?.type.name === 'table_header') inCell = pos + 2;
      return true;
    });
    expect(inCell).toBeGreaterThan(0);
    const place = textPlaceOf(doc, inCell);
    expect(place.offset).toBe(1);
    expect(positionOfPlace(doc, place)).toBe(inCell);
  });

  test('a position between blocks is the next paragraph\'s start; one past the last the last one\'s end', async () => {
    const session = await EditorDocument.open(await fixture('tables.docx'), {});
    const { doc } = stateFor(session);
    expect(textPlaceOf(doc, 0)).toEqual({ paragraph: 0, offset: 0 });
    const strings = paragraphStrings(doc);
    expect(textPlaceOf(doc, doc.content.size + 1)).toEqual({ paragraph: strings.length - 1, offset: strings[strings.length - 1]!.length });
  });

  test('a place read back from its position is the same place, over every fixture', async () => {
    for (const name of await allFixtures()) {
      const session = await EditorDocument.open(await fixture(name), {});
      const { doc } = stateFor(session);
      const strings = paragraphStrings(doc);
      const step = Math.max(1, Math.floor(strings.length / 25));
      for (let paragraph = 0; paragraph < strings.length; paragraph += step) {
        const text = strings[paragraph]!;
        for (const offset of new Set([0, Math.floor(text.length / 2), text.length])) {
          const place = { paragraph, offset };
          expect({ name, ...textPlaceOf(doc, positionOfPlace(doc, place)) }).toEqual({ name, ...place });
        }
      }
    }
  }, 120_000);
});
