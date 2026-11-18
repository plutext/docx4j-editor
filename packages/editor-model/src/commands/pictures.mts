// SPDX-License-Identifier: GPL-3.0-or-later
//
// Pictures made in the editor (ED-005 proposal 45 part 2, section 12.76). `range.insertInlinePicture`
// adds the image part and its relationship to the main part in its `prepare` half, as a link's
// relationship is made, builds the `w:drawing` through the engine and registers it under a source id
// (`session.addInlinePicture`), then inserts an `image` node at the selection, which the export emits
// as it emits any inline drawing. Editor-only, no `api` line: Office JS puts a picture at a
// paragraph's start or end (`insertInlinePictureFromBase64`), never at a caret.
import { register, sessionOf, storyOf } from './spec.mjs';
import { insertAtom } from './structure.mjs';

export interface InsertPictureArgs {
  /** The image as base64 (PNG, JPEG, GIF, BMP, TIFF, SVG, EMF or WMF). */
  base64: string;
  /** `wp:docPr/@name`; "Picture N" when omitted. */
  name?: string;
  /** The width in points; the natural size, scaled down to the text width, when omitted. */
  width?: number;
  /** Filled by `prepare`. */
  sourceId?: string;
  widthPt?: number | null;
  heightPt?: number | null;
}

register<InsertPictureArgs>({
  name: 'range.insertInlinePicture',
  label: 'Insert picture',
  args: { base64: 'the image as base64 (PNG, JPEG, GIF, BMP, TIFF, SVG, EMF or WMF)', name: 'wp:docPr/@name', width: 'points; the natural size, scaled down to the text width, when omitted' },
  prepare: async (state, args) => {
    const session = sessionOf(state);
    if (!session || !args?.base64 || storyOf(state) !== undefined) return args;
    const made = session.addInlinePicture(args.base64, { ...(args.name !== undefined ? { name: args.name } : {}), ...(args.width !== undefined ? { width: args.width } : {}) });
    return { ...args, sourceId: made.sourceId, widthPt: made.width, heightPt: made.height };
  },
  run: (args) => (state, dispatch, view) => {
    if (!args?.sourceId) return false;
    const sourceId = args.sourceId;
    return insertAtom((s, attrs) => s.schema.nodes['image']!.create({ ...attrs, sourceId, width: args.widthPt ?? null, height: args.heightPt ?? null }))(state, dispatch, view);
  },
  // The body alone: a note's or a header's part would need the image related from it (as a paste's links are refused there).
  isEnabled: (state) => storyOf(state) === undefined && (state.selection.$from.parent.type.name === 'paragraph' || state.selection.$from.parent.type.name === 'sdtInline'),
});
