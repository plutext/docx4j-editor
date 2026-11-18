// SPDX-License-Identifier: GPL-3.0-or-later
//
// The command registry (ED-002 section 4): importing this module registers every command.
import './history.mjs';
import './marks.mjs';
import './paragraph.mjs';
import './structure.mjs';
import './sections.mjs';
import './tables.mjs';
import './breaks.mjs';
import './controls.mjs';
import './lists.mjs';
import './links.mjs';
import './tracking.mjs';
import './notes.mjs';
import './comments.mjs';
import './find.mjs';
import './pictures.mjs';
import './pagenumbers.mjs';

export { register, commandNamed, listCommands, asCommand, taggedDispatch, COMMAND_META, sessionPlugin, sessionOf, sessionKey, storyPlugin, storyOf, storyKey, type CommandSpec } from './spec.mjs';
export { usedSdtIds, plainTextControlRefusal, controlToBind, dataBindRefusal, bindsPicture, bindsFragment, insideControlHoldingNoControls, PLACEHOLDER_TEXT, PLACEHOLDER_STYLE, type InsertControlArgs, type BindArgs, type RoleArgs, type PlainTextRefusal, type BindRefusal } from './controls.mjs';
export { runCommand, runCommandAsync, stateFor, type RunResult, type SelectionSpec } from './run.mjs';
export { graphemeSafeRange, atomMarks, isFormattableAtom, applyMarkChange, applyMarkChangeIn, markSetsIn, allHave, setFontField, FONT_SIZES, ATOM_TYPES, type MarkChange } from './marks.mjs';
export { paragraphsIn, updateParagraphs, paragraphValue, twips, points, TWIPS_PER_POINT, INDENT_STEP, ALIGNMENT_TO_JC, JC_TO_ALIGNMENT, tabAttrsOf, TAB_ALIGNMENTS, TAB_LEADERS, type TabStopArg } from './paragraph.mjs';
export { nextStyleOf, inRunControl } from './structure.mjs';
export { linkAt, hyperlinkTarget, insertLinkedText, documentBookmarks, bookmarkPosition, HYPERLINK_STYLE, type HyperlinkArgs, type LinkReading } from './links.mjs';
export { numberingOf, numberingParts, numberingOfPart, numberingPPrOf, listRefOf, levelIndentOf, listKindOf, detachPatch, toggleCase, clampLevel, LEVELS, type Numbering, type ListKind, type ToggleCase } from './lists.mjs';
export { sectionAt, updateSection, type SectionTarget } from './sections.mjs';
export { tableAround, normalizeRowGroups, repairStacks, TableMap, CellSelection, isInTable } from './tables.mjs';
export { OPTIONAL_BREAK, WORD_RUN, breakOffsetsIn } from './breaks.mjs';
export { selectionFormatting, effectiveFormatting, paragraphPropertiesFor, type SelectionFormatting, type FontReading, type ParagraphReading, type Mixed } from './reads.mjs';
export { WORD_KEYMAP, type KeyBinding } from './keymap.mjs';
export { REPLACED, replacedCountOf, finderFor, replaceAllInStory, type ReplaceArgs } from './find.mjs';
export { NOTE_FOCUS, noteFocusOf, NOTE_STYLES, newReference, type InsertNoteArgs, type DeleteNoteArgs } from './notes.mjs';
export { COMMENT_FOCUS, commentFocusOf, commentAuthor, commentTarget, newCommentReference, type InsertCommentArgs, type ReplyArgs, type CommentIdArgs } from './comments.mjs';
export { TRACKING_CHANGED, RESOLVED_META, type TrackingModeArg, type ResolveArgs } from './tracking.mjs';
