// SPDX-License-Identifier: GPL-3.0-or-later
//
// @docx4j/editor-model: the headless editable projection of a docx over @docx4j/core-ts
// (ED-001 sections 2 and 3; ED-002 section 3). Import a body to ProseMirror nodes held in Yjs,
// edit, export back to the typed tree with every untouched element re-emitted as it was.
export { schema, nodes, marks, FORMATTING_MARKS, type EditorSchema, type ParagraphAttrs, type Revision, type RevisionInfo, type PropertiesChange, type NumPrAttrs, type TabAttrs, type IndentAttrs, type SpacingAttrs, type FontAttrs } from './schema.mjs';
export { SourceIds } from './ids.mjs';
// The highlight palette as the engine reads and writes `w:highlight` (objects CR-014: Word's own since 0.6.0): the one table for the view's drawing too.
export { highlightHexValue, highlightNameForColor, HIGHLIGHT_COLORS } from '@docx4j/generated-objects-ts/helpers/wml';
export { projectBody, projectBlocks, newContext, paragraphAttrsOf, projectedParagraphAttrs, formattingMarks, TWIPS_PER_PX, type Projection, type ProjectionContext, type SourceInfo, type ParagraphInfo, type TableInfo, type RowInfo, type SdtGroup, type GroupInfo } from './import.mjs';
export { exportBody, exportBlocks, localRevisionIds, applyParagraphAttrs, formattingKey, marksFromKey, formattingOf, runPropertiesFor, type ExportResult, type ControlLookup, type BlockContainer, type RevisionIds } from './export.mjs';
export * from './opendope.mjs';
export {
  readControl, applyControlProperties, withBinding, withSdtId, newSdtPr, allocateSdtId, sdtItems, sdtItem, withSdtItem, withoutSdtItem,
  sdtTag, sdtTitle, sdtId, sdtKind, sdtBinding, sdtRunProperties, W_NS, W14_NS, W15_NS,
  type SdtPrItem, type SdtPrJson, type SdtKind, type ControlProperties, type ControlReading, type DataBinding, type ContentControlAppearance,
} from './sdtpr.mjs';
export { listControls, controlAt, controlAround, sdtPrOf, rowExtent, cellExtent, rowPosition, isNodeTarget, type ControlSummary, type ControlTarget, type ControlForm, type NodeTarget, type RowGroupTarget, type CellGroupTarget } from './controls.mjs';
export { TemplateModel, type Verdict, type EntryUse } from './template.mjs';
export { previewInstance, type Preview } from './preview.mjs';
export * from './commands/index.mjs';
export { tabStopsOf, paragraphTabs, nextTabStop, tabWidth, DEFAULT_TAB_STOP, type TabStop, type FollowingText } from './tabs.mjs';
export { listLabels, liveLabels, displayLabel, type ListLabel, type LiveLabel, type LabelSuffix } from './labels.mjs';
export { documentStyles, documentFonts, COMMON_FONTS, type StyleEntry } from './styles.mjs';
export { Stories, commentSourceKey, NOTE_STYLES, type Story, type Note, type NoteKind, type CommentThread, type CommentRecord, type NewComment, type HeaderFooter, type HeaderVariant, type PageStories } from './stories.mjs';
export { commentAnchors, commentParagraphs, commentElement, referenceIdOf, initialsOf, dropCommentsResolvedAway, COMMENT_STYLES } from './comments.mjs';
export { APPLICATION_NAME, appVersionOf, writeApplication, type WriteApplicationOptions } from './appxml.mjs';
export {
  newEditStore, loadDocument, readDocument, replaceDocument, syncTextFormatting, alignBlocks, editScript, fragmentIn, readFragment, replaceFragment, LOCAL_ORIGIN, IMPORT_ORIGIN, BODY_FRAGMENT, FOOTNOTES_MAP, ENDNOTES_MAP, COMMENTS_MAP, HEADERS_MAP, CONTROLS_MAP, PACKAGE_MAP, FINAL_SECT_PR,
  readControlEntry, readControlEntries, writeControl, controlWritesOf, addControlWrite, applyControlWrites, CONTROL_WRITES,
  COMMENT_THREADS_MAP, AGENTS_MAP, AGENT_LOG_ARRAY, readThreadEntry, writeThreadFields, commentWritesOf, addCommentWrite, COMMENT_WRITES, storeWritesOf, writeStoreWrite,
  SOURCES_MAP, STYLES_MAP, NUMBERING_MAP, PARTS_MAP, RELATIONSHIPS_MAP, ROOM_MAP, ROOM_SCHEMA, ROOM_ORIGIN, PROJECTION, PROJECTION_KEY, RoomNotReadyError, RoomVersionError, ProjectionVersionError,
  type EditStore, type ControlEntry, type ControlWrite, type ThreadFields, type CommentWrite, type StoreWrite,
} from './yjs.mjs';
export { isBaseId } from './ids.mjs';
export { pathIds, assignPathIds, treeEqual, type SourceEntry } from './sources.mjs';
export { compareProjections, normalizedJson, SelfCheckError, type SelfCheckResult } from './selfcheck.mjs';
export { EditorDocument, compatibilityModeOf, EDITOR_VERSION, PARTS_REVISION, PARTS_ORIGIN, HYPERLINK_RELATIONSHIP, sha256Hex, roomOf, PaginationMismatchError, type RoomRecord, type OpenOptions, type SaveOptions, type SaveResult, type SaveFormat, type TreeBindings, type TreeRunResult, type TreeRunTracking, type TreeRunOptions, type NoteSettings, type SelectXPathOptions } from './session.mjs';
export { moveMarkers, isMarker, type MarkerMove } from './paginate.mjs';
export { type Diagnostics, type OpaqueEntry, emptyDiagnostics, summarize, preservedCount } from './diagnostics.mjs';
export { dateString, revisionInfo, revisionOf, isDeleted, revisionMarksOf, deletionOf, isHolderOf, revisionHolder, formatChangeBefore, formatChangeBeforeRPr, formatChangeAttrs, rPrChangeOf, type FormatChangeAttrs } from './revisions.mjs';
export { graphemeBoundary, isGraphemeBoundary } from './text.mjs';
export { readMarkdown, webAddress, plainBlocks, inlineText, looksLikeMarkdown, buildBlocks, markdownDocument, MarkdownUnsupported, MARKDOWN_STYLES, type MarkdownDocument, type MdBlock, type Inline, type ListPlace, type BuildContext, type MarkdownRead, type ReadOptions, tableColumnWidths } from './markdown.mjs';
export { ensureStylesIn } from './knownstyles.mjs';
export { partEditing, checkPartXml, applyPartXml, type PartEditing, type PartXmlProblem, type AppliedPart, type DanglingReference } from './parts.mjs';
export { selectIn, selectScript, XPATH_NAMESPACES, type XPathHit, type SelectContainer } from './xpath.mjs';
export { compareText, paragraphTextsOf, paragraphTextsOfBytes, paragraphTextsOfSession, compareWithBytes, type AlignedPair, type TextComparison } from './compare.mjs';
export { paragraphStrings, textPlaceOf, positionOfPlace, relocator, type TextPlace, type Landing } from './relocate.mjs';
export { Finder, findIn, atomText, paragraphText, acceptedTextBetween, searchOptionsOf, hasGroupReference, hitSelection, HOLDS_MARK, type FindOptions, type FindContext, type Hit, type BlockHits, type ParagraphText } from './find.mjs';
export { embeddedFaces, deobfuscateFont, fontKeyBytes, looksLikeFont, repairFont, type EmbeddedFace } from './fonts-embedded.mjs';
export { MACROS_NAMESPACE, macrosXml, macrosOf, macrosPart, readDocumentMacros, writeDocumentMacros, type Macro } from './macros.mjs';
export { selectionSpec, selectionBinding, paragraphsWithin, textOffsetIn, type SelectionSpec, type SelectionBinding } from './bindings.mjs';
export { formatNumber, noteLabels, noteNumberingFor, type NoteNumbering } from './numbers.mjs';
export { toPlain, revive, plainEqual } from './json.mjs';
export { heldText } from './excluded.mjs';
export { textPackageOver } from './textmodel.mjs';
export { recordTransaction, trackTransaction, withoutInheritedRevisions, markComposed, trackingDate, UNTRACKED_META, PASTE_META, type Tracking, type Recorded } from './track.mjs';
export { RoomRepair, yAttributeKey, resolutionWriteOf, addResolutionWrite, writeResolution, repairWrites, REPAIR_ORIGIN, ACK_ORIGIN, RESOLUTION_WRITE, type ResolutionWrite, type ResolvedRange } from './resolutions.mjs';
export { trackedChangesOf, emptyChangesOf, isEmptyRun, preservedChangesOf, preservedGroupsOf, listingBlockWalks, whatChanged, changedValues, replacementPairs, liveMapping, resolveChange, resolveChanges, changeKey, changeOfKey, type TrackedChangeInfo, type ChangedValue, type ChangePiece, type ChangeType, type ElementLookup, type ResolveContext, type PreservedChange, type PreservedGroup, type PreservedKind, type SourceLookup } from './review.mjs';
export { preservedTargetOf, listedPreserved, resolvePreservedOnTree, type PreservedTarget, type PreservedTargetKind, type ListedPreserved } from './preserved.mjs';
export { AGENTS_ORIGIN, AGENT_LOG_CAP, isAgentStopped, stopAgent, resumeAgent, agentLog, appendAgentLog, onAgents, type AgentLogEntry } from './agents.mjs';
export { controlLockRefusal, applyControlRules, type LockRefusal } from './controlrules.mjs';
export { cellConditionsOf, tableConditionsOf, conditionalTcPr, type CellConditions, type ParagraphConditions } from './tableconditions.mjs';
export { BOUND_WRITE_META, writeBoundValue, mappedKind, isFlatOpc, boundControlsIn, boundControlsTouched, controlValue, boundNodeOf, writeBackBoundControls, refreshBoundControls, applyBoundRefresh, controlIds, boundControlsWithIds, syncBoundAfter, DEFAULT_PLACEHOLDERS, defaultPlaceholderFor, placeholderFor, namedDefaultPlaceholder, type BoundControl, type DefaultPlaceholder } from './bound.mjs';
export { XHTML_CONTENT_TYPE, isXhtmlControl, xhtmlBindingOf, xhtmlControlsIn, xhtmlKindOf, type XhtmlControl } from './xhtml.mjs';
export { WORD_DOCUMENT_PROGID, isFragmentControl, fragmentBindingOf, fragmentFactsOf, fragmentControlsIn, refreshFragments, type FragmentBinding, type FragmentFacts, type FragmentControl, type FragmentRefresh } from './fragments.mjs';
export { isPictureHandler, pictureBindingOf, pictureControlsIn, type PictureBinding, type PictureControl, looksLikeImageBase64, base64Bytes, imageBase64Of, pictureFactsOf, withPictureWidth, fittedInBox, type PictureFacts } from './pictures.mjs';
export { inlineDrawingOf, inlineOf, blipEmbedOf, type InlineDrawing } from './drawings.mjs';
export { xpathSteps, relativeToRepeat } from './opendope.mjs';
export { endCellsWithParagraph, tableAutoFitOf, textWidthTwips, columnWidthsPx, columnWidthsPt, setColumnWidths, insertGridColumns, columnWidthsAside, clearColumnWidths, MIN_COLUMN_WIDTH_PX, rowHeightOf, withoutRowHeight, withCellShading, rowShadingOf, type TableAutoFit } from './commands/tables.mjs';
export { pageNumberParagraph, hasPageField, planPageNumbers, PAGE_NUMBER_STYLE, type PageNumberPlan, type PageNumberVariant } from './pagenumbers.mjs';
export { staleBoundControls, staleTextControls, staleXhtmlControls, stalePictureControls, staleFragmentControls, textOfMarkup, type StaleBoundControls } from './stale.mjs';
export { pastedSlice, closedAtMark, commentLengthsIn, commentLengthIn, clipboardToken, CLIPBOARD_SLICE_ATTR, CLIPBOARD_SESSION_ATTR, type PasteRules } from './clipboard.mjs';
export { newCommentReference } from './commands/comments.mjs';
export { newReference as newNoteReference } from './commands/notes.mjs';
