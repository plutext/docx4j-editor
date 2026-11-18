// SPDX-License-Identifier: GPL-3.0-or-later

/** A tracked change as the editor lists it (Office JS's form). */
export interface ListedChange { type: string; author: string | null; text: string; row?: true; move?: string }

/** What the builder made (ED-005 section 9.2 item 1). */
export interface ReviewForWord {
  bytes: Uint8Array;
  /** The editor's self-check at the person's save. */
  selfCheck: boolean;
  /** The changes as the editor lists them after reopening the saved file. */
  changes: ListedChange[];
  /** The revisions shown in Word only. */
  preserved: { kind: string; author: string | null }[];
  /** The comments, a thread's replies after it. */
  comments: { author: string; text: string; done?: boolean; reply?: true }[];
}

/** Builds the file in a room: a person's session as two authors, the agent through the companion's tools. */
export function buildReviewForWord(): Promise<ReviewForWord>;

/** The listing as text (review-for-word.txt). */
export function describe(result: ReviewForWord): string;
