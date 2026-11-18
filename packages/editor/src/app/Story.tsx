// SPDX-License-Identifier: GPL-3.0-or-later
//
// A story mounted into React (ED-002 section 7.1): the `StoryView` lives as long as the element,
// and follows the story it is given.
import { useEffect, useRef } from 'react';
import type { EditorDocument, Story as StoryModel } from '@docx4j/editor-model';
import { StoryView, type StoryViewOptions } from '../view/StoryView.js';

export interface StoryProps extends StoryViewOptions {
  session: EditorDocument;
  story: StoryModel;
  className?: string;
  testId?: string;
}

export function Story({ session, story, className, testId, ...options }: StoryProps) {
  const ref = useRef<HTMLDivElement>(null);
  const viewRef = useRef<StoryView | null>(null);
  const label = options.label;
  const part = options.part;
  useEffect(() => {
    if (!ref.current) return;
    const view = new StoryView(ref.current, session, story, { ...(label ? { label } : {}), ...(part ? { part } : {}) });
    viewRef.current = view;
    return () => { view.destroy(); viewRef.current = null; };
    // The view is rebuilt for another session or part; another story of the same one is an update below.
  }, [session, label, part]);   // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => {
    if (viewRef.current && viewRef.current.state.doc !== story.doc) viewRef.current.update(story);
  }, [story]);
  return <div ref={ref} className={className} data-testid={testId} />;
}
