// SPDX-License-Identifier: GPL-3.0-or-later
//
// Paths, not blobs (vocabulary section 6; ED-005 section 8.7): a path the agent names is resolved
// against the roots the companion was started with, symlinks resolved first, and refused when it
// lies outside every root. There is no default root.
import { realpath, lstat } from 'node:fs/promises';
import { basename, dirname, isAbsolute, join, relative, resolve } from 'node:path';

/** A refusal, with the code the tool contract's errors carry (ED-005 section 8.3). */
export class PathError extends Error {
  override name = 'PathError';
  constructor(message: string, readonly code: 'outside_root' | 'not_found') {
    super(message);
  }
}

/** Whether `child` is `parent` or lies inside it. */
function inside(parent: string, child: string): boolean {
  const rel = relative(parent, child);
  return rel === '' || (!rel.startsWith('..') && !isAbsolute(rel));
}

/**
 * An existing file's real path, when it lies under one of `roots` (each resolved to its real path
 * too); a relative path is taken from the first root. Refused otherwise, naming the path.
 */
export async function existingUnderRoots(roots: readonly string[], path: string): Promise<string> {
  if (roots.length === 0) throw new PathError('the companion has no root folder (none named with --root, or a VS Code window with no folder open), so it opens nothing', 'outside_root');
  const realRoots = await Promise.all(roots.map((r) => realpath(r)));
  let target: string;
  try {
    target = await realpath(isAbsolute(path) ? path : resolve(realRoots[0]!, path));
  } catch {
    throw new PathError(`no file at ${path}`, 'not_found');
  }
  if (!realRoots.some((r) => inside(r, target))) throw new PathError(`${path} is outside the companion's roots (${roots.join(', ')})`, 'outside_root');
  return target;
}

/**
 * A path to write to (`save`'s and `read`'s `output_path`): its real parent directory under one of
 * `roots`, the file itself resolved when it exists (a symlink out of the roots is refused). Whether
 * it exists, for the caller's overwrite check. A relative path is taken from the first root.
 */
export async function writableUnderRoots(roots: readonly string[], path: string): Promise<{ path: string; exists: boolean }> {
  if (roots.length === 0) throw new PathError('the companion has no root folder (none named with --root, or a VS Code window with no folder open), so it writes nothing', 'outside_root');
  const realRoots = await Promise.all(roots.map((r) => realpath(r)));
  const given = isAbsolute(path) ? path : resolve(realRoots[0]!, path);
  let exists = false;
  try { await lstat(given); exists = true; } catch { exists = false; }
  if (exists) {
    const target = await realpath(given);
    if (!realRoots.some((r) => inside(r, target))) throw new PathError(`${path} is outside the companion's roots (${roots.join(', ')})`, 'outside_root');
    return { path: target, exists: true };
  }
  let parent: string;
  try {
    parent = await realpath(dirname(given));
  } catch {
    throw new PathError(`no folder at ${dirname(path)}`, 'not_found');
  }
  const target = join(parent, basename(given));
  if (!realRoots.some((r) => inside(r, target))) throw new PathError(`${path} is outside the companion's roots (${roots.join(', ')})`, 'outside_root');
  return { path: target, exists: false };
}
