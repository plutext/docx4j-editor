// SPDX-License-Identifier: GPL-3.0-or-later
//
// The one non-edit write (ED-001 section 11 item 2, decision 12): when the document has been
// edited, `docProps/app.xml` gets `Application` and `AppVersion`, unless turned off.
// `AppVersion` must be `XX.YYYY` (Word 2010 x64 treats anything else as corrupt), so the semver
// is mapped `major.minor.patch` to `MM.mmpp`, suffixes dropped, minor and patch clamped at 99,
// and nothing is written when the result does not validate.
import type { OpcPackage } from '@docx4j/core-ts';

/**
 * The value of `Application`: one constant, decided before the first build leaves Plutext
 * (ED-002 section 8 item 5; ED-001 decision 16). ED-001 section 11.1's value until then.
 */
export const APPLICATION_NAME = 'docx4j-editor';

const APP_VERSION = /^\d{2}\.\d{4}$/;

/** `0.3.1` → `00.0301`; `1.12.4-beta.2` → `01.1204`; `1.123.4` → `01.9904`; undefined when the input is not a version. */
export function appVersionOf(semver: string): string | undefined {
  const m = /^\s*v?(\d+)\.(\d+)(?:\.(\d+))?/.exec(semver);
  if (!m) return undefined;
  const major = Math.min(99, Number(m[1]));
  const minor = Math.min(99, Number(m[2]));
  const patch = Math.min(99, Number(m[3] ?? '0'));
  const out = `${String(major).padStart(2, '0')}.${String(minor).padStart(2, '0')}${String(patch).padStart(2, '0')}`;
  return APP_VERSION.test(out) ? out : undefined;
}

export interface WriteApplicationOptions {
  /** The editor's own version (the package version). */
  version: string;
  /** Off in the options pane: nothing is written. Default true. */
  enabled?: boolean;
  /** Override the application name (tests). */
  application?: string;
}

interface ExtendedProperties {
  TYPE_NAME?: string;
  application?: string;
  appVersion?: string;
}

/**
 * Writes `Application` and `AppVersion` into the package's extended properties, creating the
 * part when absent. Returns true when something was written. Call it only after an edit.
 */
export async function writeApplication(pkg: OpcPackage, options: WriteApplicationOptions): Promise<boolean> {
  if (options.enabled === false) return false;
  const version = appVersionOf(options.version);
  if (version === undefined) return false;
  const part = pkg.docPropsExtendedPart ?? pkg.addDocPropsExtendedPart();
  let props: ExtendedProperties;
  try {
    props = (await part.getContents()) as ExtendedProperties;
  } catch {
    props = { TYPE_NAME: 'org_docx4j_docProps_extended.Properties' };
    part.setContents(props as never);
  }
  props.application = options.application ?? APPLICATION_NAME;
  props.appVersion = version;
  return true;
}
