// SPDX-License-Identifier: GPL-3.0-or-later
export const NOTICES_FILE: string;
export const ALLOWED: Set<string>;
export function allowed(expression: unknown): boolean;
export function licenceOf(pkg: Record<string, unknown>): string | undefined;
export function packageDirOf(moduleId: string): string | undefined;
export interface PackageNotice {
  name: string;
  version: string;
  licence: string | undefined;
  homepage: string | undefined;
  author: string | undefined;
  dir: string;
}
export function describePackage(dir: string): PackageNotice | undefined;
export function withoutLicenceText(packages: Array<PackageNotice | undefined>): PackageNotice[];
export function noticesText(what: string, packages: Array<PackageNotice | undefined>, licenceFile?: string, more?: string[]): string;
