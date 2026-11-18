// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-002 section 3.7 test 5: the `XX.YYYY` rule of ED-001 section 11 item 2.
import { describe, test, expect } from 'vitest';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { appVersionOf, writeApplication, APPLICATION_NAME } from '../src/index.mjs';

describe('appVersionOf', () => {
  test('maps major.minor.patch to MM.mmpp', () => {
    expect(appVersionOf('0.3.1')).toBe('00.0301');
    expect(appVersionOf('1.12.4')).toBe('01.1204');
    expect(appVersionOf('0.1.0')).toBe('00.0100');
    expect(appVersionOf('16.0.0')).toBe('16.0000');
  });
  test('drops pre-release and build suffixes', () => {
    expect(appVersionOf('0.3.1-beta.2')).toBe('00.0301');
    expect(appVersionOf('0.3.1-SNAPSHOT')).toBe('00.0301');
    expect(appVersionOf('1.2.3+build.7')).toBe('01.0203');
    expect(appVersionOf('v2.0.1')).toBe('02.0001');
  });
  test('clamps minor and patch at 99 and major at 99', () => {
    expect(appVersionOf('1.123.4')).toBe('01.9904');
    expect(appVersionOf('1.2.345')).toBe('01.0299');
    expect(appVersionOf('123.0.0')).toBe('99.0000');
  });
  test('writes nothing for what is not a version', () => {
    expect(appVersionOf('')).toBeUndefined();
    expect(appVersionOf('latest')).toBeUndefined();
    expect(appVersionOf('1')).toBeUndefined();
  });
});

describe('writeApplication', () => {
  test('writes Application and AppVersion, creating the part when absent', async () => {
    const pkg = await WordprocessingMLPackage.createPackage();
    // createPackage gives a package extended properties; drop them to test creation.
    const wrote = await writeApplication(pkg, { version: '0.1.0' });
    expect(wrote).toBe(true);
    const part = pkg.docPropsExtendedPart!;
    const xml = await part.getXml();
    // The facade's prefix table writes the elements prefixed; Word's own file uses the default namespace. Both are valid.
    expect(xml).toMatch(new RegExp(`<(?:[\\w-]+:)?Application>${APPLICATION_NAME}</(?:[\\w-]+:)?Application>`));
    expect(xml).toMatch(/<(?:[\w-]+:)?AppVersion>00\.0100<\/(?:[\w-]+:)?AppVersion>/);
  });
  test('writes nothing when disabled or when the version is invalid', async () => {
    const pkg = await WordprocessingMLPackage.createPackage();
    expect(await writeApplication(pkg, { version: '0.1.0', enabled: false })).toBe(false);
    expect(await writeApplication(pkg, { version: 'nope' })).toBe(false);
  });
});
