// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 4.10 item 2, the OpenDoPE half: the tag syntax against the specification's
// grammar and examples (its section 3), the parts read and written (REQ-012), the data part
// rules (section 2.2), the boolean conversion modes of table 7, the condition tree of table 6
// with its cycle and missing-reference errors (REQ-029, REQ-030), the repeat base and count
// (section 8.1, REQ-038, REQ-039).
import { describe, test, expect } from 'vitest';
import {
  EditorDocument, parseTag, formatTag, readTag, withTagParam, tagParam, readXPaths, writeXPaths, readConditions, writeConditions,
  xpathsProblems, conditionsProblems, versionNewerThan, selectDataPart, evaluateEntryBoolean, evaluateCondition, evaluateEntryValue,
  repeatBase, repeatCount, repeatLooksLikeContainer, freshId, booleanIn, OpenDopeError, XPATHS_NS, CONDITIONS_NS,
  type EvaluationContext, type ConditionsPart, type XPathsPart,
} from '../src/index.mjs';
import { DefaultXPathEngine } from '@docx4j/core-ts';
import { fixture } from './helpers.mjs';

describe('the tag syntax (specification section 3)', () => {
  test('the specification\'s example: two parameters, values literal', () => {
    expect(parseTag('od:repeat=x2&od:finish=t_r')).toEqual([{ key: 'od:repeat', value: 'x2' }, { key: 'od:finish', value: 't_r' }]);
  });

  test('keys are percent-decoded with + a space; values keep = and are not decoded; a bare parameter is its own value', () => {
    expect(parseTag('a%20b+c=1%202&bare&x=a=b')).toEqual([{ key: 'a b c', value: '1%202' }, { key: 'bare', value: 'bare' }, { key: 'x', value: 'a=b' }]);
    expect(parseTag('')).toEqual([]);
    expect(parseTag(undefined)).toEqual([]);
    expect(parseTag('od:xpath=x1&')).toEqual([{ key: 'od:xpath', value: 'x1' }]);
  });

  test('formatTag is the inverse and refuses & in a value', () => {
    const tag = 'od:xpath=x1&customer+name=customer name';
    expect(formatTag(parseTag(tag))).toBe(tag);
    expect(formatTag([{ key: 'a b', value: 'c=d' }, { key: 'p%q', value: '' }])).toBe('a+b=c=d&p%25q=');
    expect(() => formatTag([{ key: 'k', value: 'a&b' }])).toThrow(OpenDopeError);
  });

  test('withTagParam sets in place, appends, and removes', () => {
    const params = parseTag('od:xpath=x1&od:finish=t');
    expect(formatTag(withTagParam(params, 'od:xpath', 'x9'))).toBe('od:xpath=x9&od:finish=t');
    expect(formatTag(withTagParam(params, 'od:condition', 'c1'))).toBe('od:xpath=x1&od:finish=t&od:condition=c1');
    expect(formatTag(withTagParam(params, 'od:xpath', undefined))).toBe('od:finish=t');
    expect(tagParam(params, 'od:finish')).toBe('t');
  });

  test('readTag: the roles, and the problems REQ-007 and REQ-077 name', () => {
    const r = readTag('od:repeat=x2&od:finish=t_r&od:RptPosCon=p1');
    expect(r).toMatchObject({ role: 'repeat', roleId: 'x2', finish: 't_r', positionCondition: 'p1', problems: [] });
    expect(readTag('od:condition=c1&od:repeat=x1').problems[0]).toMatch(/REQ-007/);
    expect(readTag('od:rptd=x2&od:RptOcc=1').problems[0]).toMatch(/REQ-077/);
    expect(readTag('od:rptd=x2&od:RptOcc=1').processorSet.map((p) => p.key)).toEqual(['od:rptd', 'od:RptOcc']);
    expect(readTag('od:xpath=x7&od:ContentType=application/xhtml+xml')).toMatchObject({ xpath: 'x7', contentType: 'application/xhtml+xml', role: 'none' });
    expect(readTag('').role).toBe('none');
  });
});

describe('the parts read and written', () => {
  test('invoice.docx: the XPaths and Conditions parts as records; no version means 2.3 (REQ-009) and the java mode', async () => {
    const session = await EditorDocument.open(await fixture('invoice.docx'));
    await session.template.load();
    const t = session.template;
    expect(t.xpaths!.entries.map((e) => e.id)).toEqual(['x1', 'x2', 'x3', 'x4', 'x5', 'x6']);
    expect(t.xpaths!.entries[1]).toMatchObject({ xpath: '/invoice[1]/items/item', storeItemID: '{8b049945-9dfe-4726-9de9-cf5691e53858}', prefixMappings: '' });
    expect(t.xpaths!.version).toBeUndefined();
    expect(t.mode).toBe('java');
    expect(t.conditions!.conditions).toEqual([
      { id: 'c5', attrs: {}, operand: { op: 'xpathref', id: 'x5' } },
      { id: 'c6', attrs: {}, operand: { op: 'xpathref', id: 'x6' } },
    ]);
    expect(t.problems.map((p) => p.requirement)).toEqual(['REQ-009']);
    expect(t.editable.ok).toBe(true);
    expect(t.isTemplate).toBe(true);
    expect(t.dataPart).toMatchObject({ rule: 1 });
    expect(t.dataPart.part!.part.partName.name).toBe('/customXml/item3.xml');
  });

  test('an entry whose binding attributes a writer qualified with the WordprocessingML namespace is read all the same (ED-005 section 12.87)', async () => {
    // core-ts's check 38 template: xpaths.xsd declares the attributes unqualified, and docx4j writes them so; the
    // engine's build wrote w:storeItemID and w:xpath, which the editor had read as no expression and no part.
    const session = await EditorDocument.open(await fixture('opendope/xpaths-qualified-attributes.docx'));
    await session.template.load();
    const record = readXPaths(session.template.openDopeParts.xpaths!);
    expect(record.entries.map((e) => [e.id, e.xpath])).toEqual([['x1', '/data/logo'], ['x2', '/data/wide'], ['x3', '/data/svg']]);
    expect(record.entries.every((e) => e.storeItemID !== '')).toBe(true);
    expect(session.template.dataPart.part).toBeDefined();
  });

  test('writeXPaths keeps the root attributes and every entry attribute it did not understand (REQ-012); writeConditions likewise', async () => {
    const session = await EditorDocument.open(await fixture('invoice.docx'));
    await session.template.load();
    const part = session.template.openDopeParts.xpaths!;
    const record = readXPaths(part);
    record.entries[0]!.attrs['name'] = 'Customer';
    record.entries[0]!.attrs['questionID'] = 'q1';
    record.entries[0]!.attrs['fieldWidth'] = '20';
    record.attrs['x-custom'] = 'kept';
    record.version = '3.0';
    record.booleanConversion = 'xpath1';
    writeXPaths(part, record);
    const again = readXPaths(part);
    expect(again.entries[0]!.attrs).toEqual({ name: 'Customer', questionID: 'q1', fieldWidth: '20' });
    expect(again.entries.map((e) => e.id)).toEqual(['x1', 'x2', 'x3', 'x4', 'x5', 'x6']);
    expect(again.attrs['x-custom']).toBe('kept');
    expect(again.version).toBe('3.0');
    expect(again.booleanConversion).toBe('xpath1');
    // The prefix as the part spells it: docx4j wrote `ns15:xpaths`, so the entries are `ns15:xpath`.
    expect(part.getXml()).toMatch(/<ns15:xpath id="x1" name="Customer"/);
    const cPart = session.template.openDopeParts.conditions!;
    const conditions = readConditions(cPart);
    conditions.conditions.push({ id: 'c7', attrs: { name: 'both' }, operand: { op: 'and', operands: [{ op: 'conditionref', id: 'c5' }, { op: 'not', operand: { op: 'xpathref', id: 'x6' } }] } });
    writeConditions(cPart, conditions);
    expect(readConditions(cPart).conditions[2]).toEqual({ id: 'c7', attrs: { name: 'both' }, operand: { op: 'and', operands: [{ op: 'conditionref', id: 'c5' }, { op: 'not', operand: { op: 'xpathref', id: 'x6' } }] } });
    expect(cPart.getXml()).toMatch(/<ns16:condition id="c7" name="both"><ns16:and><ns16:conditionref id="c5"\/><ns16:not><ns16:xpathref id="x6"\/><\/ns16:not><\/ns16:and><\/ns16:condition>/);
    // A condition cannot hold a conditionref directly (table 6).
    conditions.conditions.push({ id: 'c8', attrs: {}, operand: { op: 'conditionref', id: 'c5' } });
    expect(() => writeConditions(cPart, conditions)).toThrow(/conditionref directly/);
  });

  test('problems: duplicate ids (REQ-013, REQ-028), missing references and cycles (REQ-029), unknown modes (REQ-031), newer versions (REQ-010)', () => {
    const xpaths: XPathsPart = { attrs: {}, version: '3.0', booleanConversion: 'xpath2', entries: [
      { id: 'x1', xpath: '/a', storeItemID: '{1}', prefixMappings: '', attrs: {} },
      { id: 'x1', xpath: '', storeItemID: '{1}', prefixMappings: '', attrs: {} },
    ] };
    expect(xpathsProblems(xpaths).map((p) => p.requirement)).toEqual(['REQ-013', 'REQ-014']);
    expect(xpathsProblems({ attrs: { booleanConversion: 'perl' }, version: '3.0', entries: [] }).map((p) => p.requirement)).toEqual(['REQ-031']);
    expect(xpathsProblems({ attrs: {}, version: '3.1', booleanConversion: 'xpath2', entries: [] }).map((p) => p.requirement)).toEqual(['REQ-010']);
    expect(versionNewerThan('3.1', '3.0')).toBe(true);
    expect(versionNewerThan('2.3', '3.0')).toBe(false);
    expect(versionNewerThan('3.0', '3.0')).toBe(false);
    expect(versionNewerThan('banana', '3.0')).toBe(true);
    const conditions: ConditionsPart = { attrs: {}, conditions: [
      { id: 'c1', attrs: {}, operand: { op: 'and', operands: [{ op: 'conditionref', id: 'c2' }] } },
      { id: 'c2', attrs: {}, operand: { op: 'not', operand: { op: 'conditionref', id: 'c1' } } },
      { id: 'c3', attrs: {}, operand: { op: 'or', operands: [{ op: 'conditionref', id: 'c9' }, { op: 'xpathref', id: 'x9' }] } },
      { id: 'c3', attrs: {}, operand: { op: 'xpathref', id: 'x1' } },
    ] };
    const problems = conditionsProblems(conditions, xpaths);
    expect(problems.map((p) => p.requirement).sort()).toEqual(['REQ-028', 'REQ-029', 'REQ-029', 'REQ-029']);
    expect(problems.find((p) => /cycle/.test(p.message))!.message).toMatch(/c1 -> c2 -> c1/);
  });
});

describe('the data part (specification section 2.2)', () => {
  test('rule 2 (the first binding), rule 3 (the sole candidate) and rule 4 (several candidates, REQ-003)', async () => {
    const session = await EditorDocument.open(await fixture('HelloWordOnline.docx'));
    await session.template.load();
    // (A part is compared by identity: a deep comparison would walk its DOM.)
    const choice = (bindings: { storeItemID: string; prefixMappings: string }[] = [], xpaths?: XPathsPart) => { const c = selectDataPart(session.template.parts, xpaths, bindings); return [c.rule, c.part]; };
    expect(choice()).toEqual([4, undefined]);
    const a = session.pkg.customXmlParts.add('<a><b>1</b></a>');
    session.template.invalidate();
    expect(choice()).toEqual([3, a]);
    const b = session.pkg.customXmlParts.add('<c/>');
    session.template.invalidate();
    expect(choice()).toEqual([4, undefined]);
    expect(choice([{ storeItemID: b.id.toLowerCase(), prefixMappings: '' }])).toEqual([2, b]);
    // Bindings to the property stores are skipped (case-insensitive ids, REQ-005).
    expect(choice([{ storeItemID: b.id, prefixMappings: "xmlns:ns0='http://schemas.microsoft.com/office/2006/coverPageProps'" }, { storeItemID: a.id, prefixMappings: '' }])).toEqual([2, a]);
    // An XPaths part with entries that resolve nowhere: undetermined, not rule 2.
    expect(choice([{ storeItemID: a.id, prefixMappings: '' }], { attrs: {}, entries: [{ id: 'x1', xpath: '/a', storeItemID: '{0000}', prefixMappings: '', attrs: {} }] })).toEqual([4, undefined]);
  });

  test('binding-simple.docx has no OpenDoPE parts: rule 2 finds the part the first control binds', async () => {
    const session = await EditorDocument.open(await fixture('opendope/binding-simple.docx'));
    await session.template.load();
    expect(session.template.dataPart.rule).toBe(2);
    expect(session.template.isTemplate).toBe(true);
    expect(session.template.openDopeParts.xpaths).toBeUndefined();
  });
});

/** An evaluation context over one custom XML part added to a fixture, with the entries and conditions given. */
async function contextOver(xml: string, entries: Record<string, string>, conditions: ConditionsPart['conditions'] = [], mode: EvaluationContext['mode'] = 'java'): Promise<EvaluationContext> {
  const session = await EditorDocument.open(await fixture('HelloWordOnline.docx'));
  const part = session.pkg.customXmlParts.add(xml);
  await session.template.load();
  return {
    parts: session.template.parts,
    engine: session.pkg.xpathEngine,
    mode,
    xpaths: { attrs: {}, entries: Object.entries(entries).map(([id, xpath]) => ({ id, xpath, storeItemID: part.id, prefixMappings: '', attrs: {} })) },
    conditions: { attrs: {}, conditions },
  };
}

describe('a data part made in the editor (section 11.2 item 31)', () => {
  test('addDataPart gives a document with no data part one, which rule 3 names; malformed XML and an OpenDoPE part are refused; the save carries the part', async () => {
    const original = await fixture('HelloWordOnline.docx');
    const session = await EditorDocument.open(original);
    await session.template.load({ engine: false });
    const t = session.template;
    expect(t.dataPart.rule).toBe(4);
    expect(() => t.addDataPart('<invoice><customer>')).toThrow(/not a custom XML part/);
    expect(() => t.addDataPart('<xpaths xmlns="http://opendope.org/xpaths"/>')).toThrow(/OpenDoPE/);
    expect(t.parts).toHaveLength(0);
    const part = t.addDataPart('<invoice xmlns="urn:acme:invoice"><customer/></invoice>');
    expect(t.dataPart).toMatchObject({ rule: 3, part });
    expect(part.namespaceUri).toBe('urn:acme:invoice');
    expect(session.touchedParts.has(part.part.partName.toString())).toBe(true);
    expect(t.touched).toBe(false);   // no OpenDoPE part was written
    const saved = await session.save();
    expect(saved.changed).toBe(true);
    const reopened = await EditorDocument.open(saved.bytes);
    await reopened.template.load({ engine: false });
    expect(reopened.template.dataPart.rule).toBe(3);
    expect(reopened.template.dataPart.part!.documentElement.childElements.map((c) => c.baseName)).toEqual(['customer']);
    expect(reopened.template.openDopeParts.xpaths).toBeUndefined();
  });
});

describe('boolean conversion (table 7) and conditions (table 6)', () => {
  const DATA = '<d><t>true</t><T>True</T><f>false</f><one>1</one><yes>yes</yes><empty/><n>3</n></d>';
  const ENTRIES = { t: '/d/t', T: '/d/T', f: '/d/f', one: '/d/one', yes: '/d/yes', empty: '/d/empty', missing: '/d/nothere', gt: '/d/n > 2', count: 'count(/d/*) > 100' };

  test('java: the string equals true ignoring case; 1, yes and false are false', async () => {
    const ctx = await contextOver(DATA, ENTRIES, [], 'java');
    const results = Object.fromEntries(Object.keys(ENTRIES).map((id) => [id, evaluateEntryBoolean(ctx, id)]));
    expect(results).toEqual({ t: true, T: true, f: false, one: false, yes: false, empty: false, missing: false, gt: true, count: false });
  });

  test('xpath1: boolean(): a non-empty node-set or string is true, so "false" is true', async () => {
    const ctx = await contextOver(DATA, ENTRIES, [], 'xpath1');
    const results = Object.fromEntries(Object.keys(ENTRIES).map((id) => [id, evaluateEntryBoolean(ctx, id)]));
    expect(results).toEqual({ t: true, T: true, f: true, one: true, yes: true, empty: true, missing: false, gt: true, count: false });
  });

  test('xpath2 through the FontoXPath engine: xs:boolean of the expression, as docx4j\'s cast2 (ConditionBooleanSemanticsTest); refused over the default engine (REQ-032)', async () => {
    const ctx = await contextOver(DATA, { ...ENTRIES, two: '/d/*[self::t or self::f]', padded: '/d/padded' }, [], 'xpath2');
    const results = Object.fromEntries(['t', 'f', 'one', 'empty', 'missing', 'gt', 'count', 'T', 'yes', 'two'].map((id) => { try { return [id, evaluateEntryBoolean(ctx, id)]; } catch (e) { return [id, `error: ${(e as Error).message.slice(0, 40)}`]; } }));
    expect(results['t']).toBe(true);
    expect(results['f']).toBe(false);
    expect(results['one']).toBe(true);
    expect(results['missing']).toBe(false);   // selects nothing: false
    expect(results['gt']).toBe(true);
    expect(results['count']).toBe(false);
    for (const id of ['T', 'yes', 'empty', 'two']) expect(results[id], id).toMatch(/^error/);   // "True", "yes", an empty element, a multi-node sequence raise
    const plain = new DefaultXPathEngine();
    await plain.ready();
    expect(() => booleanIn(plain, '/d/t', ctx.parts[0]!.document, {}, 'xpath2')).toThrow(/REQ-032/);
    expect(booleanIn(plain, '/d/T', ctx.parts[0]!.document, {}, 'java')).toBe(true);
    expect(booleanIn(plain, '/d/f', ctx.parts[0]!.document, {}, 'xpath1')).toBe(true);
  });

  test('the string value: trimmed, and _xHHHH_ decoded after local-name (REQ-019, REQ-020); undefined when nothing is selected', async () => {
    const ctx = await contextOver('<d><a>  padded  </a><b_x0020_c>x</b_x0020_c></d>', { a: '/d/a', name: 'local-name(/d/*[2])', none: '/d/zzz' });
    expect(evaluateEntryValue(ctx, 'a')).toBe('padded');
    expect(evaluateEntryValue(ctx, 'name')).toBe('b c');
    expect(evaluateEntryValue(ctx, 'none')).toBeUndefined();
    expect(() => evaluateEntryValue(ctx, 'nosuch')).toThrow(/entry nosuch does not exist/);
  });

  test('conditions: and, or, not with short circuit; a missing condition is REQ-030, a missing reference or a cycle REQ-029', async () => {
    const ctx = await contextOver(DATA, ENTRIES, [
      { id: 'c1', attrs: {}, operand: { op: 'and', operands: [{ op: 'xpathref', id: 't' }, { op: 'not', operand: { op: 'xpathref', id: 'f' } }] } },
      { id: 'c2', attrs: {}, operand: { op: 'or', operands: [{ op: 'xpathref', id: 'f' }, { op: 'conditionref', id: 'c1' }] } },
      { id: 'c3', attrs: {}, operand: { op: 'and', operands: [{ op: 'xpathref', id: 'f' }, { op: 'conditionref', id: 'nosuch' }] } },
      { id: 'c4', attrs: {}, operand: { op: 'not', operand: { op: 'conditionref', id: 'c5' } } },
      { id: 'c5', attrs: {}, operand: { op: 'or', operands: [{ op: 'conditionref', id: 'c4' }] } },
    ]);
    expect(evaluateCondition(ctx, 'c1')).toBe(true);
    expect(evaluateCondition(ctx, 'c2')).toBe(true);
    expect(evaluateCondition(ctx, 'c3')).toBe(false);   // short circuit: the missing reference is never reached
    expect(() => evaluateCondition(ctx, 'c4')).toThrow(/cycle: c4 -> c5 -> c4.*REQ-029/);
    expect(() => evaluateCondition(ctx, 'nosuch')).toThrow(/REQ-030/);
  });

  test('invoice.docx in its own mode: c5 (includeBankDetails=true) kept, c6 (wantspam=false) dropped', async () => {
    const session = await EditorDocument.open(await fixture('invoice.docx'));
    await session.template.load();
    const ctx = session.template.evaluation;
    expect(evaluateCondition(ctx, 'c5')).toBe(true);
    expect(evaluateCondition(ctx, 'c6')).toBe(false);
  });
});

describe('repeats (specification section 8.1)', () => {
  test('the repeat base drops one trailing / or [1]', () => {
    expect(repeatBase('/invoice[1]/items[1]/item')).toBe('/invoice[1]/items[1]/item');
    expect(repeatBase('/invoice[1]/items[1]/item[1]')).toBe('/invoice[1]/items[1]/item');
    expect(repeatBase('/invoice[1]/items[1]/item/')).toBe('/invoice[1]/items[1]/item');
    expect(repeatBase('/a/b[1][1]')).toBe('/a/b[1]');
  });

  test('the count is the nodes the base selects (REQ-039); a base that selects the container is warned about (REQ-038)', async () => {
    const session = await EditorDocument.open(await fixture('invoice.docx'));
    await session.template.load();
    const ctx = session.template.evaluation;
    expect(repeatCount(ctx, 'x2')).toBe(3);
    expect(repeatLooksLikeContainer(ctx, 'x2')).toBe(false);
    const container: EvaluationContext = { ...ctx, xpaths: { attrs: {}, entries: [{ id: 'bad', xpath: '/invoice[1]/items[1]', storeItemID: ctx.xpaths!.entries[0]!.storeItemID, prefixMappings: '', attrs: {} }] } };
    expect(repeatCount(container, 'bad')).toBe(1);
    expect(repeatLooksLikeContainer(container, 'bad')).toBe(true);
  });

  test('freshId skips the ids in use', () => {
    expect(freshId('x', ['x1', 'x2', 'x4'])).toBe('x3');
    expect(freshId('c', [])).toBe('c1');
  });
});

describe('namespaces', () => {
  test('the constants are the specification\'s (table 1)', () => {
    expect(XPATHS_NS).toBe('http://opendope.org/xpaths');
    expect(CONDITIONS_NS).toBe('http://opendope.org/conditions');
  });
});

describe('a refused part write', () => {
  test('leaves the cached record and the part as they were (a cycle, REQ-029; a duplicate id, REQ-013)', async () => {
    const session = await EditorDocument.open(await fixture('invoice.docx'));
    await session.template.load({ engine: false });
    const t = session.template;
    const before = JSON.stringify(t.conditions);
    expect(() => t.updateConditions((c) => {
      c.conditions.push({ id: 'c7', attrs: {}, operand: { op: 'conditionref', id: 'c8' } });
      c.conditions.push({ id: 'c8', attrs: {}, operand: { op: 'conditionref', id: 'c7' } });
    })).toThrow(/cycle/);
    expect(JSON.stringify(t.conditions)).toBe(before);
    expect(t.conditions!.conditions.map((c) => c.id)).toEqual(['c5', 'c6']);
    expect(t.touched).toBe(false);
    const entries = JSON.stringify(t.xpaths);
    expect(() => t.updateXPaths((x) => { x.entries.push({ id: 'x1', xpath: '/a', storeItemID: '{1}', prefixMappings: '', attrs: {} }); })).toThrow(/x1/);
    expect(JSON.stringify(t.xpaths)).toBe(entries);
    expect(t.touched).toBe(false);
    expect((await session.save()).changed).toBe(false);
  });
});
