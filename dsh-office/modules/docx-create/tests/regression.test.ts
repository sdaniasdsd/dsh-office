import { describe, expect, it } from 'vitest';
import { unzipSync } from 'fflate';
import { createDocxCreateModule } from '../src/index';
import { memoryStore, part, rewrite, asBytes, simple } from './helpers';
import { parseXml, W } from '../src/engine/xml';

describe('template and layout regression', () => {
  it('fills split runs in body, table and headers, preserving the source and unrelated parts', async () => {
    const memory = memoryStore(); const module = createDocxCreateModule({ artifactStore: memory.store });
    const template = await module.handlers.execute({ ...simple, plan: { kind: 'create', document: {
      header: '{{company}}', footer: '{{year}}', blocks: [
        { kind: 'paragraph', id: 'title', runs: [{ text: '欢迎 {{na', bold: true }, { text: 'me}}，{{name}}' }] },
        { kind: 'table', id: 'table', rows: [['项目'], ['{{project}}']] },
      ],
    } } });
    const ref = template.result.artifactRef; const before = memory.files.get(ref.id)!.slice();
    const out = await module.handlers.execute({ operation: 'execute', requestId: 'fill', plan: { kind: 'fillTemplate', templateRef: ref,
      values: { name: '张三 & <李四>', company: '示例公司', year: '2026', project: '文档模块' } } });
    const after = memory.files.get(out.result.artifactRef.id)!;
    expect(memory.files.get(ref.id)).toEqual(before); expect(out.result.filledKeys).toEqual(['company', 'name', 'project', 'year']);
    const doc = parseXml(part(after));
    expect(Array.from(doc.getElementsByTagNameNS(W, 't')).map(t => t.textContent).join('')).toContain('欢迎 张三 & <李四>，张三 & <李四>');
    expect(part(after)).toContain('<w:b');
    expect(part(after, 'word/header1.xml')).toContain('示例公司');
    const a = unzipSync(before), b = unzipSync(after);
    for (const name of Object.keys(a)) if (!['word/document.xml', 'word/header1.xml', 'word/footer1.xml'].includes(name)) expect(b[name]).toEqual(a[name]);
    expect(out.result.paragraphIds).toEqual(template.result.paragraphIds);
  });
  it('does not evaluate or recursively substitute replacement text', async () => {
    const memory = memoryStore(); const module = createDocxCreateModule({ artifactStore: memory.store });
    const template = await module.handlers.execute({ ...simple, plan: { kind: 'create', document: { blocks: [{ kind: 'paragraph', id: 'p', runs: [{ text: '{{name}}' }] }] } } });
    const out = await module.handlers.execute({ operation: 'execute', requestId: 'fill', plan: { kind: 'fillTemplate', templateRef: template.result.artifactRef, values: { name: '{{other}} ${process.exit()}' } } });
    expect(part(memory.files.get(out.result.artifactRef.id)!)).toContain('{{other}} ${process.exit()}');
  });
  it('rejects missing and unused values and unsupported multiline replacement', async () => {
    const memory = memoryStore(); const module = createDocxCreateModule({ artifactStore: memory.store });
    const template = await module.handlers.execute({ ...simple, plan: { kind: 'create', document: { blocks: [{ kind: 'paragraph', id: 'p', runs: [{ text: '{{name}}' }] }] } } });
    for (const [values, code] of [[{}, 'TEMPLATE_VALUE_MISSING'], [{ name: 'a', extra: 'b' }, 'TEMPLATE_VALUE_UNUSED'], [{ name: 'a\nb' }, 'UNSUPPORTED_TEMPLATE']] as const) {
      await expect(module.handlers.execute({ operation: 'execute', requestId: 'fill', plan: { kind: 'fillTemplate', templateRef: template.result.artifactRef, values } })).rejects.toMatchObject({ code });
    }
    expect(memory.files.size).toBe(1);
  });
  it('retains unknown XML in changed template parts', async () => {
    const memory = memoryStore(); const module = createDocxCreateModule({ artifactStore: memory.store });
    const template = await module.handlers.execute({ ...simple, plan: { kind: 'create', document: { blocks: [{ kind: 'paragraph', id: 'p', runs: [{ text: '{{name}}' }] }] } } });
    const bytes = memory.files.get(template.result.artifactRef.id)!;
    const custom = memory.put('custom', rewrite(bytes, e => { e['word/document.xml'] = asBytes(part(bytes).replace('<w:sectPr>', '<custom:record xmlns:custom="urn:test" value="keep"/><w:sectPr>')); }));
    const out = await module.handlers.execute({ operation: 'execute', requestId: 'fill', plan: { kind: 'fillTemplate', templateRef: custom, values: { name: 'filled' } } });
    expect(part(memory.files.get(out.result.artifactRef.id)!)).toContain('value="keep"');
  });
  it('generates XML escaped text with explicit line/tab elements and stable paragraph IDs', async () => {
    const memory = memoryStore(); const module = createDocxCreateModule({ artifactStore: memory.store });
    const make = async (text: string) => module.handlers.execute({ ...simple, plan: { kind: 'create', document: { pageNumbers: false, blocks: [{ kind: 'paragraph', id: 'fixed', runs: [{ text }] }] } } });
    const a = await make('A & B\nC\tD'); const b = await make('changed');
    expect(a.result.paragraphIds).toEqual(b.result.paragraphIds); expect(a.result.fields).toBe('none');
    const xml = part(memory.files.get(a.result.artifactRef.id)!);
    expect(xml).toContain('A &amp; B'); expect(xml).toContain('<w:br/>'); expect(xml).toContain('<w:tab/>');
  });
  it('preserves unreferenced vendor parts and fills even when delimiters are split', async () => {
    const memory = memoryStore(); const module = createDocxCreateModule({ artifactStore: memory.store });
    const template = await module.handlers.execute({ ...simple, plan: { kind: 'create', document: { blocks: [
      { kind: 'paragraph', id: 'p', runs: [{ text: '{' }, { text: '{name' }, { text: '}' }, { text: '}' }] },
    ] } } });
    const bytes = memory.files.get(template.result.artifactRef.id)!;
    const added = memory.put('vendor', rewrite(bytes, entries => { entries['custom/vendor.xml'] = asBytes('<vendor xmlns="urn:vendor">keep untouched</vendor>'); }));
    const out = await module.handlers.execute({ operation: 'execute', requestId: 'fill', plan: { kind: 'fillTemplate', templateRef: added, values: { name: '完整替换' } } });
    const saved = memory.files.get(out.result.artifactRef.id)!;
    expect(part(saved)).toContain('完整替换');
    expect(unzipSync(saved)['custom/vendor.xml']).toEqual(unzipSync(memory.files.get('vendor')!)['custom/vendor.xml']);
  });
  it('refuses protected templates instead of silently bypassing editing restrictions', async () => {
    const memory = memoryStore(); const module = createDocxCreateModule({ artifactStore: memory.store });
    const template = await module.handlers.execute({ ...simple, plan: { kind: 'create', document: { blocks: [{ kind: 'paragraph', id: 'p', runs: [{ text: '{{name}}' }] }] } } });
    const bytes = memory.files.get(template.result.artifactRef.id)!;
    const added = memory.put('protected', rewrite(bytes, entries => {
      entries['word/settings.xml'] = asBytes(part(bytes, 'word/settings.xml').replace('</w:settings>', '<w:documentProtection w:edit="readOnly" w:enforcement="1"/></w:settings>'));
    }));
    await expect(module.handlers.execute({ operation: 'execute', requestId: 'fill', plan: { kind: 'fillTemplate', templateRef: added, values: { name: 'forbidden' } } })).rejects.toMatchObject({ code: 'UNSUPPORTED_TEMPLATE' });
  });
});
