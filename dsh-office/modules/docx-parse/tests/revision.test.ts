/**
 * 修订层测试 —— 「新旧修改看得准、外部意图收得住」。
 *
 * 最核心的一条断言在「中间插入一段不会让后面所有节点都失联」那个用例：
 * 它是双 IR 身份方案存在的全部理由，也是通用文本 diff 做不到的事。
 */
import { describe, expect, it, vi } from 'vitest';

import { diffDualIR, irFingerprint } from '../src/revision/diff';
import {
  createMemoryRevisionStore,
  createRevisionLedger,
  validateIntent,
} from '../src/revision/ledger';
import type { RevisionStore } from '../src/revision/contract';
import { buildIR, intent, para } from './support';

/** 造一份「三段普通文字」的双 IR。 */
function threeParagraphs(texts: string[] = ['Alpha', 'Beta', 'Gamma']) {
  return buildIR(texts.map((text, index) => para(text, index)));
}

describe('revision: diff by stable identity', () => {
  it('reports a text edit as one modification, not remove+add', () => {
    const before = threeParagraphs();
    const after = threeParagraphs(['Alpha', 'Beta changed', 'Gamma']);
    const changes = diffDualIR(before, after);

    // 只有第二段变了；稳定 id 让我们认出「这是同一个节点被改了」。
    expect(changes).toHaveLength(1);
    expect(changes[0]?.change).toBe('modified');
    expect(changes[0]?.fields).toEqual(['text']);
    expect(changes[0]?.before).toBe('Beta');
    expect(changes[0]?.after).toBe('Beta changed');
    // 能直接拿到物理落点，写入端不用再查一次。
    expect(changes[0]?.pointers[0]).toContain('word/document.xml!');
  });

  it('sees a format-only change that leaves the text untouched', () => {
    const before = buildIR([para('Same words', 0)]);
    const after = buildIR([para('Same words', 0, { styleId: 'Emphasis', styleName: 'Emphasis' })]);
    const changes = diffDualIR(before, after);

    // 一个字都没改，但样式换了 —— 这正是通用文本 diff 会漏掉的那类改动。
    expect(changes).toHaveLength(1);
    expect(changes[0]?.change).toBe('modified');
    expect(changes[0]?.fields).toEqual(['style']);
  });

  it('sees a heading level change', () => {
    const before = buildIR([para('Section', 0, { outlineLevel: 1, headingStyle: true })]);
    const after = buildIR([para('Section', 0, { outlineLevel: 2, headingStyle: true })]);
    const changes = diffDualIR(before, after);

    expect(changes[0]?.fields).toEqual(['level']);
    expect(changes[0]?.before).toBe('H2 Section');
    expect(changes[0]?.after).toBe('H3 Section');
  });

  it('reports added and removed nodes', () => {
    const before = buildIR([para('Alpha', 0), para('Beta', 1)]);
    const after = buildIR([para('Alpha', 0), para('Delta', 1)]);

    // 第二段的 id 由结构路径推导 —— 位置没变但内容换成完全不同的文字，
    // 仍然算「同一个节点被改」；真正的增删要看位置位移。
    const changes = diffDualIR(before, after);
    expect(changes.map((entry) => entry.change)).toEqual(['modified']);

    const shrunk = diffDualIR(before, buildIR([para('Alpha', 0)]));
    expect(shrunk.some((entry) => entry.change === 'removed')).toBe(true);

    const grown = diffDualIR(before, buildIR([para('Alpha', 0), para('Beta', 1), para('Gamma', 2)]));
    expect(grown.some((entry) => entry.change === 'added')).toBe(true);
  });

  it('does not cascade when a paragraph is inserted mid-document and paraIds exist', () => {
    // 关键用例：有原生 paraId 时，id 不随位置漂移。
    const before = buildIR([
      para('Alpha', 0, { paraId: 'AAAAAAAA' }),
      para('Gamma', 1, { paraId: 'CCCCCCCC' }),
    ]);
    const after = buildIR([
      para('Alpha', 0, { paraId: 'AAAAAAAA' }),
      para('Beta', 1, { paraId: 'BBBBBBBB' }),
      para('Gamma', 2, { paraId: 'CCCCCCCC' }),
    ]);

    const changes = diffDualIR(before, after);
    // 只有新增那一段；Alpha 与 Gamma 完全不受影响。
    expect(changes).toHaveLength(1);
    expect(changes[0]?.change).toBe('added');
    expect(changes[0]?.after).toBe('Beta');
  });

  it('omits unchanged nodes by default and includes them on request', () => {
    const before = threeParagraphs();
    const after = threeParagraphs(['Alpha', 'Beta changed', 'Gamma']);

    expect(diffDualIR(before, after)).toHaveLength(1);
    const withUnchanged = diffDualIR(before, after, { includeUnchanged: true });
    expect(withUnchanged).toHaveLength(3);
    expect(withUnchanged.filter((entry) => entry.change === 'unchanged')).toHaveLength(2);
  });

  it('is deterministic: same inputs give byte-identical output', () => {
    const before = threeParagraphs();
    const after = threeParagraphs(['Alpha', 'Beta changed', 'Gamma']);
    expect(JSON.stringify(diffDualIR(before, after))).toBe(
      JSON.stringify(diffDualIR(before, after)),
    );
  });

  it('covers table cell paragraphs so single-cell edits are visible', () => {
    // 改一格必须能在差异里看见；否则最常见的办公改动会完全隐身。
    const changes = diffDualIR(singleCellTable('A1'), singleCellTable('A1 changed'));

    expect(changes).toHaveLength(1);
    expect(changes[0]?.kind).toBe('paragraph');
    expect(changes[0]?.after).toBe('A1 changed');
  });
});

/** 造一个 1x1 表格，便于验证「单元格段落可被精确识别」。 */
function singleCellTable(cellText: string) {
  const cellPath = '/w:document/w:body/w:tbl[1]/w:tr[1]/w:tc[1]/w:p[1]';
  return buildIR([
    {
      kind: 'table',
      pointer: 'word/document.xml!/w:document/w:body/w:tbl[1]',
      part: 'word/document.xml',
      path: '/w:document/w:body/w:tbl[1]',
      ordinal: 0,
      text: cellText,
      gridColumns: 1,
      rows: [
        {
          cells: [
            {
              gridSpan: 1,
              vMerge: null,
              paragraphs: [
                {
                  pointer: `word/document.xml!${cellPath}`,
                  part: 'word/document.xml',
                  path: cellPath,
                  ordinal: 0,
                  paraId: null,
                  text: cellText,
                  styleId: null,
                  styleName: null,
                  outlineLevel: null,
                  styleNameLevel: null,
                  headingStyle: false,
                  list: null,
                  commentRefs: [],
                  footnoteRefs: [],
                  endnoteRefs: [],
                },
              ],
            },
          ],
        },
      ],
    },
  ]);
}

describe('revision: fingerprint', () => {
  it('is stable for identical content and changes when content changes', () => {
    const a = threeParagraphs();
    const b = threeParagraphs();
    const c = threeParagraphs(['Alpha', 'Beta changed', 'Gamma']);

    expect(irFingerprint(b)).toBe(irFingerprint(a));
    expect(irFingerprint(c)).not.toBe(irFingerprint(a));
  });

  it('detects a format-only change even though every text is unchanged', () => {
    const plain = buildIR([para('Same words', 0)]);
    const styled = buildIR([para('Same words', 0, { styleId: 'Emphasis' })]);
    expect(irFingerprint(styled)).not.toBe(irFingerprint(plain));
  });
});

describe('revision: intent validation (trust boundary)', () => {
  it('accepts a well-formed intent', () => {
    const value = intent();
    expect(validateIntent(value)).toEqual(value);
  });

  it('rejects malformed intents', () => {
    const bad: unknown[] = [
      null,
      [],
      'text',
      intent({ id: '' }),
      intent({ text: '   ' }),
      { ...intent(), stage: 'whenever' },
      { ...intent(), kind: 'colour' },
      { ...intent(), source: 'nobody' },
      { ...intent(), targetIds: 'not-an-array' },
      { ...intent(), targetIds: [123] },
    ];
    for (const value of bad) {
      expect(() => validateIntent(value)).toThrowError(
        expect.objectContaining({ code: 'INVALID_INPUT' }),
      );
    }
  });
});

describe('revision: ledger', () => {
  it('numbers revisions monotonically and keeps a baseline', async () => {
    const ledger = createRevisionLedger();
    const first = buildIR([para('Alpha', 0)]);
    const second = buildIR([para('Alpha edited', 0)]);

    const entry1 = await ledger.record({ intent: intent({ id: 'i1' }), after: first });
    const entry2 = await ledger.record({ intent: intent({ id: 'i2' }), after: second });

    expect(entry1.revision).toBe(1);
    expect(entry2.revision).toBe(2);
    // 首条没有基线，因此没有「改动」可谈。
    expect(entry1.changes).toEqual([]);
    expect(entry1.beforeFingerprint).toBeNull();
    // 第二条与第一条做差。
    expect(entry2.changes).toHaveLength(1);
    expect(entry2.beforeFingerprint).toBe(entry1.afterFingerprint);
  });

  it('records a requirement that has not produced a new version yet', async () => {
    const ledger = createRevisionLedger();
    const entry = await ledger.record({
      intent: intent({ stage: 'requirement', kind: 'format', text: '全篇用宋体小四' }),
    });
    // 只是表达了意图：没有新版本、没有差异，但意图必须留下。
    expect(entry.producedVersion).toBe(false);
    expect(entry.changes).toEqual([]);
    expect(entry.afterFingerprint).toBeNull();
    expect((await ledger.intents())[0]?.text).toBe('全篇用宋体小四');
  });

  it('does not overwrite the baseline when no new version is recorded', async () => {
    const ledger = createRevisionLedger();
    const baseline = buildIR([para('Alpha', 0)]);
    await ledger.record({ intent: intent({ id: 'i1' }), after: baseline });
    await ledger.record({ intent: intent({ id: 'i2', stage: 'requirement', kind: 'format' }) });

    const entry3 = await ledger.record({
      intent: intent({ id: 'i3' }),
      after: buildIR([para('Alpha edited', 0)]),
    });
    // 第三条仍然与第一条（真正的上一版）做差，没有被中间的空记录清空基线。
    expect(entry3.changes).toHaveLength(1);
  });

  it('pushes each entry to the onRecord callback (real-time hook)', async () => {
    const seen: number[] = [];
    const ledger = createRevisionLedger(createMemoryRevisionStore(), {
      now: () => 1_700_000_000_000,
      onRecord: (entry) => seen.push(entry.revision),
    });
    await ledger.record({ intent: intent({ id: 'i1' }), after: buildIR([para('A', 0)]) });
    await ledger.record({ intent: intent({ id: 'i2' }), after: buildIR([para('B', 0)]) });

    expect(seen).toEqual([1, 2]);
  });

  it('uses the injected clock', async () => {
    const ledger = createRevisionLedger(createMemoryRevisionStore(), { now: () => 42 });
    const entry = await ledger.record({ intent: intent() });
    expect(entry.recordedAtMs).toBe(42);
  });

  it('returns history in revision order', async () => {
    const ledger = createRevisionLedger();
    await ledger.record({ intent: intent({ id: 'i1' }) });
    await ledger.record({ intent: intent({ id: 'i2' }) });
    const history = await ledger.history();
    expect(history.map((entry) => entry.revision)).toEqual([1, 2]);
    expect((await ledger.intents()).map((entry) => entry.id)).toEqual(['i1', 'i2']);
  });

  it('works with a fully async, externally supplied store', async () => {
    // 模拟宿主换成数据库：端口允许 Promise，账本不需任何改动。
    let entries: Awaited<ReturnType<RevisionStore['loadEntries']>> = [];
    let baseline: Awaited<ReturnType<RevisionStore['loadBaseline']>> = null;
    const store: RevisionStore = {
      async loadEntries() {
        return entries;
      },
      async appendEntry(entry) {
        entries = [...entries, entry];
      },
      async loadBaseline() {
        return baseline;
      },
      async saveBaseline(content) {
        baseline = content;
      },
    };

    const ledger = createRevisionLedger(store);
    await ledger.record({ intent: intent({ id: 'i1' }), after: buildIR([para('A', 0)]) });
    expect(await ledger.history()).toHaveLength(1);
  });

  it('keeps separate instances isolated', async () => {
    const a = createRevisionLedger();
    const b = createRevisionLedger();
    await a.record({ intent: intent({ id: 'i1' }) });
    expect(await b.history()).toHaveLength(0);
  });

  it('rejects a malformed intent before touching the store', async () => {
    const store = createMemoryRevisionStore();
    const append = vi.spyOn(store, 'appendEntry');
    const ledger = createRevisionLedger(store);

    await expect(
      ledger.record({ intent: { ...intent(), stage: 'whenever' } as never }),
    ).rejects.toMatchObject({ code: 'INVALID_INPUT' });
    // 校验发生在落账之前：脏数据进不了只追加的账本。
    expect(append).not.toHaveBeenCalled();
  });
});
