/**
 * 深度引擎（可选 Docling 后端）的契约测试。
 *
 * 这里守护的是一条硬性路由不变量：**`inspect` 恒走轻量引擎**。
 * 这不是风格问题而是定位问题——`inspect` 的语义是「快速看一眼这是什么产物」，
 * 让它付出加载版面模型（冷启动以秒计）的代价会直接违背该接口的承诺。
 *
 * 除 gated 的集成用例外，本文件不需要 Python，也绝不需要 Docling：
 * 路由判定用可计数的假引擎，因此可在任何环境验证。
 */
import { describe, expect, it } from 'vitest';

import { DocxParseError, resolveConfig } from '../src/index';
// 具体引擎与其脚本路径解析是内部接缝；公开的只有 DocxEngine 端口本身。
import { DoclingEngine, createDeepEngine, defaultDoclingScriptPath } from '../src/internal';
import type { Block, ConfigOverrides, DocxEngine, ParseResult } from '../src/index';
import { configSchema } from '../src/config';
import {
  artifactRef,
  countingEngine,
  doclingAvailable,
  fakeEngine,
  fakeParseResult,
  fixturePath,
  makeModule,
  makeModuleWithEngine,
  makeModuleWithEngines,
  PROJECT_ROOT,
  pythonAvailable,
  pythonExecutable,
} from './support';

/* -------------------------------------------------------------------------- */
/* 构造测试数据                                                                 */
/* -------------------------------------------------------------------------- */

/** 构造一个段落块。 */
function para(index: number, text: string, headingLevel: number | null = null): Block {
  return {
    kind: 'paragraph',
    paragraph: {
      index,
      text,
      styleId: headingLevel === null ? null : `Heading${headingLevel}`,
      styleName: headingLevel === null ? null : `heading ${headingLevel}`,
      outlineLevel: headingLevel === null ? null : headingLevel - 1,
      headingLevel,
      alignment: null,
      listItem: false,
      runs: [{ text, bold: false, italic: false, underline: false }],
    },
  };
}

/** 一份足以走完流水线的假结果（含一个标题）。 */
function titledResult(): ParseResult {
  return fakeParseResult({ blocks: [para(0, 'Title', 1), para(1, 'Body')] });
}

/** 具名假引擎：路由测试必须能区分「谁跑的」。 */
function namedEngine(name: string): DocxEngine {
  return { ...fakeEngine(titledResult()), name };
}

/** 启用深度引擎的配置覆盖项。 */
const DEEP_OVERRIDE = { engine: { deep: { driver: 'docling' as const } } };

/** 装配一个双引擎模块，两个引擎都可计数且名字不同。 */
function moduleWithBothEngines() {
  const light = countingEngine(namedEngine('fake-light'));
  const deep = countingEngine(namedEngine('fake-deep'));
  const module = makeModuleWithEngines(light.engine, deep.engine, DEEP_OVERRIDE);
  return { module, light, deep };
}

/* -------------------------------------------------------------------------- */
/* 路由隔离                                                                     */
/* -------------------------------------------------------------------------- */

describe('deep engine: routing isolation', () => {
  it('never lets inspect touch the deep engine', async () => {
    const { module, light, deep } = moduleWithBothEngines();

    await module.handlers.inspect({
      artifactRef: artifactRef('clean.docx'),
      operation: 'inspect',
      requestId: 'req-inspect',
    });

    expect(light.calls()).toBe(1);
    expect(deep.calls()).toBe(0);
  });

  it('routes execute to the deep engine when one is configured', async () => {
    const { module, light, deep } = moduleWithBothEngines();

    await module.handlers.execute({
      artifactRef: artifactRef('clean.docx'),
      operation: 'execute',
      requestId: 'req-execute',
    });

    expect(light.calls()).toBe(0);
    expect(deep.calls()).toBe(1);
  });

  it('routes verify to the deep engine when one is configured', async () => {
    const { module, light, deep } = moduleWithBothEngines();

    await module.handlers.verify({
      artifactRef: artifactRef('clean.docx'),
      operation: 'verify',
      requestId: 'req-verify',
      policy: { id: 'deep-engine-test', allowEncrypted: false },
    });

    expect(light.calls()).toBe(0);
    expect(deep.calls()).toBe(1);
  });

  it('keeps using the lightweight engine when nothing configures a deep engine', async () => {
    // 关键回归：既没有 `engine.deep` 配置、也没有注入深度引擎时，
    // 路由必须停在轻量引擎上——不得因为「深度引擎是可选能力」就默认构造一个。
    const light = countingEngine(namedEngine('fake-light'));
    const module = makeModuleWithEngine(light.engine);

    await module.handlers.execute({
      artifactRef: artifactRef('clean.docx'),
      operation: 'execute',
      requestId: 'req-fallback',
    });

    expect(light.calls()).toBe(1);
  });

  it('records the engine that actually ran in the telemetry summary', async () => {
    const { module } = moduleWithBothEngines();

    const execute = await module.handlers.execute({
      artifactRef: artifactRef('clean.docx'),
      operation: 'execute',
      requestId: 'req-telemetry',
    });
    const inspect = await module.handlers.inspect({
      artifactRef: artifactRef('clean.docx'),
      operation: 'inspect',
      requestId: 'req-telemetry-2',
    });

    // 同一份文档在两个接口上走了不同引擎，摘要必须能解释这件事。
    expect(execute.telemetry?.engine).toBe('fake-deep');
    expect(inspect.telemetry?.engine).toBe('fake-light');
  });
});

/* -------------------------------------------------------------------------- */
/* 深度引擎配置校验                                                              */
/* -------------------------------------------------------------------------- */

describe('deep engine: configuration', () => {
  it('rejects an unsupported deep driver', () => {
    let caught: unknown;
    try {
      resolveConfig({ engine: { deep: { driver: 'marker' as never } } });
    } catch (error) {
      caught = error;
    }
    expect((caught as DocxParseError).code).toBe('INVALID_INPUT');
  });

  it('rejects a non-positive or fractional deep timeout', () => {
    for (const timeoutMs of [0, -1, 1.5]) {
      expect(() =>
        resolveConfig({ engine: { deep: { driver: 'docling', timeoutMs } } }),
      ).toThrow(DocxParseError);
    }
  });

  it('accepts a positive integer deep timeout', () => {
    const config = resolveConfig({ engine: { deep: { driver: 'docling', timeoutMs: 5000 } } });
    expect(config.engine.deep?.timeoutMs).toBe(5000);
  });

  it('freezes the nested deep configuration', () => {
    const config = resolveConfig(DEEP_OVERRIDE);
    expect(Object.isFrozen(config.engine)).toBe(true);
    expect(Object.isFrozen(config.engine.deep)).toBe(true);
  });

  it('exposes deep in the config schema', () => {
    const parsed = JSON.parse(JSON.stringify(configSchema())) as {
      properties: {
        engine: {
          properties: {
            deep: {
              type: string;
              required: string[];
              properties: { driver: { enum: string[] } };
            };
          };
        };
      };
    };
    const deep = parsed.properties.engine.properties.deep;
    expect(deep.type).toBe('object');
    expect(deep.required).toEqual(['driver']);
    expect(deep.properties.driver.enum).toEqual(['docling']);
  });

  it('constructs no deep engine when the config omits it', () => {
    expect(createDeepEngine(resolveConfig({}).engine)).toBeNull();
  });

  it('constructs a Docling engine when the config requests it', () => {
    const engine = createDeepEngine(resolveConfig(DEEP_OVERRIDE).engine);
    expect(engine).toBeInstanceOf(DoclingEngine);
    expect(engine?.name).toBe('python-docling-parse');
  });

  it('refuses to build a Docling engine without a deep configuration', () => {
    expect(() => new DoclingEngine(resolveConfig({}).engine)).toThrow(DocxParseError);
  });

  it('resolves the bridge script next to the lightweight engine', () => {
    expect(defaultDoclingScriptPath().endsWith('docx_parse_docling.py')).toBe(true);
  });
});

/* -------------------------------------------------------------------------- */
/* 标题合并逻辑（用 docling 桩模块覆盖，无需安装真实依赖）                          */
/* -------------------------------------------------------------------------- */

/** 桩模块所在的目录；通过 PYTHONPATH 注入子进程。 */
const STUB_ROOT = `${PROJECT_ROOT}/tests/stubs`;

/** 构造一个使用 docling 桩模块的深度引擎配置。 */
function stubOverride(spec: Record<string, unknown>): ConfigOverrides {
  return {
    engine: {
      deep: {
        driver: 'docling',
        pythonPath: pythonExecutable(),
        env: {
          PYTHONPATH: STUB_ROOT,
          DOCX_PARSE_STUB_DOCLING: JSON.stringify(spec),
        },
      },
    },
  };
}

/** 只取段落记录，便于按文本定位。 */
function paragraphRecords(blocks: Block[]) {
  return blocks.flatMap((block) => (block.kind === 'paragraph' ? [block.paragraph] : []));
}

/** 用真实轻量引擎跑一遍，作为「OOXML 自身的判定」基线。 */
async function executeLight() {
  const module = makeModule();
  const output = await module.handlers.execute({
    artifactRef: artifactRef('clean.docx'),
    operation: 'execute',
    requestId: 'req-light',
  });
  await module.dispose();
  return output;
}

/** 用桩驱动的深度引擎跑一遍。 */
async function executeWithStub(spec: Record<string, unknown>) {
  const override = stubOverride(spec);
  const config = resolveConfig(override);
  const module = makeModuleWithEngines(
    namedEngine('fake-light'),
    createDeepEngine(config.engine)!,
    override,
  );
  const output = await module.handlers.execute({
    artifactRef: artifactRef('clean.docx'),
    operation: 'execute',
    requestId: 'req-stub',
  });
  await module.dispose();
  return output;
}

describe.skipIf(!pythonAvailable())('deep engine: heading merge (docling stub)', () => {
  it('augments headings that the OOXML styles did not mark', async () => {
    const light = paragraphRecords((await executeLight()).result.ir.blocks);
    const target = light.find((p) => p.headingLevel === null && p.text.trim() !== '');
    expect(target, 'fixture must contain a plain paragraph').toBeDefined();

    const deep = paragraphRecords(
      (
        await executeWithStub({
          items: [{ label: 'section_header', text: target!.text, level: 0 }],
        })
      ).result.ir.blocks,
    );

    const merged = deep.find((p) => p.text === target!.text)!;
    // 桩给的层级 0 → 第 1 级标题；outlineLevel 是 0 基。
    expect(merged.headingLevel).toBe(1);
    expect(merged.outlineLevel).toBe(0);
  });

  it('never overrides a heading decision that OOXML already made', async () => {
    const light = paragraphRecords((await executeLight()).result.ir.blocks);
    const existing = light.find((p) => p.headingLevel !== null);
    expect(existing, 'fixture must contain a styled heading').toBeDefined();

    const deep = paragraphRecords(
      (
        await executeWithStub({
          // 桩声称同一段文本是第 6 级；OOXML 的权威判定必须胜出。
          items: [{ label: 'section_header', text: existing!.text, level: 5 }],
        })
      ).result.ir.blocks,
    );

    const merged = deep.find((p) => p.text === existing!.text)!;
    expect(merged.headingLevel).toBe(existing!.headingLevel);
    expect(merged.outlineLevel).toBe(existing!.outlineLevel);
  });

  it('clamps a very deep Docling level to the OOXML maximum', async () => {
    const light = paragraphRecords((await executeLight()).result.ir.blocks);
    const target = light.find((p) => p.headingLevel === null && p.text.trim() !== '')!;

    const deep = paragraphRecords(
      (
        await executeWithStub({ items: [{ label: 'title', text: target.text, level: 99 }] })
      ).result.ir.blocks,
    );

    // OOXML 只定义到 Heading 9，但 6 级以上已无区分度；越界必须收敛而不是透传。
    expect(deep.find((p) => p.text === target.text)!.headingLevel).toBe(6);
  });

  it('tolerates an iterator that yields bare items instead of tuples', async () => {
    const light = paragraphRecords((await executeLight()).result.ir.blocks);
    const target = light.find((p) => p.headingLevel === null && p.text.trim() !== '')!;

    const deep = paragraphRecords(
      (
        await executeWithStub({
          bare: true,
          items: [{ label: 'title', text: target.text, level: 3 }],
        })
      ).result.ir.blocks,
    );

    // 缺少层级信息时按第 1 级处理，而不是崩掉或整段漏掉。
    expect(deep.find((p) => p.text === target.text)!.headingLevel).toBe(1);
  });

  it('falls back to get_text() when the item has no text attribute', async () => {
    const light = paragraphRecords((await executeLight()).result.ir.blocks);
    const target = light.find((p) => p.headingLevel === null && p.text.trim() !== '')!;

    const deep = paragraphRecords(
      (
        await executeWithStub({
          items: [{ label: 'section_header', text: target.text, level: 1, textVia: 'get_text' }],
        })
      ).result.ir.blocks,
    );

    expect(deep.find((p) => p.text === target.text)!.headingLevel).toBe(2);
  });

  it('leaves blocks untouched when Docling reports no headings', async () => {
    const light = (await executeLight()).result.ir.blocks;
    const deep = (
      await executeWithStub({ items: [{ label: 'paragraph', text: 'ignored', level: 0 }] })
    ).result.ir.blocks;

    // 无标题文档是正常情况：合并必须是彻底的空操作。
    expect(deep).toEqual(light);
  });

  it('degrades to the lightweight blocks when Docling raises', async () => {
    const output = await executeWithStub({ raise: 'stub exploded' });

    // 深度引擎出错不能让整次调用失败——结果必须仍然完整。
    expect(output.result.ir.blocks.length).toBeGreaterThan(0);

    // 但降级必须留下可诊断的痕迹。
    const issueCodes = output.warnings.map((warning) => warning.details?.['issueCode']);
    expect(issueCodes).toContain('DOCLING_FAILED');
  });
});

/* -------------------------------------------------------------------------- */
/* 集成（gated）                                                                */
/* -------------------------------------------------------------------------- */

/**
 * 深度脚本不可用时的失败必须可诊断。
 *
 * 之所以用「不存在的脚本」而不是「没装 Docling」：前者与后者走的是同一条
 * 子进程失败路径（非零退出码 → ENGINE_FAILED），但不需要真的卸载依赖就能复现。
 */
describe.skipIf(!pythonAvailable())('deep engine: unusable backend', () => {
  it('fails with ENGINE_FAILED instead of silently degrading', async () => {
    const config = resolveConfig({
      engine: { deep: { driver: 'docling', scriptPath: fixturePath('missing-bridge.py') } },
    });
    const broken = createDeepEngine(config.engine);
    expect(broken).toBeInstanceOf(DoclingEngine);

    const module = makeModuleWithEngines(
      namedEngine('fake-light'),
      broken!,
      {
        engine: { deep: { driver: 'docling', scriptPath: fixturePath('missing-bridge.py') } },
      },
    );

    await expect(
      module.handlers.execute({
        artifactRef: artifactRef('clean.docx'),
        operation: 'execute',
        requestId: 'req-docling-fail',
      }),
    ).rejects.toMatchObject({ code: 'ENGINE_FAILED' });
  });
});

/**
 * Docling 缺失时的降级。
 *
 * 这是深度引擎最重要的一条安全性质：**开启它绝不能让本该解析成功的文档失败**。
 * 桥接脚本在 Docling 不可导入时回退标准库正文解析，并留下一条可诊断的告警——
 * 降级必须可见，不能静默。
 */
describe.skipIf(!pythonAvailable() || doclingAvailable())(
  'deep engine: graceful degradation without Docling',
  () => {
    it('still returns a full result and reports why it degraded', async () => {
      const config = resolveConfig(DEEP_OVERRIDE);
      const module = makeModuleWithEngines(
        namedEngine('fake-light'),
        createDeepEngine(config.engine)!,
        DEEP_OVERRIDE,
      );

      const output = await module.handlers.execute({
        artifactRef: artifactRef('clean.docx'),
        operation: 'execute',
        requestId: 'req-docling-fallback',
      });

      // 结果本身完整：这正是「回退而非失败」的含义。
      expect(output.result.ir.blocks.length).toBeGreaterThan(0);
      expect(output.result.ir.styles.length).toBeGreaterThan(0);

      // 降级必须留下痕迹。
      const issueCodes = output.warnings.map((warning) => warning.details?.['issueCode']);
      expect(issueCodes).toContain('DOCLING_UNAVAILABLE');

      await module.dispose();
    });
  },
);

describe.skipIf(!(pythonAvailable() && doclingAvailable()))(
  'deep engine: Docling integration',
  () => {
    it('parses through Docling on execute while inspect stays on the lightweight path', async () => {
      // 深度引擎必须用「装了 Docling 的那个解释器」，而不是 PATH 上的默认 python；
      // 探测与执行共用 pythonExecutable()，两者不会脱节。
      const override = {
        engine: { deep: { driver: 'docling' as const, pythonPath: pythonExecutable() } },
      };
      const config = resolveConfig(override);
      const module = makeModuleWithEngines(
        namedEngine('fake-light'),
        createDeepEngine(config.engine)!,
        override,
      );

      const execute = await module.handlers.execute({
        artifactRef: artifactRef('clean.docx'),
        operation: 'execute',
        requestId: 'req-docling',
      });

      // 引擎名必须可区分，否则无法解释两条路径的耗时差异。
      expect(execute.telemetry?.engine).toBe('python-docling-parse');
      expect(execute.result.ir.blocks.length).toBeGreaterThan(0);
      // 深度引擎接管正文后，OPC 层（样式/关系/批注/脚注）不得退化。
      expect(execute.result.ir.styles.length).toBeGreaterThan(0);

      await module.dispose();
    });
  },
);
