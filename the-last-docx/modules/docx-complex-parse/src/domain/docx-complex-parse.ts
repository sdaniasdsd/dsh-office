/**
 * 本模块的【领域模型】与信任边界校验。
 *
 * 这一层的存在意义只有一条：**把「引擎说什么」与「模块信什么」分开**。
 * 引擎通过 stdout 上的 JSON 提交一份「原始观察」，本文件负责校验它是否
 * 真的长成我们以为的样子，并在不符时给出本模块自己的错误分类——
 * 而不是让一个类型错误的字段一路渗透到 IR 里，最后变成一个无解的 bug。
 *
 * 三条约束：
 *   1. 这里只出现普通数据（对象/数组/数字/字符串），不存在任何引擎对象；
 *   2. 校验失败一律抛 `ENGINE_PROTOCOL_ERROR`（是引擎违约，不是文档有问题）；
 *   3. 本文件不 import 任何运行时依赖，因此可被离线测试单独覆盖。
 */
import {
  DOCX_COMPLEX_PARSE_WARNING_CODES,
  type DocxComplexParseWarningCode,
  type FloatAnchor,
  type FloatingKind,
  type PageGeometry,
  type SectionInfo,
} from '../contract';
import { DocxComplexParseError } from '../errors';

/** 引擎协议版本；与 Python 侧 `PARSE_VERSION` 必须一致。 */
export const SUPPORTED_PARSE_VERSION = 1;

/** 合并状态。`restart` 是纵向合并的起点，`continue` 是后续格。 */
export type RawVMerge = 'restart' | 'continue' | null;

/** 一个原始表格单元格。 */
export interface RawTableCell {
  /** 横向跨度；至少 1。 */
  gridSpan: number;
  vMerge: RawVMerge;
  /** 单元格纯文本。 */
  text: string;
}

/** 一个原始表格行。 */
export interface RawTableRow {
  /**
   * 本行开头跳过的网格列数（`w:gridBefore`）。
   *
   * 这是「主流实现也常漏」的属性：它让左右边缘不必落在第 0 列，
   * 直接按顺序累加 `gridSpan` 会把整行格子左移错位。
   */
  gridBefore: number;
  /** 本行结尾跳过的网格列数（`w:gridAfter`）。 */
  gridAfter: number;
  cells: RawTableCell[];
}

/** 一个原始表格块（对应一个 `w:tbl`）。 */
export interface RawTableBlock {
  kind: 'table';
  /** 部件内结构路径，例如 `/w:document/w:body/w:tbl[1]`。 */
  path: string;
  /** `w:tblGrid` 声明的逻辑列数；无法判定时为 0。 */
  tblGridColumns: number;
  /** 重复表头行数（`w:tblHeader`）。 */
  headerRowCount: number;
  rows: RawTableRow[];
}

/** 一个原始段落块。本模块只对表格做深度建模，段落保持轻量。 */
export interface RawParagraphBlock {
  kind: 'paragraph';
  path: string;
  text: string;
}

/** 正文块联合类型。 */
export type RawBlock = RawParagraphBlock | RawTableBlock;

/** 版面片段：一个正文块在页面上的位置。 */
export interface RawLayoutFragment {
  /** 正文块序号（与 `blocks` 下标一一对应）。 */
  bodyIndex: number;
  /** 物理页序号（1 基，与文档顺序一致）。 */
  physicalPage: number;
  /** 显示页序号（1 基，可能因分节重置而与物理页不同）。 */
  displayedPage: number;
  /** 页内矩形，单位为 `RawLayout.unit`。 */
  x: number;
  y: number;
  width: number;
  height: number;
}

/** 一页的尺寸。 */
export interface RawLayoutPage {
  /** 0 基物理页序号。 */
  index: number;
  width: number;
  height: number;
}

/** 版面信息。 */
export interface RawLayout {
  /** 版面模型是否可用。false 时 `pages`/`fragments` 必为空。 */
  available: boolean;
  /** 坐标单位。 */
  unit: 'pt' | 'emu' | 'twip' | 'px';
  pages: RawLayoutPage[];
  fragments: RawLayoutFragment[];
}

/** 引擎上报的告警：它说的是「结果需要被注意」，不是「解析失败」。 */
export interface RawWarning {
  code: DocxComplexParseWarningCode;
  message: string;
  path?: string;
}

/** 粗粒度格式身份；与 office-core 的 `FormatKind` 同构。 */
export type RawFormatKind = 'docx' | 'docm' | 'dotx' | 'dotm' | 'ole' | 'rtf' | 'zip' | 'unknown';

/** 容器种类。 */
export type RawContainerKind = 'zip' | 'ole' | 'rtf' | 'unknown';

/** 包级格式身份。 */
export interface RawPackageProfile {
  container: RawContainerKind;
  format: RawFormatKind;
  /** `null` 表示「无法判断」，而非「未加密」——两者对策略层含义完全不同。 */
  encrypted: boolean | null;
  /** 主文档部件的媒体类型；无法判定时为 null。 */
  mediaType: string | null;
  /** 命中的探测信号，例如 `magic:zip`。 */
  signatures: string[];
}

/** 包内一个部件。 */
export interface RawPackagePart {
  name: string;
  /** 解压后字节数。 */
  sizeBytes: number;
  /** ZIP 中压缩后的字节数。 */
  compressedSize: number;
}

/** 包内一条 OPC 关系。 */
export interface RawRelationship {
  /** 关系的来源部件；包根为 `''`。 */
  sourcePart: string;
  id: string;
  type: string;
  target: string;
  targetMode: 'Internal' | 'External';
}

/** 宏指标：只说明「存在宏」，不判定其内容或风险。 */
export interface RawMacroIndicator {
  kind: 'vba' | 'xlm';
  part: string;
  sizeBytes: number;
}

/** 外部引用指标。 */
export interface RawExternalReference {
  sourcePart: string;
  relationshipId: string;
  relationshipType: string;
  target: string;
  /** 分类标签，供策略层做差异化处理。 */
  category: string;
}

/** 嵌入对象指标。 */
export interface RawEmbeddedObject {
  part: string;
  kind: string;
  mediaType: string | null;
  sizeBytes: number;
  name: string | null;
}

/**
 * 包级指标汇总。
 *
 * 这三个数组**不是**「顺便统计的计数」：空数组等于断言「没有宏 / 没有外部引用 /
 * 没有嵌入对象」。因此它们必须来自引擎对容器的实际读取，且必须是完整枚举——
 * 任何「读到一半就停」的实现都会把安全结论变成谎言。
 */
export interface RawIndicators {
  macros: RawMacroIndicator[];
  externalReferences: RawExternalReference[];
  embeddedObjects: RawEmbeddedObject[];
}

/** 引擎提交的完整原始观察。 */
export interface RawParseResult {
  parseVersion: number;
  /** 包内主文档部件名，例如 `word/document.xml`。 */
  mainPart: string;
  /** 包级格式身份。 */
  profile: RawPackageProfile;
  /** 包内部件清单。 */
  parts: RawPackagePart[];
  relationships: RawRelationship[];
  indicators: RawIndicators;
  blocks: RawBlock[];
  warnings: RawWarning[];
  layout: RawLayout;
  /** Optional v1 extensions: omission means not observed, not observed empty. */
  breaks?: RawPageBreak[];
  sections?: RawSection[];
  floats?: RawFloat[];
}

export interface RawPageBreak {
  path: string;
  kind: 'explicit' | 'rendered' | 'section';
  beforeBodyIndex: number | null;
}

export interface RawSection {
  startBodyIndex: number | null;
  geometry: PageGeometry | null;
  columnCount: number;
  startType: SectionInfo['startType'];
}

export interface RawFloat {
  path: string;
  bodyIndex: number;
  kind: FloatingKind;
  anchor: FloatAnchor;
  width: number | null;
  height: number | null;
  unit: RawLayout['unit'] | null;
  relationshipId: string | null;
  name: string | null;
  altText: string | null;
}

/* -------------------------------------------------------------------------- */
/* 校验原语                                                                    */
/* -------------------------------------------------------------------------- */

function protocolError(message: string, details?: Record<string, unknown>): DocxComplexParseError {
  return new DocxComplexParseError('ENGINE_PROTOCOL_ERROR', message, { details });
}

function asRecord(value: unknown, where: string): Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw protocolError(`${where} must be an object`);
  }
  return value as Record<string, unknown>;
}

function asArray(value: unknown, where: string): unknown[] {
  if (!Array.isArray(value)) {
    throw protocolError(`${where} must be an array`);
  }
  return value;
}

function asString(value: unknown, where: string): string {
  if (typeof value !== 'string') {
    throw protocolError(`${where} must be a string`);
  }
  return value;
}

function asFiniteNumber(value: unknown, where: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw protocolError(`${where} must be a finite number`);
  }
  return value;
}

/** 取整数并夹到 `[min, +∞)`；非整数或越界一律视为引擎违约。 */
function asInt(value: unknown, where: string, min: number): number {
  const n = asFiniteNumber(value, where);
  if (!Number.isInteger(n) || n < min) {
    throw protocolError(`${where} must be an integer >= ${min}`, { value: n });
  }
  return n;
}

function asBoolean(value: unknown, where: string): boolean {
  if (typeof value !== 'boolean') {
    throw protocolError(`${where} must be a boolean`);
  }
  return value;
}

function asEnum<T extends string>(value: unknown, where: string, allowed: readonly T[]): T {
  if (typeof value !== 'string' || !allowed.includes(value as T)) {
    throw protocolError(`${where} must be one of ${allowed.join(', ')}`, { value });
  }
  return value as T;
}

/** 允许 null（表示「无法判断」）；缺省与 null 等价。 */
function asNullableString(value: unknown, where: string): string | null {
  return value === undefined || value === null ? null : asString(value, where);
}

function asStringArray(value: unknown, where: string): string[] {
  return asArray(value, where).map((item, i) => asString(item, `${where}[${i}]`));
}

/** 允许字段缺失并回落到默认值；存在则必须合法。 */
function optional<T>(value: unknown, fallback: T, read: (v: unknown) => T): T {
  return value === undefined || value === null ? fallback : read(value);
}

/* -------------------------------------------------------------------------- */
/* 分项校验                                                                    */
/* -------------------------------------------------------------------------- */

const VMERGE_VALUES = ['restart', 'continue'] as const;
const LAYOUT_UNITS = ['pt', 'emu', 'twip', 'px'] as const;
const CONTAINER_KINDS = ['zip', 'ole', 'rtf', 'unknown'] as const;
const FORMAT_KINDS = ['docx', 'docm', 'dotx', 'dotm', 'ole', 'rtf', 'zip', 'unknown'] as const;
const TARGET_MODES = ['Internal', 'External'] as const;
const MACRO_KINDS = ['vba', 'xlm'] as const;

function parseCell(value: unknown, where: string): RawTableCell {
  const rec = asRecord(value, where);
  return {
    gridSpan: optional(rec['gridSpan'], 1, (v) => asInt(v, `${where}.gridSpan`, 1)),
    vMerge: optional(rec['vMerge'], null, (v) => asEnum(v, `${where}.vMerge`, VMERGE_VALUES)),
    text: optional(rec['text'], '', (v) => asString(v, `${where}.text`)),
  };
}

function parseRow(value: unknown, where: string): RawTableRow {
  const rec = asRecord(value, where);
  const cells = asArray(rec['cells'], `${where}.cells`).map((cell, i) =>
    parseCell(cell, `${where}.cells[${i}]`),
  );
  return {
    gridBefore: optional(rec['gridBefore'], 0, (v) => asInt(v, `${where}.gridBefore`, 0)),
    gridAfter: optional(rec['gridAfter'], 0, (v) => asInt(v, `${where}.gridAfter`, 0)),
    cells,
  };
}

function parseBlock(value: unknown, where: string): RawBlock {
  const rec = asRecord(value, where);
  const kind = asEnum(rec['kind'], `${where}.kind`, ['paragraph', 'table'] as const);
  const path = asString(rec['path'], `${where}.path`);

  if (kind === 'paragraph') {
    return { kind, path, text: optional(rec['text'], '', (v) => asString(v, `${where}.text`)) };
  }

  const rows = asArray(rec['rows'], `${where}.rows`).map((row, i) =>
    parseRow(row, `${where}.rows[${i}]`),
  );
  return {
    kind,
    path,
    tblGridColumns: optional(rec['tblGridColumns'], 0, (v) =>
      asInt(v, `${where}.tblGridColumns`, 0),
    ),
    headerRowCount: optional(rec['headerRowCount'], 0, (v) =>
      asInt(v, `${where}.headerRowCount`, 0),
    ),
    rows,
  };
}

function parseLayout(value: unknown, where: string): RawLayout {
  const rec = asRecord(value, where);
  const available = asBoolean(rec['available'], `${where}.available`);
  const unit = optional(rec['unit'], 'pt', (v) => asEnum(v, `${where}.unit`, LAYOUT_UNITS));

  const pages = asArray(optional(rec['pages'], [], (v) => asArray(v, `${where}.pages`)), `${where}.pages`).map(
    (page, i) => {
      const p = asRecord(page, `${where}.pages[${i}]`);
      return {
        index: asInt(p['index'], `${where}.pages[${i}].index`, 0),
        width: asFiniteNumber(p['width'], `${where}.pages[${i}].width`),
        height: asFiniteNumber(p['height'], `${where}.pages[${i}].height`),
      };
    },
  );

  const fragments = asArray(
    optional(rec['fragments'], [], (v) => asArray(v, `${where}.fragments`)),
    `${where}.fragments`,
  ).map((fragment, i) => {
    const f = asRecord(fragment, `${where}.fragments[${i}]`);
    const at = `${where}.fragments[${i}]`;
    return {
      bodyIndex: asInt(f['bodyIndex'], `${at}.bodyIndex`, 0),
      physicalPage: asInt(f['physicalPage'], `${at}.physicalPage`, 1),
      displayedPage: asInt(f['displayedPage'], `${at}.displayedPage`, 1),
      x: asFiniteNumber(f['x'], `${at}.x`),
      y: asFiniteNumber(f['y'], `${at}.y`),
      width: asFiniteNumber(f['width'], `${at}.width`),
      height: asFiniteNumber(f['height'], `${at}.height`),
    };
  });

  // 版面可用却没有任何片段，说明引擎自称有版面但没给数据——宁可当成不可用。
  const usable = available && fragments.length > 0;
  return { available: usable, unit, pages: usable ? pages : [], fragments: usable ? fragments : [] };
}

function bodyIndex(value: unknown, where: string, count: number): number {
  const index = asInt(value, where, 0);
  if (index >= count) throw protocolError(`${where} references a missing body block`);
  return index;
}

function nullableNumber(value: unknown, where: string): number | null {
  return value === null ? null : asFiniteNumber(value, where);
}

function dimension(value: unknown, where: string): number | null {
  const result = nullableNumber(value, where);
  if (result !== null && result < 0) throw protocolError(`${where} must be nonnegative`);
  return result;
}

function parseGeometry(value: unknown, where: string): PageGeometry | null {
  if (value === null) return null;
  const rec = asRecord(value, where);
  const margins = asRecord(rec['margins'], `${where}.margins`);
  const width = asFiniteNumber(rec['width'], `${where}.width`);
  const height = asFiniteNumber(rec['height'], `${where}.height`);
  if (width <= 0 || height <= 0) throw protocolError(`${where} needs positive dimensions`);
  return {
    width, height,
    unit: asEnum(rec['unit'], `${where}.unit`, LAYOUT_UNITS),
    orientation: asEnum(rec['orientation'], `${where}.orientation`, ['portrait', 'landscape']),
    margins: {
      top: asFiniteNumber(margins['top'], `${where}.margins.top`),
      right: asFiniteNumber(margins['right'], `${where}.margins.right`),
      bottom: asFiniteNumber(margins['bottom'], `${where}.margins.bottom`),
      left: asFiniteNumber(margins['left'], `${where}.margins.left`),
    },
  };
}

function parseObservations(root: Record<string, unknown>, count: number):
  Pick<RawParseResult, 'breaks' | 'sections' | 'floats'> {
  const result: Pick<RawParseResult, 'breaks' | 'sections' | 'floats'> = {};
  if (root['breaks'] !== undefined) {
    result.breaks = asArray(root['breaks'], 'breaks').map((value, index) => {
      const at = `breaks[${index}]`;
      const rec = asRecord(value, at);
      return {
        path: asString(rec['path'], `${at}.path`),
        kind: asEnum(rec['kind'], `${at}.kind`, ['explicit', 'rendered', 'section']),
        beforeBodyIndex: rec['beforeBodyIndex'] === null ? null : bodyIndex(rec['beforeBodyIndex'], at, count),
      };
    });
  }
  if (root['sections'] !== undefined) {
    result.sections = asArray(root['sections'], 'sections').map((value, index) => {
      const at = `sections[${index}]`;
      const rec = asRecord(value, at);
      return {
        startBodyIndex: rec['startBodyIndex'] === null ? null : bodyIndex(rec['startBodyIndex'], at, count),
        geometry: parseGeometry(rec['geometry'], `${at}.geometry`),
        columnCount: asInt(rec['columnCount'], `${at}.columnCount`, 1),
        startType: asEnum(rec['startType'], `${at}.startType`, ['continuous', 'nextPage', 'evenPage', 'oddPage', 'unknown']),
      };
    });
  }
  if (root['floats'] !== undefined) {
    result.floats = asArray(root['floats'], 'floats').map((value, index): RawFloat => {
      const at = `floats[${index}]`;
      const rec = asRecord(value, at);
      const anchor = asRecord(rec['anchor'], `${at}.anchor`);
      return {
        path: asString(rec['path'], `${at}.path`),
        bodyIndex: bodyIndex(rec['bodyIndex'], at, count),
        kind: asEnum(rec['kind'], `${at}.kind`, ['image', 'chart', 'shape', 'textbox', 'equation', 'ole', 'unknown']),
        anchor: {
          relativeFromHorizontal: asString(anchor['relativeFromHorizontal'], `${at}.relativeFromHorizontal`),
          relativeFromVertical: asString(anchor['relativeFromVertical'], `${at}.relativeFromVertical`),
          alignHorizontal: anchor['alignHorizontal'] === null ? null : asEnum(anchor['alignHorizontal'], at, ['left', 'center', 'right'] as const),
          alignVertical: anchor['alignVertical'] === null ? null : asEnum(anchor['alignVertical'], at, ['top', 'center', 'bottom'] as const),
          offsetX: nullableNumber(anchor['offsetX'], `${at}.offsetX`),
          offsetY: nullableNumber(anchor['offsetY'], `${at}.offsetY`),
          unit: anchor['unit'] === null ? null : asEnum(anchor['unit'], at, LAYOUT_UNITS),
          wrap: asEnum(anchor['wrap'], `${at}.wrap`, ['inline', 'square', 'tight', 'through', 'topAndBottom', 'behindText', 'inFrontOfText', 'none']),
          behindText: asBoolean(anchor['behindText'], `${at}.behindText`),
        },
        width: dimension(rec['width'], `${at}.width`), height: dimension(rec['height'], `${at}.height`),
        unit: rec['unit'] === null ? null : asEnum(rec['unit'], at, LAYOUT_UNITS),
        relationshipId: asNullableString(rec['relationshipId'], `${at}.relationshipId`),
        name: asNullableString(rec['name'], `${at}.name`),
        altText: asNullableString(rec['altText'], `${at}.altText`),
      };
    });
  }
  return result;
}

/* -------------------------------------------------------------------------- */
/* 包级事实                                                                    */
/* -------------------------------------------------------------------------- */

function parsePackageProfile(value: unknown, where: string): RawPackageProfile {
  const rec = asRecord(value, where);
  return {
    container: asEnum(rec['container'], `${where}.container`, CONTAINER_KINDS),
    format: asEnum(rec['format'], `${where}.format`, FORMAT_KINDS),
    // 这里刻意接受 null 而不是回落到 false：`encrypted: null` 是「无法判断」，
    // 把它写成 false 会让策略层以为「已确认未加密」。
    encrypted: rec['encrypted'] === undefined ? null : asBooleanOrNull(rec['encrypted'], `${where}.encrypted`),
    mediaType: asNullableString(rec['mediaType'], `${where}.mediaType`),
    signatures: optional(rec['signatures'], [], (v) => asStringArray(v, `${where}.signatures`)),
  };
}

function asBooleanOrNull(value: unknown, where: string): boolean | null {
  return value === null ? null : asBoolean(value, where);
}

function parsePackagePart(value: unknown, where: string): RawPackagePart {
  const rec = asRecord(value, where);
  return {
    name: asString(rec['name'], `${where}.name`),
    sizeBytes: asInt(rec['sizeBytes'], `${where}.sizeBytes`, 0),
    compressedSize: asInt(rec['compressedSize'], `${where}.compressedSize`, 0),
  };
}

function parseRelationship(value: unknown, where: string): RawRelationship {
  const rec = asRecord(value, where);
  return {
    // 来源部件允许为空字符串（包根 `_rels/.rels` 的关系就来自包根）。
    sourcePart: optional(rec['sourcePart'], '', (v) => asString(v, `${where}.sourcePart`)),
    id: asString(rec['id'], `${where}.id`),
    type: asString(rec['type'], `${where}.type`),
    target: asString(rec['target'], `${where}.target`),
    targetMode: optional(rec['targetMode'], 'Internal', (v) =>
      asEnum(v, `${where}.targetMode`, TARGET_MODES),
    ),
  };
}

function parseIndicators(value: unknown, where: string): RawIndicators {
  const rec = asRecord(value, where);
  return {
    // 三个数组都**必须存在**：缺一个就等于少断言一类安全指标。
    macros: asArray(rec['macros'], `${where}.macros`).map((item, i) => {
      const entry = asRecord(item, `${where}.macros[${i}]`);
      return {
        kind: asEnum(entry['kind'], `${where}.macros[${i}].kind`, MACRO_KINDS),
        part: asString(entry['part'], `${where}.macros[${i}].part`),
        sizeBytes: asInt(entry['sizeBytes'], `${where}.macros[${i}].sizeBytes`, 0),
      };
    }),
    externalReferences: asArray(rec['externalReferences'], `${where}.externalReferences`).map(
      (item, i) => {
        const entry = asRecord(item, `${where}.externalReferences[${i}]`);
        const at = `${where}.externalReferences[${i}]`;
        return {
          sourcePart: asString(entry['sourcePart'], `${at}.sourcePart`),
          relationshipId: asString(entry['relationshipId'], `${at}.relationshipId`),
          relationshipType: asString(entry['relationshipType'], `${at}.relationshipType`),
          target: asString(entry['target'], `${at}.target`),
          category: asString(entry['category'], `${at}.category`),
        };
      },
    ),
    embeddedObjects: asArray(rec['embeddedObjects'], `${where}.embeddedObjects`).map((item, i) => {
      const entry = asRecord(item, `${where}.embeddedObjects[${i}]`);
      const at = `${where}.embeddedObjects[${i}]`;
      return {
        part: asString(entry['part'], `${at}.part`),
        kind: asString(entry['kind'], `${at}.kind`),
        mediaType: asNullableString(entry['mediaType'], `${at}.mediaType`),
        sizeBytes: asInt(entry['sizeBytes'], `${at}.sizeBytes`, 0),
        name: asNullableString(entry['name'], `${at}.name`),
      };
    }),
  };
}

/** 判断告警码是否在本模块固定的表内。 */
export function isKnownWarningCode(code: string): code is DocxComplexParseWarningCode {
  return Object.prototype.hasOwnProperty.call(DOCX_COMPLEX_PARSE_WARNING_CODES, code);
}

function parseWarning(value: unknown, where: string): RawWarning {
  const rec = asRecord(value, where);
  const code = asString(rec['code'], `${where}.code`);
  // 告警码必须来自本模块固定的表：引擎不能自造一个上层读不懂的码。
  if (!isKnownWarningCode(code)) {
    throw protocolError(`unknown warning code "${code}"`, { code });
  }
  const message = asString(rec['message'], `${where}.message`);
  const path = rec['path'];
  return path === undefined || path === null
    ? { code, message }
    : { code, message, path: asString(path, `${where}.path`) };
}

/* -------------------------------------------------------------------------- */
/* 失败信封                                                                    */
/* -------------------------------------------------------------------------- */

/**
 * 引擎上报的失败信封。
 *
 * 刻意不叫「错误码」：引擎说的是【原因】（`reason`），错误码由 TS 侧决定。
 * 这样换一个引擎时，上层看到的错误分类体系完全不变。
 */
export interface RawEngineFailure {
  reason: string;
  message: string;
  detail: Record<string, unknown>;
}

/**
 * 识别引擎的失败信封；返回 null 表示「这不是失败信封」，调用方应继续按成功结果解析。
 *
 * @throws DocxComplexParseError 信封存在但结构不合法时抛出 `ENGINE_PROTOCOL_ERROR`。
 */
export function parseEngineFailure(value: unknown): RawEngineFailure | null {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  if (!('failure' in record)) return null;

  const failure = asRecord(record['failure'], 'failure');
  const detail = failure['detail'];
  return {
    reason: asString(failure['reason'], 'failure.reason'),
    message: asString(failure['message'], 'failure.message'),
    detail: detail === undefined || detail === null ? {} : asRecord(detail, 'failure.detail'),
  };
}

/* -------------------------------------------------------------------------- */
/* 入口                                                                        */
/* -------------------------------------------------------------------------- */

/**
 * 校验并归一化引擎输出。
 *
 * @throws DocxComplexParseError 编码为 `ENGINE_PROTOCOL_ERROR`。
 */
export function parseRawResult(value: unknown): RawParseResult {
  const root = asRecord(value, 'engine output');

  const parseVersion = asInt(root['parseVersion'], 'parseVersion', 1);
  if (parseVersion !== SUPPORTED_PARSE_VERSION) {
    throw protocolError(
      `unsupported engine parse version ${parseVersion}; expected ${SUPPORTED_PARSE_VERSION}`,
      { parseVersion },
    );
  }

  const blocks = asArray(root['blocks'], 'blocks').map((block, i) =>
    parseBlock(block, `blocks[${i}]`),
  );

  const layout = parseLayout(root['layout'], 'layout');
  for (const fragment of layout.fragments) bodyIndex(fragment.bodyIndex, 'layout fragment', blocks.length);

  return {
    parseVersion,
    mainPart: asString(root['mainPart'], 'mainPart'),
    // 包级事实是必填而不是可选：缺失意味着引擎没有真正读过容器，
    // 此时 `indicators` 之类的安全结论无从谈起，必须按违约处理。
    profile: parsePackageProfile(root['profile'], 'profile'),
    parts: asArray(root['parts'], 'parts').map((part, i) => parsePackagePart(part, `parts[${i}]`)),
    relationships: asArray(root['relationships'], 'relationships').map((item, i) =>
      parseRelationship(item, `relationships[${i}]`),
    ),
    indicators: parseIndicators(root['indicators'], 'indicators'),
    blocks,
    warnings: asArray(optional(root['warnings'], [], (v) => asArray(v, 'warnings')), 'warnings').map(
      (warning, i) => parseWarning(warning, `warnings[${i}]`),
    ),
    layout,
    ...parseObservations(root, blocks.length),
  };
}
