/**
 * 引擎适配器 —— 本模块中【唯一】知道「解析是 spawn 一个 Python 进程跑 rdocx」
 * 这一事实的文件。
 *
 * 三条设计纪律（与兄弟模块同构，换引擎时上层无感）：
 *   1. 适配器只返回 `RawParseResult` 这种普通领域对象；子进程句柄、流、临时文件
 *      一律不得逃出本文件；
 *   2. 引擎通过 stdout 上的 JSON 通信，因此换实现（进程内 TS 读取器、Open XML SDK
 *      服务、远程解析器）只需替换 `DocxEngine` 的一个实现；
 *   3. 引擎失败会被翻译成本模块的错误分类体系——**引擎只报原因，错误码在此决定**。
 */
import { spawn } from 'node:child_process';
import { statSync } from 'node:fs';
import { dirname, isAbsolute, join, resolve as resolvePath } from 'node:path';
import { fileURLToPath } from 'node:url';

import {
  IMPLEMENTED_ENGINE_DRIVERS,
  type DocxComplexParseErrorCode,
  type EngineConfig,
  type FeatureFlags,
  type LimitConfig,
  type TelemetrySink,
} from '../contract';
import {
  parseEngineFailure,
  parseRawResult,
  type RawParseResult,
} from '../domain/docx-complex-parse';
import { DocxComplexParseError, isDocxComplexParseError } from '../errors';

/**
 * 从子进程流中保留的字节上限。
 *
 * 复杂文档的载荷比普通文档大得多（网格会按行列展开），但仍必须设上限：
 * 一个异常文档不能让 Node 堆内存暴涨。64 MiB 远超正常文档，又足以兜住失控输出。
 */
const MAX_ENGINE_OUTPUT_BYTES = 64 * 1024 * 1024;

/** 默认解析器明确不支持的 URI scheme；遇到时应提示调用方注入 office-files 物化器。 */
const UNSUPPORTED_SCHEMES = ['http:', 'https:', 'memory:', 'data:', 's3:', 'gs:'];

/** 一次解析请求所需的全部输入。 */
export interface ParseRequest {
  /** 已解析为本地可读路径的 artifact。 */
  artifactPath: string;
  limits: LimitConfig;
  featureFlags: FeatureFlags;
  timeoutMs: number;
}

/**
 * 所有引擎后端都必须实现的端口。
 *
 * 刻意保持极窄：只回答一个有界、无副作用的问题——「这份文档里有什么原始结构？」
 */
export interface DocxEngine {
  /** 引擎名，用于遥测与诊断。 */
  readonly name: string;
  parse(request: ParseRequest): Promise<RawParseResult>;
  /** 释放资源（Python 桥接为无状态，因此是空实现）。 */
  dispose(): Promise<void>;
}

/** 内置解析脚本的绝对路径。 */
export function defaultScriptPath(): string {
  // 用 import.meta.url 而非 __dirname：本模块以 ESM 发布，且需要跨平台正确解析。
  return join(dirname(fileURLToPath(import.meta.url)), 'docx_complex_parse.py');
}

/**
 * 把 `ArtifactRef.uri` 解析为本地可读路径。
 *
 * 只接受独立模块确实读得懂的形态；其余一律以 `UNSUPPORTED_ARTIFACT_URI` 拒绝，
 * 让 Profile 有机会注入 `office-files` 物化器，而不是「默默读错东西」。
 */
export function resolveArtifactPath(uri: string): string {
  if (typeof uri !== 'string' || uri.trim() === '') {
    throw new DocxComplexParseError('INVALID_INPUT', 'artifactRef.uri must be a non-empty string');
  }

  const lowered = uri.toLowerCase();
  if (UNSUPPORTED_SCHEMES.some((scheme) => lowered.startsWith(scheme))) {
    throw new DocxComplexParseError(
      'UNSUPPORTED_ARTIFACT_URI',
      'The default resolver only reads local files; inject an office-files materializer for other schemes.',
      { details: { uri } },
    );
  }

  if (lowered.startsWith('file://')) {
    try {
      return fileURLToPath(uri);
    } catch (error) {
      throw new DocxComplexParseError('UNSUPPORTED_ARTIFACT_URI', 'Malformed file:// URI', {
        details: { uri },
        cause: error,
      });
    }
  }

  // 出现任意 scheme 但不在白名单内（例如 `ftp://`）。
  if (uri.includes('://')) {
    throw new DocxComplexParseError('UNSUPPORTED_ARTIFACT_URI', 'Unsupported URI scheme', {
      details: { uri },
    });
  }

  // 相对路径按进程工作目录解析；绝对路径直接使用。
  return isAbsolute(uri) ? uri : resolvePath(process.cwd(), uri);
}

/**
 * 引擎上报的原因 → 本模块错误码。
 *
 * 这是「引擎无权选择错误码」这条纪律的落点：映射集中在这里，一眼能看全。
 */
const FAILURE_CODE_BY_REASON: Record<string, DocxComplexParseErrorCode> = {
  not_found: 'ARTIFACT_NOT_FOUND',
  format_mismatch: 'FORMAT_MISMATCH',
  encrypted: 'UNSUPPORTED_CONTAINER',
  parse_failed: 'PARSE_FAILED',
  limit_exceeded: 'LIMIT_EXCEEDED',
  layout_unavailable: 'LAYOUT_UNAVAILABLE',
  runtime_missing: 'ENGINE_UNAVAILABLE',
  bad_config: 'INVALID_INPUT',
};

/** 未知原因一律落到 `ENGINE_FAILED`——宁可笼统，也不要凭空发明一个码。 */
function codeForReason(reason: string): DocxComplexParseErrorCode {
  return FAILURE_CODE_BY_REASON[reason] ?? 'ENGINE_FAILED';
}

/**
 * Python + rdocx 引擎桥接。
 *
 * 每次 `parse` 都 spawn 一个短命子进程：进程间完全隔离，因此一份恶意文档
 * 无法影响宿主进程状态。
 */
export class PythonRdocxEngine implements DocxEngine {
  readonly name: string;
  private readonly config: EngineConfig;
  private readonly telemetry: TelemetrySink | undefined;

  constructor(config: EngineConfig, telemetry?: TelemetrySink) {
    this.config = config;
    this.name = `python-rdocx-${config.driver}`;
    this.telemetry = telemetry;
  }

  async parse(request: ParseRequest): Promise<RawParseResult> {
    // 路径解析与可读性检查在 spawn 之前完成，
    // 这样「文件不存在」这类常见失败不会被误报成引擎错误。
    const artifactPath = resolveArtifactPath(request.artifactPath);
    assertReadableFile(artifactPath);

    const scriptPath = this.config.scriptPath ?? defaultScriptPath();
    // limits / featureFlags 以 JSON 通过命令行传入，让引擎保持无状态、无配置文件。
    const payload = JSON.stringify({
      limits: request.limits,
      featureFlags: request.featureFlags,
    });

    const startedAt = Date.now();
    const outcome = await runParse(this.config, scriptPath, artifactPath, payload, request.timeoutMs);

    // 遥测只记规模与耗时，不记路径与内容。
    this.telemetry?.emit({
      level: 'debug',
      name: 'docx-complex-parse.engine.parse',
      attributes: {
        engine: this.name,
        engineMs: Date.now() - startedAt,
        exitCode: outcome.exitCode ?? -1,
        stdoutBytes: outcome.stdout.length,
        stderrBytes: outcome.stderr.length,
        timedOut: outcome.timedOut,
      },
    });

    // --- 失败路径：逐一映射为本模块的稳定错误码 ---------------------------- //
    if (outcome.timedOut) {
      throw new DocxComplexParseError(
        'ENGINE_TIMEOUT',
        `Parsing exceeded the ${request.timeoutMs} ms budget`,
        { details: { artifactPath } },
      );
    }
    if (outcome.exitCode !== 0) {
      throw new DocxComplexParseError('ENGINE_FAILED', 'Parse process exited with an error', {
        details: {
          artifactPath,
          exitCode: outcome.exitCode ?? -1,
          stderr: truncate(outcome.stderr, 2048),
        },
      });
    }

    // --- 协议解析 ---------------------------------------------------------- //
    let parsed: unknown;
    try {
      parsed = JSON.parse(outcome.stdout);
    } catch (error) {
      throw new DocxComplexParseError('ENGINE_PROTOCOL_ERROR', 'Parse did not emit valid JSON', {
        details: { stdout: truncate(outcome.stdout, 2048), stderr: truncate(outcome.stderr, 2048) },
        cause: error,
      });
    }

    try {
      // 失败信封优先：引擎没能读完文档时，不要拿半份结果去套成功路径的校验。
      const failure = parseEngineFailure(parsed);
      if (failure) {
        throw new DocxComplexParseError(codeForReason(failure.reason), failure.message, {
          details: { reason: failure.reason, ...failure.detail },
        });
      }
      // parseRawResult 是信任边界：畸形结构在这里被拒绝。
      return parseRawResult(parsed);
    } catch (error) {
      if (isDocxComplexParseError(error)) throw error;
      throw new DocxComplexParseError('ENGINE_PROTOCOL_ERROR', messageOf(error), {
        details: { artifactPath },
        cause: error,
      });
    }
  }

  async dispose(): Promise<void> {
    // 桥接层无状态：每次解析都会自行 spawn 并回收子进程，无需释放任何资源。
  }
}

/** 把未知值收敛成一句可读消息。 */
function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

/**
 * 按配置创建引擎。
 *
 * 选了尚未接入的候选引擎时快速失败——静默换成别的引擎会让宿主以为自己拿到的是
 * 另一个引擎的结果，这比直接报错危险得多。
 */
export function createEngine(config: EngineConfig, telemetry?: TelemetrySink): DocxEngine {
  const implemented: readonly string[] = IMPLEMENTED_ENGINE_DRIVERS;
  if (!implemented.includes(config.driver)) {
    throw new DocxComplexParseError(
      'ENGINE_UNAVAILABLE',
      `Engine driver "${config.driver}" is not implemented in this module.`,
      { details: { driver: config.driver, implemented: [...implemented] } },
    );
  }
  return new PythonRdocxEngine(config, telemetry);
}

/** 确认路径是「存在且为普通文件」，否则抛出 `ARTIFACT_NOT_FOUND`。 */
function assertReadableFile(path: string): void {
  let stats;
  try {
    stats = statSync(path);
  } catch (error) {
    throw new DocxComplexParseError('ARTIFACT_NOT_FOUND', 'Artifact path is not readable', {
      details: { artifactPath: path },
      cause: error,
    });
  }
  // 目录也能 stat 成功，但它不是 artifact，需要单独拒绝。
  if (!stats.isFile()) {
    throw new DocxComplexParseError('ARTIFACT_NOT_FOUND', 'Artifact path is not a regular file', {
      details: { artifactPath: path },
    });
  }
}

/** 子进程执行结果。 */
interface ParseOutcome {
  exitCode: number | null;
  stdout: string;
  stderr: string;
  timedOut: boolean;
}

/**
 * 启动解析子进程并收集输出。
 *
 * 该函数只负责「跑起来 + 收结果」，不做任何语义判断——
 * 超时、退出码、协议解析都在 `parse()` 里决策，便于单测替换。
 */
function runParse(
  config: EngineConfig,
  scriptPath: string,
  artifactPath: string,
  payload: string,
  timeoutMs: number,
): Promise<ParseOutcome> {
  return new Promise<ParseOutcome>((resolve, reject) => {
    const child = spawn(config.pythonPath, [scriptPath, '--path', artifactPath, '--config', payload], {
      // cwd 设为脚本所在目录：让脚本内的相对路径行为可预期。
      cwd: dirname(scriptPath),
      env: {
        ...process.env,
        ...config.env,
        // Windows 控制台默认编码不是 UTF-8，这里强制两侧统一，
        // 避免中文路径/内容在 stdout 上变成乱码导致 JSON 解析失败。
        PYTHONIOENCODING: 'utf-8',
        PYTHONDONTWRITEBYTECODE: '1',
        PYTHONUTF8: '1',
      },
      stdio: ['ignore', 'pipe', 'pipe'],
      // Windows 下不弹出控制台窗口。
      windowsHide: true,
    });

    let stdout = '';
    let stderr = '';
    let stdoutBytes = 0;
    let stderrBytes = 0;
    // settled 保证「超时」与「正常结束」只会有一个生效，避免重复 resolve/reject。
    let settled = false;
    let timedOut = false;

    const timer = setTimeout(() => {
      timedOut = true;
      // SIGKILL：解析脚本不写盘，无需优雅退出，直接强杀以防僵死。
      child.kill('SIGKILL');
    }, timeoutMs);

    child.stdout.on('data', (chunk: Buffer) => {
      if (stdoutBytes >= MAX_ENGINE_OUTPUT_BYTES) return;
      stdoutBytes += chunk.byteLength;
      stdout += chunk.toString('utf8');
    });

    child.stderr.on('data', (chunk: Buffer) => {
      if (stderrBytes >= MAX_ENGINE_OUTPUT_BYTES) return;
      stderrBytes += chunk.byteLength;
      stderr += chunk.toString('utf8');
    });

    // `error` 事件表示「进程根本没起来」（可执行文件不存在、无执行权限等），
    // 这与「进程起来了但失败」是两种不同的错误，需要区分。
    child.on('error', (error) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      reject(
        new DocxComplexParseError(
          'ENGINE_UNAVAILABLE',
          `Unable to launch the parse interpreter "${config.pythonPath}"`,
          { details: { scriptPath, reason: error.message }, cause: error },
        ),
      );
    });

    child.on('close', (exitCode) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve({ exitCode, stdout, stderr, timedOut });
    });
  });
}

/** 截断过长文本，用于错误上下文，避免错误对象体积失控。 */
function truncate(value: string, limit: number): string {
  if (value.length <= limit) return value;
  return `${value.slice(0, limit)}…`;
}
