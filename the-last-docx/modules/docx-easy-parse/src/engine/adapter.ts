/**
 * 引擎适配器 —— 本模块中【唯一】知道「解析是用 Python 的 zipfile + lxml 实现的」
 * 这一事实的文件。
 *
 * 让引擎可替换而又不改动公开契约的三条设计纪律：
 *   1. 适配器只返回 `ParseResult` 这种普通领域对象；解释器句柄、文件流、临时文件
 *      一律不得逃出本文件；
 *   2. 引擎通过 stdout 上的 JSON 与模块通信，因此换实现（python-docx、Docling、
 *      进程内 TS 读取器或远程服务）只需替换 `DocxEngine` 的一个实现；
 *   3. 引擎失败会被翻译成本模块的错误分类体系，引擎无权决定错误码。
 */
import { spawn } from 'node:child_process';
import { statSync } from 'node:fs';
import { dirname, isAbsolute, join, resolve as resolvePath } from 'node:path';
import { fileURLToPath } from 'node:url';

import type { EngineConfig, FeatureFlags, LimitConfig, TelemetrySink } from '../contract';
import { DocxParseError } from '../errors';
import { ParseProtocolError, parseParseResult, type ParseResult } from '../domain/docx-parse';

/**
 * 从子进程流中保留的字节上限。
 *
 * 为什么需要：子进程若因 bug 或恶意输入陷入死循环并狂写 stdout，
 * 无上限累积字符串会迅速耗尽 Node 堆内存。这里只保留前 16 MiB，
 * 后续内容直接丢弃——协议载荷远小于该阈值，丢的都是异常输出。
 *
 * 比 docx-inspect 的 8 MiB 略高，因为解析 IR 会把正文结构完整带回来。
 */
const MAX_ENGINE_OUTPUT_BYTES = 16 * 1024 * 1024;

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
 * 刻意保持极窄：只回答一个有界、无副作用的问题——「这个文档的结构是什么？」
 */
export interface DocxEngine {
  /** 引擎名，用于遥测与诊断。 */
  readonly name: string;
  parse(request: ParseRequest): Promise<ParseResult>;
  /** 释放资源（Python 桥接为无状态，因此为空实现）。 */
  dispose(): Promise<void>;
}

/** 内置解析脚本的绝对路径。 */
export function defaultScriptPath(): string {
  // 用 import.meta.url 而非 __dirname：本模块以 ESM 发布，且需要跨平台正确解析。
  return join(dirname(fileURLToPath(import.meta.url)), 'docx_parse.py');
}

/** 默认解析器明确不支持的 URI scheme；遇到时应提示调用方注入 office-files 物化器。 */
const UNSUPPORTED_SCHEMES = ['http:', 'https:', 'memory:', 'data:', 's3:', 'gs:'];

/**
 * 把 `ArtifactRef.uri` 解析为本地可读路径。
 *
 * 只接受独立模块确实读得懂的形态；其余一律以 `UNSUPPORTED_ARTIFACT_URI` 拒绝，
 * 让 Profile 有机会注入 `office-files` 物化器，而不是「默默读错东西」。
 *
 * @throws DocxParseError(INVALID_INPUT)             uri 为空
 * @throws DocxParseError(UNSUPPORTED_ARTIFACT_URI)  scheme 不受支持或 file:// 畸形
 */
export function resolveArtifactPath(uri: string): string {
  if (typeof uri !== 'string' || uri.trim() === '') {
    throw new DocxParseError('INVALID_INPUT', 'artifactRef.uri must be a non-empty string');
  }

  const lowered = uri.toLowerCase();
  if (UNSUPPORTED_SCHEMES.some((scheme) => lowered.startsWith(scheme))) {
    throw new DocxParseError(
      'UNSUPPORTED_ARTIFACT_URI',
      'The default resolver only reads local files; inject an office-files materializer for other schemes.',
      { details: { uri } },
    );
  }

  if (lowered.startsWith('file://')) {
    try {
      return fileURLToPath(uri);
    } catch (error) {
      throw new DocxParseError('UNSUPPORTED_ARTIFACT_URI', 'Malformed file:// URI', {
        details: { uri },
        cause: error,
      });
    }
  }

  // 出现任意 scheme 但不在白名单内（例如 `ftp://`）。
  if (uri.includes('://')) {
    throw new DocxParseError('UNSUPPORTED_ARTIFACT_URI', 'Unsupported URI scheme', {
      details: { uri },
    });
  }

  // 相对路径按进程工作目录解析；绝对路径直接使用。
  return isAbsolute(uri) ? uri : resolvePath(process.cwd(), uri);
}

/**
 * Python 桥接实现。
 *
 * 每次 `parse` 都会 spawn 一个短命子进程：进程间完全隔离，因此一份恶意文档
 * 无法影响宿主进程状态；代价是每次调用约几十毫秒的启动开销，
 * 对于「解析」这类低频重操作是可以接受的。
 */
export class PythonDocxEngine implements DocxEngine {
  readonly name = 'python-zipfile-lxml-parse';
  private readonly config: EngineConfig;
  private readonly telemetry?: TelemetrySink;

  constructor(config: EngineConfig, telemetry?: TelemetrySink) {
    this.config = config;
    if (telemetry !== undefined) {
      this.telemetry = telemetry;
    }
  }

  async parse(request: ParseRequest): Promise<ParseResult> {
    // 路径解析与可读性检查在 spawn 之前完成，
    // 这样「文件不存在」这类常见失败不会被误报成引擎错误。
    const artifactPath = resolveArtifactPath(request.artifactPath);
    assertReadableFile(artifactPath);

    const scriptPath = this.config.scriptPath ?? defaultScriptPath();
    // 把 limits / featureFlags 以 JSON 形式通过命令行传入，
    // 从而让引擎保持「无状态、无配置文件」的简单模型。
    const payload = JSON.stringify({
      limits: request.limits,
      featureFlags: request.featureFlags,
    });

    return executeScript(
      this.config,
      scriptPath,
      artifactPath,
      payload,
      request.timeoutMs,
      this.name,
      this.telemetry,
    );
  }

  async dispose(): Promise<void> {
    // 桥接层无状态：每次解析都会自行 spawn 并回收子进程，无需释放任何资源。
  }
}

/** 内置 Docling 桥接脚本的绝对路径。 */
export function defaultDoclingScriptPath(): string {
  return join(dirname(fileURLToPath(import.meta.url)), 'docx_parse_docling.py');
}

/**
 * Docling 深度引擎。
 *
 * 定位：为 `execute` / `verify` 提供解析质量更高的正文结构（段落、标题、表格、
 * 阅读顺序）。`inspect` 绝不会使用它——模型冷启动的成本与「先看一眼」的诉求相悖。
 *
 * 与轻量引擎共享同一套失败语义：超时、退出码、协议校验的映射逻辑完全一致，
 * 因此换引擎不会改变上层看到的错误分类（这是 spec 第八节的硬要求）。
 */
export class DoclingEngine implements DocxEngine {
  readonly name = 'python-docling-parse';
  private readonly spec: LaunchSpec;
  private readonly scriptPath: string;
  private readonly telemetry?: TelemetrySink;

  constructor(config: EngineConfig, telemetry?: TelemetrySink) {
    const deep = config.deep;
    if (deep === undefined) {
      throw new DocxParseError(
        'INVALID_INPUT',
        'DoclingEngine requires engine.deep to be configured',
      );
    }
    // env 合并顺序：全局引擎 env 为底，深度引擎 env 覆盖——
    // 允许为深度引擎单独设置模型缓存路径等变量而不影响轻量路径。
    const env = { ...config.env, ...deep.env };
    this.spec = {
      pythonPath: deep.pythonPath ?? config.pythonPath,
      ...(Object.keys(env).length > 0 ? { env } : {}),
    };
    this.scriptPath = deep.scriptPath ?? defaultDoclingScriptPath();
    if (telemetry !== undefined) {
      this.telemetry = telemetry;
    }
  }

  async parse(request: ParseRequest): Promise<ParseResult> {
    const artifactPath = resolveArtifactPath(request.artifactPath);
    assertReadableFile(artifactPath);

    // 桥接脚本以同一 payload 协议接收配置：换引擎不改协议，只改实现。
    const payload = JSON.stringify({
      limits: request.limits,
      featureFlags: request.featureFlags,
    });

    return executeScript(
      this.spec,
      this.scriptPath,
      artifactPath,
      payload,
      request.timeoutMs,
      this.name,
      this.telemetry,
    );
  }

  async dispose(): Promise<void> {
    // 同轻量引擎：无状态，无需释放资源。
  }
}

/**
 * 当且仅当配置声明了深度引擎时构造它。
 *
 * 返回 null 表示「未配置」，调用方据此回退到轻量引擎——
 * 这样默认行为与引入本能力之前完全一致。
 */
export function createDeepEngine(
  config: EngineConfig,
  telemetry?: TelemetrySink,
): DocxEngine | null {
  if (config.deep === undefined) return null;
  return new DoclingEngine(config, telemetry);
}


/**
 * 「启动脚本 → 收结果 → 翻译失败」的共享实现。
 *
 * 两个引擎都委托到这里，是为了让【错误分类完全一致】：spec 第八节要求
 * 更换底层引擎不能改写错误分类，因此超时/退出码/协议校验的映射必须只有一份。
 *
 * @throws DocxParseError(ENGINE_TIMEOUT|ENGINE_FAILED|ENGINE_PROTOCOL_ERROR|ENGINE_UNAVAILABLE)
 */
async function executeScript(
  spec: LaunchSpec,
  scriptPath: string,
  artifactPath: string,
  payload: string,
  timeoutMs: number,
  engineName: string,
  telemetry: TelemetrySink | undefined,
): Promise<ParseResult> {
  const startedAt = Date.now();
  const outcome = await runEngine(spec, scriptPath, artifactPath, payload, timeoutMs);

  // 记录引擎层遥测：只记规模与耗时，不记路径与内容。
  telemetry?.record({
    name: 'docx-parse.engine.parse',
    level: 'debug',
    timestampMs: startedAt,
    durationMs: Date.now() - startedAt,
    attributes: {
      engine: engineName,
      exitCode: outcome.exitCode,
      stdoutBytes: outcome.stdout.length,
      stderrBytes: outcome.stderr.length,
      timedOut: outcome.timedOut,
    },
  });

  // --- 失败路径：逐一映射为本模块的稳定错误码 ---------------------------- //
  if (outcome.timedOut) {
    throw new DocxParseError('ENGINE_TIMEOUT', `Parse exceeded the ${timeoutMs} ms budget`, {
      path: artifactPath,
      details: { engine: engineName },
    });
  }
  if (outcome.exitCode !== 0) {
    throw new DocxParseError('ENGINE_FAILED', 'Parse process exited with an error', {
      path: artifactPath,
      details: {
        engine: engineName,
        exitCode: outcome.exitCode ?? -1,
        // stderr 是人类诊断信息，会截断后再上报，避免污染错误对象体积。
        stderr: truncate(outcome.stderr, 2048),
      },
    });
  }

  // --- 协议解析 ---------------------------------------------------------- //
  let parsed: unknown;
  try {
    parsed = JSON.parse(outcome.stdout);
  } catch (error) {
    throw new DocxParseError('ENGINE_PROTOCOL_ERROR', 'Parse engine did not emit valid JSON', {
      path: artifactPath,
      details: {
        engine: engineName,
        stdout: truncate(outcome.stdout, 2048),
        stderr: truncate(outcome.stderr, 2048),
      },
      cause: error,
    });
  }

  try {
    // parseParseResult 是信任边界：畸形结构在这里被拒绝。
    // 深度引擎的载荷同样要过这一关——换引擎不降低校验强度。
    return parseParseResult(parsed);
  } catch (error) {
    if (error instanceof ParseProtocolError) {
      throw new DocxParseError('ENGINE_PROTOCOL_ERROR', error.message, {
        path: artifactPath,
        details: { engine: engineName, pointer: error.pointer },
        cause: error,
      });
    }
    throw error;
  }
}

/** 确认路径是「存在且为普通文件」，否则抛出 `ARTIFACT_NOT_FOUND`。 */
function assertReadableFile(path: string): void {
  let stats;
  try {
    stats = statSync(path);
  } catch (error) {
    throw new DocxParseError('ARTIFACT_NOT_FOUND', 'Artifact path is not readable', {
      path,
      cause: error,
    });
  }
  // 目录也能 stat 成功，但它不是 artifact，需要单独拒绝。
  if (!stats.isFile()) {
    throw new DocxParseError('ARTIFACT_NOT_FOUND', 'Artifact path is not a regular file', {
      path,
    });
  }
}

/** 子进程执行结果。 */
interface EngineOutcome {
  exitCode: number | null;
  stdout: string;
  stderr: string;
  timedOut: boolean;
}

/**
 * 启动子进程所需的最小信息。
 *
 * 刻意不复用 `EngineConfig`：轻量与深度引擎各有一份解释器与 env，
 * 抽成窄接口后两者可以共用同一套「启动 + 收字节 + 判超时」逻辑，
 * 而调用方负责决定用哪一份配置。
 */
interface LaunchSpec {
  pythonPath: string;
  env?: Record<string, string> | undefined;
}

/**
 * 启动解析子进程并收集输出。
 *
 * 该函数只负责「跑起来 + 收结果」，不做任何语义判断——
 * 超时、退出码、协议解析都在各自的引擎实现里决策，便于单测替换。
 */
function runEngine(
  spec: LaunchSpec,
  scriptPath: string,
  artifactPath: string,
  payload: string,
  timeoutMs: number,
): Promise<EngineOutcome> {
  return new Promise<EngineOutcome>((resolve, reject) => {
    const child = spawn(spec.pythonPath, [scriptPath, '--path', artifactPath, '--config', payload], {
      // cwd 设为脚本所在目录：让脚本内的相对路径行为可预期，
      // 同时使桥接脚本可以用 `import docx_parse` 复用同一目录下的 OPC 层实现。
      cwd: dirname(scriptPath),
      env: {
        ...process.env,
        ...spec.env,
        // Windows 控制台默认编码不是 UTF-8，这里强制两侧统一，
        // 避免中文路径/内容在 stdout 上变成乱码导致 JSON 解析失败。
        PYTHONIOENCODING: 'utf-8',
        // 不生成 .pyc，保持部署目录干净（解析是短命功能）。
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
        new DocxParseError(
          'ENGINE_UNAVAILABLE',
          `Unable to launch the parse interpreter "${spec.pythonPath}"`,
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
