/**
 * 引擎适配器 —— 本模块中【唯一】知道「探测是用 Python 的 zipfile + lxml 实现的」
 * 这一事实的文件。
 *
 * 让引擎可替换而又不改动公开契约的三条设计纪律：
 *   1. 适配器只返回 `ProbeResult` 这种普通领域对象；解释器句柄、文件流、临时文件
 *      一律不得逃出本文件；
 *   2. 引擎通过 stdout 上的 JSON 与模块通信，因此换实现（进程内 TS 读取器、
 *      基于 oletools 的分析器、远程服务）只需替换 `DocxEngine` 的一个实现；
 *   3. 引擎失败会被翻译成本模块的错误分类体系，引擎无权决定错误码。
 */
import { spawn } from 'node:child_process';
import { statSync } from 'node:fs';
import { dirname, isAbsolute, join, resolve as resolvePath } from 'node:path';
import { fileURLToPath } from 'node:url';

import type { TelemetrySink } from '../contract';
import type { EngineConfig, FeatureFlags, LimitConfig } from '../contract';
import { DocxInspectError } from '../errors';
import { ProbeProtocolError, parseProbeResult, type ProbeResult } from '../domain/docx-inspect';

/**
 * 从子进程流中保留的字节上限。
 *
 * 为什么需要：子进程若因 bug 或恶意输入陷入死循环并狂写 stdout，
 * 无上限累积字符串会迅速耗尽 Node 堆内存。这里只保留前 8 MiB，
 * 后续内容直接丢弃——协议载荷远小于该阈值，丢的都是异常输出。
 */
const MAX_ENGINE_OUTPUT_BYTES = 8 * 1024 * 1024;

/** 一次探测请求所需的全部输入。 */
export interface ProbeRequest {
  /** 已解析为本地可读路径的 artifact。 */
  artifactPath: string;
  limits: LimitConfig;
  featureFlags: FeatureFlags;
  timeoutMs: number;
}

/**
 * 所有引擎后端都必须实现的端口。
 *
 * 刻意保持极窄：只回答一个有界、无副作用的问题——「这个产物里有什么？」
 */
export interface DocxEngine {
  /** 引擎名，用于遥测与诊断。 */
  readonly name: string;
  probe(request: ProbeRequest): Promise<ProbeResult>;
  /** 释放资源（Python 桥接为无状态，因此为空实现）。 */
  dispose(): Promise<void>;
}

/** 内置探测脚本的绝对路径。 */
export function defaultScriptPath(): string {
  // 用 import.meta.url 而非 __dirname：本模块以 ESM 发布，且需要跨平台正确解析。
  return join(dirname(fileURLToPath(import.meta.url)), 'docx_probe.py');
}

/** 默认解析器明确不支持的 URI scheme；遇到时应提示调用方注入 office-files 物化器。 */
const UNSUPPORTED_SCHEMES = ['http:', 'https:', 'memory:', 'data:', 's3:', 'gs:'];

/**
 * 把 `ArtifactRef.uri` 解析为本地可读路径。
 *
 * 只接受独立模块确实读得懂的形态；其余一律以 `UNSUPPORTED_ARTIFACT_URI` 拒绝，
 * 让 Profile 有机会注入 `office-files` 物化器，而不是「默默读错东西」。
 *
 * @throws DocxInspectError(INVALID_INPUT)              uri 为空
 * @throws DocxInspectError(UNSUPPORTED_ARTIFACT_URI)   scheme 不受支持或 file:// 畸形
 */
export function resolveArtifactPath(uri: string): string {
  if (typeof uri !== 'string' || uri.trim() === '') {
    throw new DocxInspectError('INVALID_INPUT', 'artifactRef.uri must be a non-empty string');
  }

  const lowered = uri.toLowerCase();
  if (UNSUPPORTED_SCHEMES.some((scheme) => lowered.startsWith(scheme))) {
    throw new DocxInspectError(
      'UNSUPPORTED_ARTIFACT_URI',
      'The default resolver only reads local files; inject an office-files materializer for other schemes.',
      { details: { uri } },
    );
  }

  if (lowered.startsWith('file://')) {
    try {
      return fileURLToPath(uri);
    } catch (error) {
      throw new DocxInspectError('UNSUPPORTED_ARTIFACT_URI', 'Malformed file:// URI', {
        details: { uri },
        cause: error,
      });
    }
  }

  // 出现任意 scheme 但不在白名单内（例如 `ftp://`）。
  if (uri.includes('://')) {
    throw new DocxInspectError('UNSUPPORTED_ARTIFACT_URI', 'Unsupported URI scheme', {
      details: { uri },
    });
  }

  // 相对路径按进程工作目录解析；绝对路径直接使用。
  return isAbsolute(uri) ? uri : resolvePath(process.cwd(), uri);
}

/**
 * Python 桥接实现。
 *
 * 每次 `probe` 都会 spawn 一个短命子进程：进程间完全隔离，因此一份恶意文档
 * 无法影响宿主进程状态；代价是每次调用约几十毫秒的启动开销，
 * 对于「探测」这类低频重操作是可以接受的。
 */
export class PythonDocxEngine implements DocxEngine {
  readonly name = 'python-zipfile-lxml-probe';
  private readonly config: EngineConfig;
  private readonly telemetry?: TelemetrySink;

  constructor(config: EngineConfig, telemetry?: TelemetrySink) {
    this.config = config;
    if (telemetry !== undefined) {
      this.telemetry = telemetry;
    }
  }

  async probe(request: ProbeRequest): Promise<ProbeResult> {
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

    const startedAt = Date.now();
    const outcome = await runProbe(this.config, scriptPath, artifactPath, payload, request.timeoutMs);

    // 记录引擎层遥测：只记规模与耗时，不记路径与内容。
    this.telemetry?.record({
      name: 'docx-inspect.engine.probe',
      level: 'debug',
      timestampMs: startedAt,
      durationMs: Date.now() - startedAt,
      attributes: {
        engine: this.name,
        exitCode: outcome.exitCode,
        stdoutBytes: outcome.stdout.length,
        stderrBytes: outcome.stderr.length,
        timedOut: outcome.timedOut,
      },
    });

    // --- 失败路径：逐一映射为本模块的稳定错误码 ---------------------------- //
    if (outcome.timedOut) {
      throw new DocxInspectError(
        'ENGINE_TIMEOUT',
        `Probe exceeded the ${request.timeoutMs} ms budget`,
        { path: artifactPath },
      );
    }
    if (outcome.exitCode !== 0) {
      throw new DocxInspectError('ENGINE_FAILED', 'Probe process exited with an error', {
        path: artifactPath,
        details: {
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
      throw new DocxInspectError('ENGINE_PROTOCOL_ERROR', 'Probe did not emit valid JSON', {
        path: artifactPath,
        details: { stdout: truncate(outcome.stdout, 2048), stderr: truncate(outcome.stderr, 2048) },
        cause: error,
      });
    }

    try {
      // parseProbeResult 是信任边界：畸形结构在这里被拒绝。
      return parseProbeResult(parsed);
    } catch (error) {
      if (error instanceof ProbeProtocolError) {
        throw new DocxInspectError('ENGINE_PROTOCOL_ERROR', error.message, {
          path: artifactPath,
          details: { pointer: error.pointer },
          cause: error,
        });
      }
      throw error;
    }
  }

  async dispose(): Promise<void> {
    // 桥接层无状态：每次探测都会自行 spawn 并回收子进程，无需释放任何资源。
  }
}

/** 确认路径是「存在且为普通文件」，否则抛出 `ARTIFACT_NOT_FOUND`。 */
function assertReadableFile(path: string): void {
  let stats;
  try {
    stats = statSync(path);
  } catch (error) {
    throw new DocxInspectError('ARTIFACT_NOT_FOUND', 'Artifact path is not readable', {
      path,
      cause: error,
    });
  }
  // 目录也能 stat 成功，但它不是 artifact，需要单独拒绝。
  if (!stats.isFile()) {
    throw new DocxInspectError('ARTIFACT_NOT_FOUND', 'Artifact path is not a regular file', {
      path,
    });
  }
}

/** 子进程执行结果。 */
interface ProbeOutcome {
  exitCode: number | null;
  stdout: string;
  stderr: string;
  timedOut: boolean;
}

/**
 * 启动探测子进程并收集输出。
 *
 * 该函数只负责「跑起来 + 收结果」，不做任何语义判断——
 * 超时、退出码、协议解析都在 `probe()` 里决策，便于单测替换。
 */
function runProbe(
  config: EngineConfig,
  scriptPath: string,
  artifactPath: string,
  payload: string,
  timeoutMs: number,
): Promise<ProbeOutcome> {
  return new Promise<ProbeOutcome>((resolve, reject) => {
    const child = spawn(
      config.pythonPath,
      [scriptPath, '--path', artifactPath, '--config', payload],
      {
        // cwd 设为脚本所在目录：让脚本内的相对路径行为可预期。
        cwd: dirname(scriptPath),
        env: {
          ...process.env,
          ...config.env,
          // Windows 控制台默认编码不是 UTF-8，这里强制两侧统一，
          // 避免中文路径/内容在 stdout 上变成乱码导致 JSON 解析失败。
          PYTHONIOENCODING: 'utf-8',
          // 不生成 .pyc，保持部署目录干净（探测是短命功能）。
          PYTHONDONTWRITEBYTECODE: '1',
          PYTHONUTF8: '1',
        },
        stdio: ['ignore', 'pipe', 'pipe'],
        // Windows 下不弹出控制台窗口。
        windowsHide: true,
      },
    );

    let stdout = '';
    let stderr = '';
    let stdoutBytes = 0;
    let stderrBytes = 0;
    // settled 保证「超时」与「正常结束」只会有一个生效，避免重复 resolve/reject。
    let settled = false;
    let timedOut = false;

    const timer = setTimeout(() => {
      timedOut = true;
      // SIGKILL：探测脚本不写盘，无需优雅退出，直接强杀以防僵死。
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
        new DocxInspectError(
          'ENGINE_UNAVAILABLE',
          `Unable to launch the probe interpreter "${config.pythonPath}"`,
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
