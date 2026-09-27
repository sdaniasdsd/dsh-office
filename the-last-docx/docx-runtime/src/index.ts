import { existsSync, readFileSync, statSync } from 'node:fs';
import { dirname, isAbsolute, join, relative, resolve, sep } from 'node:path';

export const DSH_RUNTIME_SCHEMA = 'dsh-office-runtime/v1' as const;
export const DEFAULT_RUNTIME_PACKAGES = ['@deepseek-ai/dsh-docx-runtime'] as const;

export type RuntimeSource = 'config-path' | 'config-root' | 'env-path' | 'env-root' | 'runtime-package' | 'bundled' | 'unresolved';

export interface RuntimeComponentSpec {
  /** Stable component identifier used by the runtime manifest. */
  id: string;
  /** Environment variable that can override this component path. */
  envVar?: string;
  /** Backward-compatible path when a root has no usable manifest. */
  defaultEntry?: string;
}

export interface RuntimeResolverOptions {
  /** The Profile-configured root, either the platform directory or a parent of it. */
  runtimeRoot?: string;
  /** Direct, component-level Profile configuration. */
  componentPaths?: Readonly<Record<string, string | undefined>>;
  /** Components needed by this consumer. Any identifier is supported. */
  components: readonly RuntimeComponentSpec[];
  /** Extra runtime packages. The DSH DOCX runtime remains the first default. */
  runtimePackageNames?: readonly string[];
  /** A consumer package file or directory used to locate sibling node_modules. */
  callerPath?: string;
  /** Optional legacy bundled runtime root. */
  bundledRuntimeRoot?: string;
  /** Injectable for deterministic tests and hosts that sandbox environment access. */
  environment?: Readonly<Record<string, string | undefined>>;
}

export interface ResolvedRuntimeComponent {
  id: string;
  path?: string;
  source: RuntimeSource;
  available: boolean;
  /** Candidate paths considered in precedence order, excluding direct command fallbacks. */
  candidates: readonly string[];
}

export interface RuntimeResolution {
  schema?: string;
  runtimeRoot?: string;
  source: RuntimeSource;
  components: Readonly<Record<string, ResolvedRuntimeComponent>>;
  missing: readonly string[];
}

interface RuntimeManifest {
  schema?: unknown;
  components?: Record<string, { present?: unknown; entry?: unknown }>;
}

interface RuntimeLocation {
  root: string;
  manifest?: RuntimeManifest;
}

function nonBlank(value: string | undefined): string | undefined {
  return typeof value === 'string' && value.trim() !== '' ? value.trim() : undefined;
}

function isDirectory(path: string): boolean {
  try {
    return statSync(path).isDirectory();
  } catch {
    return false;
  }
}

function safeReadManifest(packageRoot: string): RuntimeManifest | undefined {
  const manifestPath = join(packageRoot, 'runtime.json');
  try {
    const parsed: unknown = JSON.parse(readFileSync(manifestPath, 'utf8'));
    return parsed !== null && typeof parsed === 'object' ? parsed as RuntimeManifest : undefined;
  } catch {
    return undefined;
  }
}

function normalizedRoot(path: string): string | undefined {
  const absolute = resolve(path);
  const candidates = [
    join(absolute, 'runtime', 'win32-x64'),
    join(absolute, 'win32-x64'),
    absolute,
  ];
  return candidates.find(isDirectory);
}

function safeEntry(root: string, entry: unknown): string | undefined {
  if (typeof entry !== 'string' || entry.trim() === '' || isAbsolute(entry)) return undefined;
  const candidate = resolve(root, entry);
  const withinRoot = relative(root, candidate);
  if (withinRoot === '..' || withinRoot.startsWith(`..${sep}`) || isAbsolute(withinRoot)) return undefined;
  return candidate;
}

function locationFromRoot(path: string | undefined): RuntimeLocation | undefined {
  const supplied = nonBlank(path);
  if (!supplied) return undefined;
  const root = normalizedRoot(supplied);
  if (!root) return undefined;
  const packageRoot = dirname(dirname(root));
  return { root, manifest: safeReadManifest(packageRoot) };
}

function findRuntimePackage(callerPath: string | undefined, packageNames: readonly string[]): RuntimeLocation | undefined {
  const start = callerPath === undefined ? process.cwd() : (isDirectory(callerPath) ? callerPath : dirname(callerPath));
  let current = resolve(start);
  while (true) {
    for (const packageName of packageNames) {
      const packageRoot = join(current, 'node_modules', packageName);
      const root = normalizedRoot(packageRoot);
      if (root) return { root, manifest: safeReadManifest(packageRoot) };
    }
    const parent = dirname(current);
    if (parent === current) return undefined;
    current = parent;
  }
}

function candidateFromLocation(location: RuntimeLocation | undefined, spec: RuntimeComponentSpec): string | undefined {
  if (!location) return undefined;
  const component = location.manifest?.components?.[spec.id];
  if (component?.present === false) return undefined;
  return safeEntry(location.root, component?.entry) ?? safeEntry(location.root, spec.defaultEntry);
}

/**
 * Resolve component paths using the DSH runtime contract. This function never
 * throws because an unavailable runtime is a supported, diagnosable state.
 */
export function resolveRuntime(options: RuntimeResolverOptions): RuntimeResolution {
  const environment = options.environment ?? process.env;
  const packageNames = [...DEFAULT_RUNTIME_PACKAGES, ...(options.runtimePackageNames ?? [])]
    .filter((name, index, names) => nonBlank(name) !== undefined && names.indexOf(name) === index);
  const configRoot = locationFromRoot(options.runtimeRoot);
  const envRoot = locationFromRoot(environment['DSH_OFFICE_RUNTIME_ROOT']);
  const packageLocation = findRuntimePackage(options.callerPath, packageNames);
  const bundledLocation = locationFromRoot(options.bundledRuntimeRoot);
  const resolved: Record<string, ResolvedRuntimeComponent> = {};

  for (const spec of options.components) {
    const directConfig = nonBlank(options.componentPaths?.[spec.id]);
    const directEnvironment = nonBlank(spec.envVar === undefined ? undefined : environment[spec.envVar]);
    const candidatesBySource: Array<{ path: string; source: RuntimeSource }> = [
      directConfig === undefined ? undefined : { path: directConfig, source: 'config-path' },
      candidateFromLocation(configRoot, spec) === undefined ? undefined : { path: candidateFromLocation(configRoot, spec)!, source: 'config-root' },
      directEnvironment === undefined ? undefined : { path: directEnvironment, source: 'env-path' },
      candidateFromLocation(envRoot, spec) === undefined ? undefined : { path: candidateFromLocation(envRoot, spec)!, source: 'env-root' },
      candidateFromLocation(packageLocation, spec) === undefined ? undefined : { path: candidateFromLocation(packageLocation, spec)!, source: 'runtime-package' },
      candidateFromLocation(bundledLocation, spec) === undefined ? undefined : { path: candidateFromLocation(bundledLocation, spec)!, source: 'bundled' },
    ].filter((value): value is { path: string; source: RuntimeSource } => value !== undefined);
    const selected = candidatesBySource.find((candidate) => existsSync(candidate.path));
    const source: RuntimeSource = selected?.source ?? 'unresolved';
    resolved[spec.id] = {
      id: spec.id,
      path: selected?.path,
      source,
      available: selected !== undefined,
      candidates: candidatesBySource.map((candidate) => candidate.path),
    };
  }

  const values = Object.values(resolved);
  const source = values.find((component) => component.source !== 'unresolved')?.source ?? 'unresolved';
  return {
    schema: packageLocation?.manifest?.schema === DSH_RUNTIME_SCHEMA ? DSH_RUNTIME_SCHEMA : undefined,
    runtimeRoot: configRoot?.root ?? envRoot?.root ?? packageLocation?.root ?? bundledLocation?.root,
    source,
    components: resolved,
    missing: values.filter((component) => !component.available).map((component) => component.id),
  };
}

