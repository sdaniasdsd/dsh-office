/**
 * Vitest 全局准备：在跑任何测试之前，用生成器重建全部 fixture。
 *
 * 为什么不在仓库里直接提交二进制样本：
 *   - 二进制不可审计，也无法解释「它为什么长这样」；
 *   - 生成器本身就是文档：读一遍就知道每个样本覆盖哪个场景；
 *   - 避免把恶意样本特征提交进仓库触发杀软告警。
 */
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

export default function setup(): void {
  const projectRoot = fileURLToPath(new URL('..', import.meta.url));
  const python = process.env['DOCX_PARSE_PYTHON'] ?? 'python';

  // 先探测解释器：契约测试必须在没有 Python 的环境下也能跑通
  // （spec 完成标准第一条），因此缺解释器时只是跳过 fixture 生成，
  // 由依赖 fixture 的用例自行 skip。
  const probe = spawnSync(python, ['-c', 'print(1)'], { encoding: 'utf8' });
  if (probe.status !== 0) {
    console.warn(
      `[docx-parse] python "${python}" is unavailable; real-engine tests will be skipped.`,
    );
    return;
  }

  const result = spawnSync(
    python,
    ['fixtures/generate_fixtures.py', '--out', 'fixtures/_generated'],
    { cwd: projectRoot, encoding: 'utf8' },
  );

  if (result.status !== 0) {
    // fixture 缺失会连带影响全部用例，因此直接失败并打印原因。
    throw new Error(
      `fixture generation failed (exit ${result.status}):\n${result.stderr || result.stdout}`,
    );
  }
}
