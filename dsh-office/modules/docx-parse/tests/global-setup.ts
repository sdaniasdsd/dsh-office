/**
 * Vitest 全局准备：在跑任何测试之前，用生成器重建全部 fixture。
 *
 * 为什么不在仓库里直接提交二进制样本：
 *   - 二进制样本不可审计，也无法解释「它为什么长这样」；
 *   - 生成器本身就是文档：读一遍就知道每个样本覆盖哪个场景；
 *   - 避免杀软把恶意样本特征提交进仓库。
 */
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

export default function setup(): void {
  const projectRoot = fileURLToPath(new URL('..', import.meta.url));
  const result = spawnSync(
    process.env['DOCX_PARSE_PYTHON'] ?? 'python',
    ['fixtures/generate_fixtures.py', '--out', 'fixtures/_generated'],
    { cwd: projectRoot, encoding: 'utf8' },
  );

  if (result.status !== 0) {
    // 直接失败并打印 stderr：fixture 缺失会连带影响全部测试，不值得继续跑。
    throw new Error(
      `fixture generation failed (exit ${result.status}):\n${result.stderr || result.stdout}`,
    );
  }
}
