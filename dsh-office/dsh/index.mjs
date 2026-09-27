/** DSH namespace plugin. The host bridge owns tool effects and process teardown. */
import * as mcpClient from '@deepseek-ai/dsh-mcp-client';
import { resolveDshConfig } from './runtime-config.mjs';
export const name = 'docx';
export const inject = ['tools'];
export async function apply(ctx, config = {}) {
  const connection = await resolveDshConfig(config);
  await ctx.plugin(mcpClient, connection);
}
