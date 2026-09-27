// 顺带导出纯文本，方便直接改词：内容取自同一份 speech-content.mjs，
// 所以文本与文档不会走偏。
import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { SPEECH_DOCUMENT } from './speech-content.mjs';

const lines = [];
let characters = 0;
for (const block of SPEECH_DOCUMENT.blocks) {
  if (block.kind === 'paragraph') {
    const text = block.runs.map((run) => run.text).join('');
    if (block.style === 'Title') lines.push(`# ${text}`, '');
    else if (block.style === 'Subtitle') lines.push(`## ${text}`, '');
    else if (block.style?.startsWith('Heading')) lines.push(`### ${text}`, '');
    else lines.push(text, '');
    if (block.style !== 'Title' && block.style !== 'Subtitle') characters += text.length;
  } else if (block.kind === 'list') {
    block.items.forEach((item, index) => lines.push(`${block.ordered ? `${index + 1}.` : '-'} ${item}`));
    lines.push('');
    characters += block.items.join('').length;
  }
}
const out = join('C:\\Users\\AA\\Desktop\\bench\\out-speech', '嵌入式的发展路径-演讲稿.md');
writeFileSync(out, lines.join('\n'), 'utf8');
console.log(`wrote ${out}`);
console.log(`字数（不含标题、副标题）：${characters}`);
console.log(`按每分钟 200 字估算：约 ${Math.round(characters / 200)} 分钟`);
