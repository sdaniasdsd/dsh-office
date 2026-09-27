// Import the INSTALLED plugin entry exactly the way DSH does. Node resolves the
// entry's own bare import ('@deepseek-ai/dsh-mcp-client') from the entry's real
// location; if that fails at boot, the whole web profile fails to load.
const root = 'file:///C:/Users/AA/AppData/Roaming/com.yeagoo.dsh-desktop/harness/profiles/web/node_modules/@deepseek-ai/dsh-docx/';

const mod = await import(root + 'dsh/index.mjs');
console.log('entry module loaded OK');
console.log('  name   =', mod.name);
console.log('  inject =', JSON.stringify(mod.inject));
console.log('  apply  =', typeof mod.apply);
