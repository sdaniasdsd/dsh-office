import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { createHost } from './host';
import { createMcpServer } from './mcp';
const profile=await createHost(),server=createMcpServer(profile);
// All protocol output belongs to MCP; engine diagnostics must stay on stderr.
const stop=async()=>{await server.close();await profile.dispose();};
process.once('SIGTERM',()=>{void stop().finally(()=>process.exit(0));});
process.once('SIGINT',()=>{void stop().finally(()=>process.exit(0));});
await server.connect(new StdioServerTransport());
