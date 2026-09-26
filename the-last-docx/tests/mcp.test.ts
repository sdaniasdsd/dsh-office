import { describe,it,expect } from 'vitest';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { fileURLToPath } from 'node:url';
import { createMcpServer } from '../src/mcp';
import { context,createRequest } from './support';
describe('MCP plugin protocol',()=>{
  it('lists real tools, executes creation and returns classified errors',async()=>{
    const c=await context(),server=createMcpServer(c.profile),client=new Client({name:'test',version:'1.0.0'}),[left,right]=InMemoryTransport.createLinkedPair();
    try{await server.connect(left);await client.connect(right);
      expect((await client.listTools()).tools.map(t=>t.name)).toContain('docx_call');
      const result=await client.callTool({name:'docx_call',arguments:{moduleId:'docx-create',operation:'execute',input:createRequest}});expect(result.isError).not.toBe(true);
      const invalid=await client.callTool({name:'docx_call',arguments:{moduleId:'docx-create',operation:'execute',input:{requestId:'bad'}}});expect(invalid.isError).toBe(true);
    }finally{await client.close();await server.close();await c.close();}
  });
  it('boots through the actual relocatable stdio launcher',async()=>{
    const c=await context(),client=new Client({name:'stdio-test',version:'1.0.0'});
    const env=Object.fromEntries(Object.entries(process.env).filter((entry):entry is [string,string]=>typeof entry[1]==='string'));
    const transport=new StdioClientTransport({command:process.execPath,args:[fileURLToPath(new URL('../scripts/start.mjs',import.meta.url))],env:{...env,THE_LAST_DOCX_WORKSPACE:c.inputs,THE_LAST_DOCX_DATA:c.files.root},stderr:'pipe'});
    try{await client.connect(transport);
      // Asserted as names, not a count: a bare `toHaveLength(6)` went stale the
      // moment a tool was added, and said nothing about which tool was missing.
      const tools=(await client.listTools()).tools.map(t=>t.name).sort();
      expect(tools).toEqual(['docx_analyze','docx_call','docx_doctor','docx_from_reference','docx_import','docx_modules','docx_read_artifact']);
    }finally{await client.close();await c.close();}
  });
});
