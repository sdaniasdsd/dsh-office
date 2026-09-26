import { readFile } from 'node:fs/promises';
import { createHost } from './host';
import { doctor } from './doctor';
import type { ModuleId } from './profile';
const [command,...args]=process.argv.slice(2);
const profile=await createHost();
try{
  let result:unknown;
  if(command==='doctor')result=await doctor(profile.options);
  else if(command==='list')result=profile.list();
  else if(command==='import'&&args[0])result=await profile.files.importFile(args[0]);
  else if(command==='call'&&args[0]&&args[1]&&args[2])result=await profile.call(args[0] as ModuleId,args[1] as 'inspect'|'execute'|'verify',JSON.parse(await readFile(args[2],'utf8')));
  else throw new Error('Usage: doctor | list | import <path> | call <moduleId> <inspect|execute|verify> <request.json>');
  process.stdout.write(JSON.stringify(result,null,2)+'\n');
}catch(error){process.stderr.write(JSON.stringify({error:error instanceof Error?error.message:'failed',code:(error as {code?:string}).code??'HOST_FAILED'})+'\n');process.exitCode=1;}
finally{await profile.dispose();}
