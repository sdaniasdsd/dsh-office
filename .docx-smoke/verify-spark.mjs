// Pre-restart gate: this is exactly what failed at boot (loader-import-failure).
// Import spark-plugin's entry the way the DSH loader does, from its real path,
// and confirm every bare specifier it needs now resolves.
const entry = 'file:///E:/dsh_project/spark-plugin-dsh/index.mjs';

try {
  const mod = await import(entry);
  console.log('SPARK ENTRY IMPORT: OK');
  console.log('  name   =', mod.name);
  console.log('  inject =', JSON.stringify(mod.inject));
  console.log('  apply  =', typeof mod.apply);
} catch (err) {
  console.log('SPARK ENTRY IMPORT: FAILED');
  console.log('  ' + (err && err.message ? err.message : String(err)));
  if (err && err.code) console.log('  code =', err.code);
  process.exitCode = 1;
}
