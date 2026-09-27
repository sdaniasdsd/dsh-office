# DOCX runtime discovery

This package implements the `dsh-office-runtime/v1` consumer contract without
shipping a runtime itself. Modules describe the components they need, while the
resolver handles configuration, environment variables, runtime packages, and
the old bundled layout.

Discovery order for every component is: explicit component path, configured
runtime root, component environment variable, runtime-root environment
variable, sibling runtime package, then a bundled runtime. Missing components
remain non-fatal and are reported in the returned diagnostic data.

The default runtime package is `@deepseek-ai/dsh-docx-runtime`. Consumers can
append package names and component specifications without changing this
package, so the resolver can be shared by future Office plugins.

## Module configuration

The adapted DOCX modules accept the following top-level configuration:

```ts
{
  // May be the package root, runtime/, or runtime/win32-x64.
  runtimeRoot: 'D:\\toolchains\\dsh-docx-runtime',
  // Optional: add a host-owned runtime package after the DSH default package.
  runtimePackageNames: ['@acme/office-runtime'],
}
```

Existing `engine.pythonPath`, `engine.sofficePath`, and
`engine.pdftoppmPath` settings remain the highest-priority per-component
configuration. Without an explicit configuration, the resolver also recognizes
`DOCX_PYTHON`, `DOCX_SOFFICE`, `DOCX_PDFTOPPM`, and
`DSH_OFFICE_RUNTIME_ROOT`.

When a component is absent, no nonexistent executable is injected. The module
keeps its normal PATH fallback and reports the existing `ENGINE_UNAVAILABLE`
error only if an operation actually needs that component. Hosts implementing a
doctor command can use `resolveRuntime(...).missing` for per-component status.

