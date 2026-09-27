# DSH Office stress baseline and replay

This runbook freezes the 15-material source corpus and replays the known Windows PPTX encoding failure through the installed DSH Profile MCP stdio server. It is test-only: it must not edit any production module or modify the frozen `office-stress-20260927-run10` evidence.

## Frozen baseline

Baseline root:

```text
dist/office-stress-20260927-run10
```

The baseline contains 5 DOCX, 5 PPTX and 5 XLSX input specimens. The five PDF stress materials are DSH render outputs derived from the five DOCX specimens, not five independent native PDFs. `manifest.json` records source and isolated-workspace SHA-256 values, URLs, category and sanitization metadata. `report.json` records 200 operation rounds, request IDs, durations and failure text; the historical run has 231 DSH tool calls including imports, preflights and Doctor.

The three DOCX files with external relationships were sanitized only in isolated workspace copies to comply with the Profile's offline safety policy. The originals remain under `sources/`; the two hashes in the manifest must not be conflated.

## Verify frozen corpus and historical report

From the repository root, use the configured DSH Node runtime:

```powershell
$node = 'C:\Users\AA\AppData\Local\DSH Desktop\runtime\node.exe'
$baseline = 'dist\office-stress-20260927-run10'
$evidence = Join-Path 'dist' ('office-stress-b0-replay-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
if (Test-Path -LiteralPath $evidence) { throw "Evidence path already exists: $evidence" }
New-Item -ItemType Directory -Path $evidence | Out-Null
& $node 'scripts\verify-office-stress-baseline.mjs' $baseline "$evidence\corpus-verification.json"
```

Verification reads every source and workspace file, checks its SHA-256 against the manifest, confirms five files per input format, confirms the 200-operation/231-tool-call historical run and the 30 expected PPTX failures, and writes a new report with create-only semantics. It never writes into the frozen baseline.

## Reproduce the PPTX failure without mutating the baseline

```powershell
& $node 'scripts\replay-pptx-baseline.mjs' $baseline "$evidence\pptx-default-env-50.json"
```

This makes 50 extract calls using the stored artifact references and the default Windows process encoding. It checks all 15 corpus hashes before and after execution, stores every request ID, duration and full DSH error text, and expects 10 failures each for `civil-service-line-management.pptx`, `timms-workshop.pptx` and `civil-society-covenant.pptx`; the remaining two decks must pass 10/10. It clears inherited `PYTHONIOENCODING` and `PYTHONUTF8` for this default-environment reproduction.

## Repeat the UTF-8 contrast

The UTF-8 contrast is intentionally separate from the default-environment replay:

```powershell
& $node 'scripts\probe-pptx-utf8.mjs' $baseline "$evidence\pptx-utf8-contrast.json"
```

It reads the frozen DSH artifact references, enables UTF-8 only for the child DSH environment, and writes to a new output file. It does not edit the source or workspace specimens.

## Repeat the native PDF boundary probe in an isolated replay root

This probe creates `workspace/pdf/nist-test-report.pdf` and DSH data artifacts, so use a new root containing only a copy of the frozen `sources` directory. Do not point it at the frozen baseline.

```powershell
$pdfRoot = Join-Path $evidence 'native-pdf-probe'
if (Test-Path -LiteralPath $pdfRoot) { throw "Probe path already exists: $pdfRoot" }
New-Item -ItemType Directory -Path $pdfRoot | Out-Null
Copy-Item -LiteralPath (Join-Path $baseline 'sources') -Destination (Join-Path $pdfRoot 'sources') -Recurse
& $node 'scripts\probe-native-pdf.mjs' $pdfRoot "$evidence\native-pdf-boundary.json"
```

The expected result is that DSH import succeeds but `docx-render.inspect` rejects the native PDF with `FORMAT_MISMATCH`. This tests the current capability boundary only; it is not a PDF parsing regression test.

## Repeat all 200 operations

The full run launches LibreOffice/Poppler 50 times and is expected to take tens of minutes. Create a new output root, copy only `sources`, and never run against the frozen baseline:

```powershell
$fullRoot = Join-Path 'dist' ('office-stress-replay-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
if (Test-Path -LiteralPath $fullRoot) { throw "Replay root already exists: $fullRoot" }
New-Item -ItemType Directory -Path $fullRoot | Out-Null
Copy-Item -LiteralPath (Join-Path $baseline 'sources') -Destination (Join-Path $fullRoot 'sources') -Recurse
Remove-Item Env:PYTHONIOENCODING -ErrorAction SilentlyContinue
Remove-Item Env:PYTHONUTF8 -ErrorAction SilentlyContinue
& $node 'scripts\office-stress-dsh.mjs' $fullRoot 10
& $node 'scripts\probe-pptx-utf8.mjs' $fullRoot (Join-Path $fullRoot 'utf8-contrast.json')
$pdfRoot = Join-Path $fullRoot 'native-pdf-probe'
New-Item -ItemType Directory -Path $pdfRoot | Out-Null
Copy-Item -LiteralPath (Join-Path $fullRoot 'sources') -Destination (Join-Path $pdfRoot 'sources') -Recurse
& $node 'scripts\probe-native-pdf.mjs' $pdfRoot (Join-Path $fullRoot 'native-pdf-boundary.json')
```

The runner emits 200 operation records (50 per DOCX, PDF-output, PPTX and XLSX path); imports, preflights and Doctor are counted separately. A full run is sequential; it is not a concurrency or semantic-correctness certification.

To exercise an explicitly selected source-built or local Profile rather than the default installed Profile, set `DSH_DOCX_PROFILE_ROOT`, `DSH_RUNTIME_ROOT` and optionally `DSH_PROFILE_LABEL` in the process environment before launching the runner. The manifest records both the profile label and root. Native PDF input reading is a separate capability and is covered by the `pdf-office` replay, not by the DOCX-to-PDF output lane here.

## B0 acceptance checklist

- [x] Corpus verifier exits 0 with exactly 15 specimens and all source/workspace hashes matching.
- [x] PPTX default-locale replay records 50 request IDs and reproduces the 30 encoding failures; all expected failures retain full DSH error text.
- [x] Hashes before and after focused replay are identical.
- [x] UTF-8 contrast and native-PDF boundary probes can write to caller-selected new files without overwriting prior evidence.
- [x] Full runner progress denominator now uses `plannedMaterialRounds` (200 for 10 repeats); the full render workload was not rerun in this batch.
- [x] Only test scripts and test documentation changed; no production module, package version, source specimen or frozen result was edited.

## B0 execution record

Executed on 2026-09-27 against the frozen run, without editing it:

- Corpus verification: PASS; all 15 source/workspace SHA-256 pairs match; the historical report contains 200 operation records and 231 DSH calls.
- Default Windows PPTX replay: PASS as a reproducer; 50/50 operations were recorded, with the expected 30 failures (10 each in the three affected decks) and 20 passes. All 15 corpus hashes were unchanged before/after; every call has a request ID and its full DSH error text is retained.
- UTF-8 contrast: PASS; the three affected decks all extracted successfully.
- Native PDF boundary: PASS; DSH import succeeded and `docx-render.inspect` returned the expected `FORMAT_MISMATCH`.
- Updated MJS scripts pass `node --check`. The full 200-operation workload was not repeated in B0; its existing 200-record/231-call report was hash-verified. Full rerun instructions remain available above.
- Production source files, installed profile package and frozen run directory were not modified.
