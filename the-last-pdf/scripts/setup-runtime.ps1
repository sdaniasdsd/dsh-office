param([switch]$WithDocling, [string]$Python = 'python')
$ErrorActionPreference = 'Stop'
$pdfWorkspace = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$pdfRuntime = Join-Path $pdfWorkspace 'runtime'
New-Item -ItemType Directory -Path $pdfRuntime -Force | Out-Null
$qpdfArchive = Join-Path $pdfRuntime 'qpdf-12.4.1-msvc64.zip'
$qpdfExpected = '3cd016cd433ef7232e42f4c13348a49cc14907a3c7278ef4f99120593126f7a6'
if (-not (Test-Path -LiteralPath $qpdfArchive)) {
  Invoke-WebRequest -Uri 'https://github.com/qpdf/qpdf/releases/download/v12.4.1/qpdf-12.4.1-msvc64.zip' -OutFile $qpdfArchive
}
if ((Get-FileHash -LiteralPath $qpdfArchive -Algorithm SHA256).Hash.ToLowerInvariant() -ne $qpdfExpected) { throw 'qpdf archive checksum mismatch' }
$qpdfDirectory = Join-Path $pdfRuntime 'qpdf-12.4.1'
$qpdfExecutable = Join-Path $qpdfDirectory 'qpdf-12.4.1-msvc64/bin/qpdf.exe'
if (-not (Test-Path -LiteralPath $qpdfExecutable)) { Expand-Archive -LiteralPath $qpdfArchive -DestinationPath $qpdfDirectory }
& $qpdfExecutable --version
if ($LASTEXITCODE -ne 0) { throw 'qpdf failed to start' }
if ($WithDocling) {
  $pdfVenv = Join-Path $pdfWorkspace '.venv'
  $pdfPython = Join-Path $pdfVenv 'Scripts/python.exe'
  if (-not (Test-Path -LiteralPath $pdfPython)) {
    & $Python -m venv $pdfVenv
    if ($LASTEXITCODE -ne 0) { throw 'venv creation failed' }
  }
  & $pdfPython -c 'import sys; print(sys.version); sys.exit(0 if sys.version_info[:2] == (3, 11) else 1)'
  if ($LASTEXITCODE -ne 0) { throw 'The pinned Docling runtime in this workspace requires Python 3.11.' }
  & $pdfPython -m pip install -r (Join-Path $pdfWorkspace 'requirements-docling.txt')
  if ($LASTEXITCODE -ne 0) { throw 'Docling installation failed' }
  & $pdfPython (Join-Path $PSScriptRoot 'download-models.py')
  if ($LASTEXITCODE -ne 0) { throw 'Model provisioning failed' }
  & $pdfPython (Join-Path $PSScriptRoot 'inventory.py')
  if ($LASTEXITCODE -ne 0) { throw 'Inventory failed' }
}
