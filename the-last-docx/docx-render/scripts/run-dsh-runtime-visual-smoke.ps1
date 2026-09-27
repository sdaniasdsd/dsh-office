param(
    [Parameter(Mandatory = $true)]
    [string]$RuntimeRoot,
    [string]$OutputRoot = (Join-Path $PSScriptRoot '..\artifacts\runtime-visual-smoke')
)

$ErrorActionPreference = 'Stop'
$runtime = (Resolve-Path -LiteralPath $RuntimeRoot).Path
$soffice = Join-Path $runtime 'libreoffice\program\soffice.com'
$pdftoppm = Join-Path $runtime 'poppler\poppler-26.09.0\Library\bin\pdftoppm.exe'
foreach ($executable in @($soffice, $pdftoppm)) {
    if (-not (Test-Path -LiteralPath $executable -PathType Leaf)) {
        throw "DSH runtime executable is missing: $executable"
    }
}

$workspace = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$createRoot = Join-Path $workspace 'docx-create\fixtures\_generated'
$fixtureDir = Get-ChildItem -LiteralPath $createRoot -Directory |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
if ($null -eq $fixtureDir) { throw "No generated docx-create fixtures at $createRoot" }

$sources = @(
    (Join-Path $fixtureDir.FullName 'technical-sample.docx'),
    (Join-Path $fixtureDir.FullName 'report-sample.docx'),
    (Join-Path $fixtureDir.FullName 'chinese-long-sample.docx'),
    (Join-Path $workspace 'DSH-Office办公室文件强度测试与分批修复报告.docx')
)
foreach ($source in $sources) {
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Visual smoke source is missing: $source" }
}

$runRoot = Join-Path $OutputRoot (Get-Date -Format 'yyyyMMdd-HHmmss')
$pdfRoot = Join-Path $runRoot 'pdf'
$pagesRoot = Join-Path $runRoot 'pages'
New-Item -ItemType Directory -Force -Path $pdfRoot, $pagesRoot | Out-Null

$ledger = @()
foreach ($source in $sources) {
    $sourcePath = (Resolve-Path -LiteralPath $source).Path
    $name = [IO.Path]::GetFileNameWithoutExtension($sourcePath)
    $profile = Join-Path $runRoot ("lo-profile-" + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path $profile | Out-Null
    $profileUri = [Uri]::new($profile).AbsoluteUri
    $lo = Start-Process -FilePath $soffice -ArgumentList @(
        '--headless', '--nologo', '--nodefault', '--norestore',
        "-env:UserInstallation=$profileUri",
        '--convert-to', 'pdf', '--outdir', $pdfRoot, $sourcePath
    ) -Wait -PassThru -WindowStyle Hidden
    if ($lo.ExitCode -ne 0) { throw "LibreOffice conversion failed for $sourcePath (exit $($lo.ExitCode))." }

    $pdf = Join-Path $pdfRoot "$name.pdf"
    if (-not (Test-Path -LiteralPath $pdf -PathType Leaf)) { throw "LibreOffice reported success without PDF: $pdf" }
    $pageDir = Join-Path $pagesRoot $name
    New-Item -ItemType Directory -Force -Path $pageDir | Out-Null
    $prefix = Join-Path $pageDir 'page'
    $poppler = Start-Process -FilePath $pdftoppm -ArgumentList @('-png', '-r', '144', $pdf, $prefix) -Wait -PassThru -WindowStyle Hidden
    if ($poppler.ExitCode -ne 0) { throw "Poppler conversion failed for $pdf (exit $($poppler.ExitCode))." }
    # The disposable office profile avoids cross-document state, but is not QA evidence.
    Remove-Item -LiteralPath $profile -Recurse -Force
    $pages = @(Get-ChildItem -LiteralPath $pageDir -Filter 'page-*.png' -File | Sort-Object Name)
    if ($pages.Count -eq 0) { throw "Poppler reported success without PNG pages: $pdf" }
    $ledger += [PSCustomObject]@{
        source = $sourcePath
        sourceSha256 = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
        pdf = $pdf
        pdfSha256 = (Get-FileHash -LiteralPath $pdf -Algorithm SHA256).Hash
        pageCount = $pages.Count
        pages = @($pages | ForEach-Object {
            [PSCustomObject]@{ path = $_.FullName; sha256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash; bytes = $_.Length }
        })
    }
}

$ledgerPath = Join-Path $runRoot 'manifest.json'
([PSCustomObject]@{
    runtimeRoot = $runtime
    runtime = @{
        soffice = $soffice
        pdftoppm = $pdftoppm
    }
    dpi = 144
    generatedAt = (Get-Date).ToUniversalTime().ToString('o')
    cases = $ledger
} | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $ledgerPath -Encoding utf8
Write-Output "Visual smoke complete: $runRoot"
Write-Output "Manifest: $ledgerPath"

