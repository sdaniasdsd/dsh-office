param(
    [Parameter(Mandatory = $true)]
    [string]$RuntimeRoot,
    [string]$OutputRoot = (Join-Path $PSScriptRoot '..\artifacts\cool-corporate-visual-smoke')
)

$ErrorActionPreference = 'Stop'
$runtime = (Resolve-Path -LiteralPath $RuntimeRoot).Path
$python = Join-Path $runtime 'python\python.exe'
$soffice = Join-Path $runtime 'libreoffice\program\soffice.com'
$pdftoppm = Join-Path $runtime 'poppler\poppler-26.09.0\Library\bin\pdftoppm.exe'
foreach ($executable in @($python, $soffice, $pdftoppm)) {
    if (-not (Test-Path -LiteralPath $executable -PathType Leaf)) { throw "DSH runtime executable is missing: $executable" }
}

$moduleRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$runRoot = Join-Path $OutputRoot (Get-Date -Format 'yyyyMMdd-HHmmss')
New-Item -ItemType Directory -Force -Path $runRoot | Out-Null
$input = Join-Path $runRoot 'input.pptx'
$styled = Join-Path $runRoot 'cool-corporate.pptx'
$pdfRoot = Join-Path $runRoot 'pdf'
$pagesRoot = Join-Path $runRoot 'pages'
New-Item -ItemType Directory -Force -Path $pdfRoot, $pagesRoot | Out-Null

$pythonCode = 'from pptx import Presentation; import sys; p=Presentation(); s=p.slides.add_slide(p.slide_layouts[0]); s.shapes.title.text="Quarterly Review"; s.placeholders[1].text="Evidence and decisions"; p.save(sys.argv[1])'
& $python -c $pythonCode $input
if ($LASTEXITCODE -ne 0) { throw 'Could not create visual-smoke input presentation.' }

$request = @{ limits = @{ maxArchiveEntries = 4096; maxEntryUncompressedBytes = 67108864; maxTotalUncompressedBytes = 268435456; maxSlides = 500; maxShapes = 100000; maxTextChars = 2000000 }; styleId = 'cool-corporate-field-v1'; artWord = 'Q3' } | ConvertTo-Json -Compress
$bridge = Join-Path $moduleRoot 'src\engine\pptx_bridge.py'
$bridgeResult = $request | & $python $bridge applyArtStyle $input $styled
if ($LASTEXITCODE -ne 0) { throw 'Cool-corporate profile application failed.' }

$profile = Join-Path $runRoot ('lo-profile-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $profile | Out-Null
$profileUri = [Uri]::new($profile).AbsoluteUri
$lo = Start-Process -FilePath $soffice -ArgumentList @('--headless', '--nologo', '--nodefault', '--norestore', "-env:UserInstallation=$profileUri", '--convert-to', 'pdf', '--outdir', $pdfRoot, $styled) -Wait -PassThru -WindowStyle Hidden
if ($lo.ExitCode -ne 0) { throw "LibreOffice conversion failed (exit $($lo.ExitCode))." }
Remove-Item -LiteralPath $profile -Recurse -Force

$pdf = Join-Path $pdfRoot 'cool-corporate.pdf'
if (-not (Test-Path -LiteralPath $pdf -PathType Leaf)) { throw 'LibreOffice did not produce the expected PDF.' }
$poppler = Start-Process -FilePath $pdftoppm -ArgumentList @('-png', '-r', '144', $pdf, (Join-Path $pagesRoot 'slide')) -Wait -PassThru -WindowStyle Hidden
if ($poppler.ExitCode -ne 0) { throw "Poppler conversion failed (exit $($poppler.ExitCode))." }
$pages = @(Get-ChildItem -LiteralPath $pagesRoot -Filter 'slide-*.png' -File | Sort-Object Name)
if ($pages.Count -ne 1) { throw "Expected one visual-smoke PNG, received $($pages.Count)." }

@{ runtimeRoot = $runtime; profile = 'cool-corporate-field-v1'; bridge = $bridgeResult | ConvertFrom-Json; pptx = $styled; pdf = $pdf; page = $pages[0].FullName } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $runRoot 'manifest.json') -Encoding utf8
Write-Output "Design-profile visual smoke complete: $runRoot"
