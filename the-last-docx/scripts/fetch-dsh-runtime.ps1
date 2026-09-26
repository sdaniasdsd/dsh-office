$ErrorActionPreference = 'Stop'
$pluginRoot = Split-Path $PSScriptRoot -Parent
$cache = Join-Path $pluginRoot '.build-cache'
$runtime = Join-Path $pluginRoot 'dist/dsh-docx/runtime/win32-x64'
New-Item -ItemType Directory -Force -Path $cache,$runtime | Out-Null
function Fetch-Verified($url, $name, $sha256) {
    $target = Join-Path $cache $name
    $partial = "$target.part"
    if (!(Test-Path -LiteralPath $target) -or (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash.ToLowerInvariant() -ne $sha256) {
        if (Test-Path -LiteralPath $target) { Move-Item -LiteralPath $target -Destination "$target.invalid.$([DateTime]::UtcNow.ToString('yyyyMMddHHmmss'))" }
        & curl.exe --fail --location --ssl-revoke-best-effort --retry 3 --retry-all-errors --continue-at - --connect-timeout 10 --max-time 600 --silent --show-error --output $partial $url
        if ($LASTEXITCODE -ne 0) { throw "Download failed: $name" }
        Move-Item -LiteralPath $partial -Destination $target
    }
    if ((Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash.ToLowerInvariant() -ne $sha256) { throw "Checksum mismatch: $name" }
    Write-Host "Verified $name"
    return $target
}
$python = Fetch-Verified 'https://www.python.org/ftp/python/3.13.15/python-3.13.15-embed-amd64.zip' 'python-3.13.15.zip' 'd1f04d990aee1253d8569e8e5104e30fa9f5fa830899f14843448872d936a2cf'
$poppler = Fetch-Verified 'https://github.com/oschwartz10612/poppler-windows/releases/download/v26.09.0-0/Release-26.09.0-0.zip?download=1' 'poppler-26.09.0.zip' '7a6f256a0ddf7536182246a5733331bf4677cbcc34f4663774947ad34556c8d0'
$office = Fetch-Verified 'https://mirror.clarkson.edu/tdf/libreoffice/stable/26.8.0/win/x86_64/LibreOffice_26.8.0_Win_x86-64.msi' 'LibreOffice_26.8.0_Win_x86-64.msi' '4aa6c6e1895f4055104effcb556bd3362d20c6ad707c149543304f395ef9db95'
if (!(Test-Path -LiteralPath (Join-Path $runtime 'python/python.exe'))) { Expand-Archive -LiteralPath $python -DestinationPath (Join-Path $runtime 'python') }
if (!(Test-Path -LiteralPath (Join-Path $runtime 'poppler/poppler-26.09.0/Library/bin/pdftoppm.exe'))) {
    $sevenZipPackage=Join-Path $cache '7zip-package'
    if (!(Test-Path -LiteralPath (Join-Path $sevenZipPackage 'package/win/x64/7za.exe'))) { New-Item -ItemType Directory -Force -Path $sevenZipPackage | Out-Null; tar.exe -xzf (Join-Path $cache '7zip-bin-5.2.0.tgz') -C $sevenZipPackage }
    $sevenZip=Join-Path $sevenZipPackage 'package/win/x64/7za.exe'
    New-Item -ItemType Directory -Force -Path (Join-Path $runtime 'poppler') | Out-Null
    & $sevenZip x $poppler "-o$(Join-Path $runtime 'poppler')" -y | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'Poppler extraction failed' }
}
$sofficePortable=Join-Path $runtime 'libreoffice/program/soffice.com'
if (!(Test-Path -LiteralPath $sofficePortable)) {
    $officeDir=Join-Path $runtime 'libreoffice'
    New-Item -ItemType Directory -Force -Path $officeDir | Out-Null
    $arguments="/a `"$office`" /qn /norestart TARGETDIR=`"$officeDir`""
    $extractor=Start-Process -FilePath 'msiexec.exe' -ArgumentList $arguments -Wait -WindowStyle Hidden -PassThru
    if ($extractor.ExitCode -ne 0 -or !(Test-Path -LiteralPath $sofficePortable)) { throw "LibreOffice MSI administrative extraction failed (exit $($extractor.ExitCode))." }
}
$installerPayload=Join-Path $runtime 'libreoffice/LibreOffice_26.8.0_Win_x86-64.msi'
if (Test-Path -LiteralPath $installerPayload) { Remove-Item -LiteralPath $installerPayload }
$wheel=Fetch-Verified 'https://files.pythonhosted.org/packages/8e/63/981401c5680c1eb30893f00a19641ac80db5d1e7086c62cb4b13ed813038/lxml-6.1.0-cp313-cp313-win_amd64.whl' 'lxml-6.1.0-cp313-cp313-win_amd64.whl' '4a1503c56e4e2b38dc76f2f2da7bae69670c0f1933e27cfa34b2fa5876410b16'
$sitePackages=Join-Path $runtime 'python/Lib/site-packages'
if (!(Test-Path -LiteralPath (Join-Path $sitePackages 'lxml/__init__.py'))) {
    New-Item -ItemType Directory -Force -Path $sitePackages | Out-Null
    $sevenZipPackage=Join-Path $cache '7zip-package';$sevenZip=Join-Path $sevenZipPackage 'package/win/x64/7za.exe'
    & $sevenZip x $wheel "-o$sitePackages" -y | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'lxml extraction failed' }
    $dist=Get-ChildItem -LiteralPath $sitePackages -Directory -Filter 'lxml-*.dist-info' | Select-Object -First 1
    if ($dist -and $dist.Name -ne 'lxml-6.1.0.dist-info') { Move-Item -LiteralPath $dist.FullName -Destination (Join-Path $sitePackages 'lxml-6.1.0.dist-info') }
}
if (!(Test-Path -LiteralPath (Join-Path $runtime 'python/python313._pth'))) { throw 'Python embeddable path file is missing.' }
$pth=@('python313.zip','.','Lib/site-packages','../../../lib/engines/docx-complex-parse','import site')
Set-Content -LiteralPath (Join-Path $runtime 'python/python313._pth') -Value $pth -Encoding ascii
Write-Host "Offline runtimes are hash-verified and extracted under $runtime"
