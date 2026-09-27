# 从上游逐件下载并校验运行时（Windows x64）。
#
# 注意：这份**配方的家已经搬到 sdaniasdsd/dsh-toolchain**（scripts/fetch-runtime.ps1）。
# 本仓库的常规路径是 `npm run fetch:toolchain` —— 按 toolchain.lock.json 钉住的
# 工具链工件下载解包，不直接打上游。
# 这个脚本保留给"要从上游重建一份运行时"的场合（例如工具链仓库自己发新版时）。
# 它按旧布局写到 dist/dsh-docx/runtime/win32-x64，build-dsh.mjs 仍认这个位置（会改名搬进运行时包）。
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
function Expand-Wheel($wheelPath) {
    $archivePath = Join-Path $cache ((Split-Path -Leaf $wheelPath) + '.zip')
    Copy-Item -LiteralPath $wheelPath -Destination $archivePath -Force
    Expand-Archive -LiteralPath $archivePath -DestinationPath $sitePackages -Force
    Remove-Item -LiteralPath $archivePath
}
$python = Fetch-Verified 'https://www.python.org/ftp/python/3.13.15/python-3.13.15-embed-amd64.zip' 'python-3.13.15.zip' 'd1f04d990aee1253d8569e8e5104e30fa9f5fa830899f14843448872d936a2cf'
$poppler = Fetch-Verified 'https://github.com/oschwartz10612/poppler-windows/releases/download/v26.09.0-0/Release-26.09.0-0.zip?download=1' 'poppler-26.09.0.zip' '7a6f256a0ddf7536182246a5733331bf4677cbcc34f4663774947ad34556c8d0'
$office = Fetch-Verified 'https://mirror.clarkson.edu/tdf/libreoffice/stable/26.8.0/win/x86_64/LibreOffice_26.8.0_Win_x86-64.msi' 'LibreOffice_26.8.0_Win_x86-64.msi' '4aa6c6e1895f4055104effcb556bd3362d20c6ad707c149543304f395ef9db95'
if (!(Test-Path -LiteralPath (Join-Path $runtime 'python/python.exe'))) { Expand-Archive -LiteralPath $python -DestinationPath (Join-Path $runtime 'python') }
if (!(Test-Path -LiteralPath (Join-Path $runtime 'poppler/poppler-26.09.0/Library/bin/pdftoppm.exe'))) {
    New-Item -ItemType Directory -Force -Path (Join-Path $runtime 'poppler') | Out-Null
    Expand-Archive -LiteralPath $poppler -DestinationPath (Join-Path $runtime 'poppler') -Force
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
    Expand-Wheel $wheel
    $dist=Get-ChildItem -LiteralPath $sitePackages -Directory -Filter 'lxml-*.dist-info' | Select-Object -First 1
    if ($dist -and $dist.Name -ne 'lxml-6.1.0.dist-info') { Move-Item -LiteralPath $dist.FullName -Destination (Join-Path $sitePackages 'lxml-6.1.0.dist-info') }
}
$pptxWheels = @(
    @{ Url='https://files.pythonhosted.org/packages/d9/4f/00be2196329ebbff56ce564aa94efb0fbc828d00de250b1980de1a34ab49/python_pptx-1.0.2-py3-none-any.whl'; Name='python_pptx-1.0.2-py3-none-any.whl'; Sha256='160838e0b8565a8b1f67947675886e9fea18aa5e795db7ae531606d68e785cba' },
    @{ Url='https://files.pythonhosted.org/packages/a6/9b/7a58e61d62be561da3a356fe2384d4059a6345fc130e23ef1c36a5b81d24/pillow-12.3.0-cp313-cp313-win_amd64.whl'; Name='pillow-12.3.0-cp313-cp313-win_amd64.whl'; Sha256='1cca606cd25738df4ed873d5ad46bbdb3d83b5cbca291f6b4ff13a4df6b0bbe8' },
    @{ Url='https://files.pythonhosted.org/packages/3a/0c/3662f4a66880196a590b202f0db82d919dd2f89e99a27fadef91c4a33d41/xlsxwriter-3.2.9-py3-none-any.whl'; Name='xlsxwriter-3.2.9-py3-none-any.whl'; Sha256='9a5db42bc5dff014806c58a20b9eae7322a134abb6fce3c92c181bfb275ec5b3' },
    @{ Url='https://files.pythonhosted.org/packages/49/d3/b8441a820a491ddfc024b0b0cf0393375b75ea13866d9c66727e54c2fc80/typing_extensions-4.16.0-py3-none-any.whl'; Name='typing_extensions-4.16.0-py3-none-any.whl'; Sha256='481caa481374e813c1b176ada14e97f1f67a4539ce9cfeb3f350d78d6370c2e8' }
)
New-Item -ItemType Directory -Force -Path $sitePackages | Out-Null
foreach ($item in $pptxWheels) {
    $wheelPath = Fetch-Verified $item.Url $item.Name $item.Sha256
    Expand-Wheel $wheelPath
}
if (!(Test-Path -LiteralPath (Join-Path $runtime 'python/python313._pth'))) { throw 'Python embeddable path file is missing.' }
# 只列解释器自身的搜索路径。引擎目录【不能】写在这里：._pth 的相对路径是相对
# python.exe 所在目录解析的，而引擎随「核心包」走（node_modules/@deepseek-ai/dsh-docx/lib/engines），
# 与「运行时包」是两棵不同的树，任何写死的相对路径都会在换一种安装布局后失效。
# 引擎需要同目录的兄弟模块时，由引擎自己 sys.path.insert。
$pth=@('python313.zip','.','Lib/site-packages','import site')
Set-Content -LiteralPath (Join-Path $runtime 'python/python313._pth') -Value $pth -Encoding ascii
Write-Host "Offline runtimes are hash-verified and extracted under $runtime"
