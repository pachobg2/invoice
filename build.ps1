# Builds the single-file app:
#   1. Lite edition (Windows XP..11, .NET 2.0)       -> host\lite-bin\   (gets embedded in the main exe)
#   2. native host (self-contained, web app embedded) -> invoice.exe
# Usage:  powershell -ExecutionPolicy Bypass -File build.ps1   (close the running app first - it locks the exe)
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$out  = Join-Path $root 'dist'

# 1. Lite edition
dotnet build (Join-Path $root 'lite') -c Release
if ($LASTEXITCODE -ne 0) { throw 'lite build failed' }
$bin = Join-Path $root 'host\lite-bin'
New-Item -ItemType Directory -Force $bin | Out-Null
Get-ChildItem $bin -File | ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force }   # drop files from older names
Copy-Item (Join-Path $root 'lite\bin\Release\net20\invoice-lite.exe') $bin -Force
Copy-Item (Join-Path $root 'lite\bin\Release\net20\invoice-lite.exe.config') $bin -Force

# 2. main app
dotnet publish (Join-Path $root 'host') -c Release -r win-x64 --self-contained true `
    -p:PublishSingleFile=true -p:EnableCompressionInSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true `
    -o $out
if ($LASTEXITCODE -ne 0) { throw 'publish failed' }

Copy-Item (Join-Path $out 'Faktura.exe') (Join-Path $root 'invoice.exe') -Force
Remove-Item (Join-Path $root 'host\bin'), (Join-Path $root 'host\obj'), (Join-Path $root 'lite\bin'), (Join-Path $root 'lite\obj'), $out -Recurse -Force
'{0}: {1:n1} MB' -f 'invoice.exe', ((Get-Item (Join-Path $root 'invoice.exe')).Length / 1MB)
