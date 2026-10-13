# Builds and publishes a GitHub release (tag vX.Y.Z from host\Faktura.csproj) with invoice.exe, invoice-lite.exe and SHA256SUMS.txt.
# Usage:  powershell -ExecutionPolicy Bypass -File release.ps1      (commit your changes first; close the running app)
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
Set-Location $root

$ver = ([xml](Get-Content (Join-Path $root 'host\Faktura.csproj') -Raw)).Project.PropertyGroup.Version | Where-Object { $_ } | Select-Object -First 1
if ($ver -notmatch '^\d+\.\d+\.\d+$') { throw "bad version '$ver' in host\Faktura.csproj" }
$tag = "v$ver"

# release notes = first CHANGELOG entry of the web app; it must belong to this version
$js = Get-Content (Join-Path $root 'faktura\app.js') -Raw -Encoding UTF8
$m = [regex]::Match($js, "const CHANGELOG = \[\s*\['([^']+)',\s*'([^']+)',\s*\[(.*?)\]\]")
if (-not $m.Success -or $m.Groups[1].Value -ne $ver) { throw "the first CHANGELOG entry in faktura\app.js is not $ver" }
$items = [regex]::Matches($m.Groups[3].Value, "'((?:[^'\\]|\\.)*)'") | ForEach-Object { '- ' + $_.Groups[1].Value }
$notes = ($items -join "`n")

if (git status --porcelain) { throw 'uncommitted changes - commit first' }
if (git tag -l $tag) { throw "tag $tag already exists - bump the version" }

& (Join-Path $root 'build.ps1')
if ($LASTEXITCODE -ne 0) { throw 'build failed' }

$stage = Join-Path $env:TEMP ("invoice-release-" + $ver)
New-Item -ItemType Directory -Force $stage | Out-Null
Copy-Item (Join-Path $root 'invoice.exe') $stage -Force
Copy-Item (Join-Path $root 'host\lite-bin\invoice-lite.exe') $stage -Force
$sums = foreach ($n in 'invoice.exe', 'invoice-lite.exe') { '{0} *{1}' -f (Get-FileHash (Join-Path $stage $n) -Algorithm SHA256).Hash.ToLower(), $n }
[IO.File]::WriteAllText((Join-Path $stage 'SHA256SUMS.txt'), (($sums -join "`n") + "`n"), (New-Object Text.UTF8Encoding($false)))

git tag $tag
git push origin HEAD
git push origin $tag
$notesFile = Join-Path $stage 'notes.md'
[IO.File]::WriteAllText($notesFile, $notes, (New-Object Text.UTF8Encoding($false)))
gh release create $tag (Join-Path $stage 'invoice.exe') (Join-Path $stage 'invoice-lite.exe') (Join-Path $stage 'SHA256SUMS.txt') --repo pachobg2/invoice --title $tag --notes-file $notesFile
if ($LASTEXITCODE -ne 0) { throw 'gh release failed' }
"released $tag"
