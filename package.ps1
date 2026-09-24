# Genera dist\ClaudeWord-instalador.zip con lo necesario para instalar en otra PC.
# El workspace personal NO se incluye: solo la plantilla generica.
#   powershell -ExecutionPolicy Bypass -File .\package.ps1

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$dist = Join-Path $root 'dist'
$stage = Join-Path $dist 'ClaudeWord'
$zip = Join-Path $dist 'ClaudeWord-instalador.zip'

Remove-Item $stage -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $zip -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path $stage | Out-Null

$items = 'instalar.cmd', 'desinstalar.cmd', 'install.ps1', 'uninstall.ps1',
         'INSTALAR.md', 'README.md', 'src', 'workspace-template'
foreach ($item in $items) {
    Copy-Item (Join-Path $root $item) $stage -Recurse
}

Compress-Archive -Path $stage -DestinationPath $zip
Remove-Item $stage -Recurse -Force
Write-Host "Listo: $zip" -ForegroundColor Green
