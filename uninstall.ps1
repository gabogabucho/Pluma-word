# Desinstala Pluma para Word. El workspace (rol, tareas, agentes) NO se borra.
#   powershell -ExecutionPolicy Bypass -File .\uninstall.ps1

$ErrorActionPreference = 'Stop'

if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    throw 'Cerra Word antes de desinstalar.'
}

$word = New-Object -ComObject Word.Application
try { $startup = $word.StartupPath }
finally {
    $word.Quit([ref]0)
    [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}

Remove-Item (Join-Path $startup 'Pluma.dotm') -ErrorAction SilentlyContinue
Remove-Item (Join-Path $startup 'ClaudeWord.dotm') -ErrorAction SilentlyContinue
Remove-Item (Join-Path $env:LOCALAPPDATA 'Pluma') -Recurse -Force -ErrorAction SilentlyContinue

Write-Host 'Pluma desinstalado. El workspace quedo intacto.' -ForegroundColor Green
