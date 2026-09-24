# Instalador de Claude para Word.
# Ejecutar UNA vez, con Word cerrado:
#   powershell -ExecutionPolicy Bypass -File .\install.ps1
#
# Hace:
#   1. Compila ClaudeWord.exe (helper oculto) con el csc.exe de .NET Framework.
#   2. Escribe config.json en %LOCALAPPDATA%\ClaudeWord.
#   3. Crea el workspace (rol/soul/contexto) si no existe.
#   4. Genera ClaudeWord.dotm con la macro, el atajo Alt+Enter y la opcion
#      "Enviar a Claude" del menu contextual, y la copia a la carpeta STARTUP de Word.
#
# Este archivo es ASCII a proposito: PowerShell 5.1 lee los .ps1 sin BOM como ANSI.

param(
    [string]$Workspace = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Claude-Word'),
    [string]$ClaudePath = '',
    [string]$Model = '',
    [string]$ClaudeConfigDir = ''
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$installDir = Join-Path $env:LOCALAPPDATA 'ClaudeWord'
$utf8 = New-Object System.Text.UTF8Encoding $false

function Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }

if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    throw 'Cerra Word antes de instalar.'
}

# --- Office -----------------------------------------------------------------
Step 'Office'
$c2r = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
if ($c2r) {
    Write-Host "    Version $($c2r.VersionToReport) - $($c2r.ProductReleaseIds)"
} else {
    Write-Host '    No es Click-to-Run (MSI). VBA + COM funciona igual en Office 2016 o superior.'
}

# --- Claude Code ------------------------------------------------------------
Step 'Claude Code'
if (-not $ClaudePath) {
    $cmd = Get-Command claude.exe -ErrorAction SilentlyContinue
    $ClaudePath = if ($cmd) { $cmd.Source } else { Join-Path $env:USERPROFILE '.local\bin\claude.exe' }
}
if (-not (Test-Path $ClaudePath)) {
    throw "No encuentro claude.exe ($ClaudePath). Instala Claude Code, inicia sesion con 'claude' y volve a correr este script."
}
Write-Host "    $ClaudePath"

# --- Helper -----------------------------------------------------------------
Step "Compilando helper en $installDir"
New-Item -ItemType Directory -Force -Path $installDir | Out-Null
$csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path $csc)) { throw "No encuentro csc.exe de .NET Framework 4 ($csc)." }

& $csc /nologo /target:winexe /codepage:65001 /optimize+ `
    "/out:$installDir\ClaudeWord.exe" `
    /r:System.Web.Extensions.dll /r:Microsoft.CSharp.dll /r:System.Core.dll `
    "$root\src\helper\ClaudeWord.cs"
if ($LASTEXITCODE -ne 0) { throw 'Fallo la compilacion del helper.' }

Copy-Item "$root\src\helper\word-rules.md" $installDir -Force

$config = [ordered]@{
    claudePath        = $ClaudePath
    workspace         = $Workspace
    model             = $Model
    claudeConfigDir   = $ClaudeConfigDir
    timeoutSeconds    = 600
    # Ritmo de escritura: palabra por palabra, como alguien tipeando.
    charsPerSecond    = 20
    startDelayMs      = 800
    sentencePauseMs   = 450
    paragraphPauseMs  = 900
    catchUpAfterChars = 600
}
# Al reinstalar se conserva el ritmo ajustado a mano.
$configPath = "$installDir\config.json"
if (Test-Path $configPath) {
    $old = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($k in 'timeoutSeconds', 'charsPerSecond', 'startDelayMs', 'sentencePauseMs', 'paragraphPauseMs', 'catchUpAfterChars') {
        if ($null -ne $old.$k) { $config[$k] = $old.$k }
    }
}
[IO.File]::WriteAllText($configPath, ($config | ConvertTo-Json), $utf8)

# --- Workspace --------------------------------------------------------------
Step "Workspace: $Workspace"
if (Test-Path $Workspace) {
    Write-Host '    Ya existe: no se toca.'
} else {
    Copy-Item "$root\workspace-template" $Workspace -Recurse
    Write-Host '    Creado desde workspace-template. Edita CLAUDE.md con el rol y el soul.'
}

# --- Plantilla de Word --------------------------------------------------------
Step 'Generando ClaudeWord.dotm'

# La importacion de modulos VBA exige CRLF y ANSI.
$bas = Join-Path $env:TEMP 'ClaudeWord.bas'
$src = [IO.File]::ReadAllText("$root\src\word\ClaudeWord.bas") -replace "`r?`n", "`r`n"
[IO.File]::WriteAllText($bas, $src, [Text.Encoding]::GetEncoding(1252))
$tmpDotm = Join-Path $env:TEMP 'ClaudeWord.dotm'
Remove-Item $tmpDotm -ErrorAction SilentlyContinue

# Acceso al modelo de objetos de VBA: se habilita solo durante la generacion.
$secKey = 'HKCU:\Software\Microsoft\Office\16.0\Word\Security'
if (-not (Test-Path $secKey)) { New-Item -Path $secKey | Out-Null }
$prevVbom = (Get-ItemProperty $secKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
Set-ItemProperty $secKey -Name AccessVBOM -Value 1 -Type DWord

$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $startup = $word.StartupPath

    $doc = $word.Documents.Add()
    $doc.VBProject.VBComponents.Import($bas) | Out-Null

    # Atajo y menu contextual se guardan DENTRO de la plantilla, no en Normal.dotm.
    # Los parametros opcionales de Word son ByRef: desde PowerShell van con [ref].
    $word.CustomizationContext = $doc
    # Alt + Enter (en Word: "Repetir", tambien en Ctrl+Y). Ctrl+Alt+Shift+letra choca
    # con atajos globales de drivers (AMD Adrenalin), que tienen prioridad sobre Word.
    $keyCode = $word.BuildKeyCode(13, [ref]1024)   # wdKeyReturn + wdKeyAlt
    $word.KeyBindings.Add(1, 'EnviarAClaude', $keyCode) | Out-Null

    $btn = $word.CommandBars.Item('Text').Controls.Add(1)   # msoControlButton
    $btn.Caption = 'Enviar a Claude'
    $btn.OnAction = 'EnviarAClaude'
    $btn.BeginGroup = $true

    # Late binding: con [ref] PowerShell pasa el string envuelto en PSObject y COM lo rechaza.
    $saveArgs = [object[]]@([string]$tmpDotm, 15)   # 15 = wdFormatXMLTemplateMacroEnabled
    [void]$doc.GetType().InvokeMember('SaveAs2', [Reflection.BindingFlags]::InvokeMethod, $null, $doc, $saveArgs)
    $doc.Close([ref]0)
}
finally {
    # Cada limpieza por separado: si una falla, las demas corren igual.
    if ($word) {
        try { $word.Quit([ref]0) } catch { Write-Warning "No pude cerrar Word oculto: $_" }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch { }
    }
    try {
        if ($null -eq $prevVbom) {
            Remove-ItemProperty $secKey -Name AccessVBOM -ErrorAction SilentlyContinue
        } else {
            Set-ItemProperty $secKey -Name AccessVBOM -Value $prevVbom -Type DWord
        }
    } catch { Write-Warning "No pude restaurar AccessVBOM: $_" }
    Remove-Item $bas -ErrorAction SilentlyContinue
}

if (-not (Test-Path $startup)) { New-Item -ItemType Directory -Force -Path $startup | Out-Null }
Copy-Item $tmpDotm (Join-Path $startup 'ClaudeWord.dotm') -Force
Remove-Item $tmpDotm -ErrorAction SilentlyContinue
Write-Host "    $startup\ClaudeWord.dotm"

Write-Host ''
Write-Host 'Listo. Abri Word, pone el cursor en un parrafo y presiona Alt+Enter' -ForegroundColor Green
Write-Host '(o clic derecho > Enviar a Claude). Log del ultimo pedido:'
Write-Host "    $installDir\logs\last.log"
