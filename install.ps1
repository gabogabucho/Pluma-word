# Instalador de Pluma para Word.
# Ejecutar UNA vez, con Word cerrado:
#   powershell -ExecutionPolicy Bypass -File .\install.ps1
#
# Hace:
#   1. Detecta los CLI instalados (claude, codex, gemini, grok, hermes, aider).
#   2. Compila Pluma.exe con el csc.exe de .NET Framework.
#   3. Escribe config.json en %LOCALAPPDATA%\Pluma.
#   4. Crea el workspace si no existe, y agents.json con los CLI encontrados.
#   5. Genera Pluma.dotm (macro, Alt+Enter, menu contextual) y lo copia a STARTUP.
#
# Este archivo es ASCII a proposito: PowerShell 5.1 lee los .ps1 sin BOM como ANSI.

param(
    [string]$Workspace = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Pluma'),
    [string]$ClaudePath = '',
    [string]$Model = '',
    [string]$ClaudeConfigDir = '',
    [string]$DefaultAgent = ''
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$installDir = Join-Path $env:LOCALAPPDATA 'Pluma'
$utf8 = New-Object System.Text.UTF8Encoding $false
$prefer = @('claude', 'codex', 'gemini', 'grok', 'hermes', 'aider')

function Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }

function Add-PlumaButton($parent, $caption, $action, $group) {
    $b = $parent.Controls.Add(1)   # msoControlButton
    $b.Caption = $caption
    $b.OnAction = $action
    if ($group) { $b.BeginGroup = $true }
}

function Find-Cli([string]$Name, [string[]]$Extra) {
    $list = @()
    $exe = Get-Command ($Name + '.exe') -ErrorAction SilentlyContinue
    if ($exe -and $exe.Source) { $list += $exe.Source }
    $any = Get-Command $Name -ErrorAction SilentlyContinue
    if ($any -and $any.Source) { $list += $any.Source }
    if ($Extra) { $list += $Extra }
    foreach ($c in $list) {
        if ($c -and (Test-Path -LiteralPath $c)) {
            return (Resolve-Path -LiteralPath $c).Path
        }
    }
    return $null
}

function Quote-Json([string]$s) {
    if ($null -eq $s) { $s = '' }
    $s = $s.Replace('\', '\\').Replace('"', '\"').Replace("`r", '\r').Replace("`n", '\n').Replace("`t", '\t')
    return '"' + $s + '"'
}

function Format-Json([object]$v, [int]$indent, [string[]]$preferKeys) {
    $pad = ' ' * $indent
    $pad2 = ' ' * ($indent + 2)
    if ($null -eq $v) { return 'null' }
    if ($v -is [bool]) { if ($v) { return 'true' } else { return 'false' } }
    if ($v -is [string]) { return (Quote-Json $v) }
    if ($v -is [System.Collections.IDictionary]) {
        $keys = @($v.Keys)
        $ordered = @()
        foreach ($k in $preferKeys) { if ($keys -contains $k) { $ordered += $k } }
        foreach ($k in $keys) { if ($ordered -notcontains $k) { $ordered += $k } }
        $parts = @()
        foreach ($key in $ordered) {
            $childPrefer = @()
            if ($key -eq 'agents') { $childPrefer = @('id','enabled','label','command','note','args','promptVia','promptFormat','promptArg','parser','stream','images','imageArg','maxImages','modelArg','rules','rulesArg','context','env') }
            $parts += ($pad2 + (Quote-Json ([string]$key)) + ': ' + (Format-Json $v[$key] ($indent + 2) $childPrefer))
        }
        if ($parts.Count -eq 0) { return '{}' }
        return "{`r`n" + ($parts -join ",`r`n") + "`r`n" + $pad + '}'
    }
    if ($v -is [System.Collections.IEnumerable]) {
        $parts = @()
        foreach ($item in $v) {
            $parts += ($pad2 + (Format-Json $item ($indent + 2) @('id','enabled','label','command','note','args','promptVia','promptFormat','promptArg','parser','stream','images','imageArg','maxImages','modelArg','rules','rulesArg','context','env')))
        }
        if ($parts.Count -eq 0) { return '[]' }
        return "[`r`n" + ($parts -join ",`r`n") + "`r`n" + $pad + ']'
    }
    return ([string]$v).Replace(',', '.')
}

function Write-DetectedAgents([string]$templatePath, [string]$destPath, $found, [string]$defaultId) {
    Add-Type -AssemblyName System.Web.Extensions
    $ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
    $ser.MaxJsonLength = 16777216
    $rootObj = $ser.DeserializeObject([IO.File]::ReadAllText($templatePath, $utf8))
    $rootObj['default'] = $defaultId
    foreach ($a in $rootObj['agents']) {
        $id = [string]$a['id']
        if ($found.ContainsKey($id) -and $found[$id]) {
            $a['command'] = [string]$found[$id]
            $a['enabled'] = $true
        } else {
            $a['enabled'] = $false
        }
    }
    $preferKeys = @('default', 'agents')
    [IO.File]::WriteAllText($destPath, (Format-Json $rootObj 0 $preferKeys) + "`r`n", $utf8)
}

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

# --- Agentes ----------------------------------------------------------------
Step 'Agentes CLI'
$extras = @{
    claude = @(
        (Join-Path $env:USERPROFILE '.local\bin\claude.exe'),
        (Join-Path $env:USERPROFILE '.claude\local\claude.exe')
    )
    codex = @(
        (Join-Path $env:USERPROFILE '.local\bin\codex.exe'),
        (Join-Path $env:APPDATA 'npm\codex.cmd')
    )
    gemini = @(
        (Join-Path $env:APPDATA 'npm\gemini.cmd'),
        (Join-Path $env:USERPROFILE '.local\bin\gemini.exe')
    )
    grok = @(
        (Join-Path $env:USERPROFILE '.grok\bin\grok.exe')
    )
    hermes = @(
        (Join-Path $env:LOCALAPPDATA 'hermes\bin\hermes.exe'),
        (Join-Path $env:USERPROFILE '.local\bin\hermes.exe')
    )
    aider = @(
        (Join-Path $env:USERPROFILE '.local\bin\aider.exe')
    )
}

$found = @{}
foreach ($id in $prefer) {
    $found[$id] = Find-Cli $id $extras[$id]
}
if ($ClaudePath) {
    if (-not (Test-Path -LiteralPath $ClaudePath)) {
        throw "No encuentro claude.exe ($ClaudePath)."
    }
    $found['claude'] = (Resolve-Path -LiteralPath $ClaudePath).Path
}

$defaultId = $null
if ($DefaultAgent) {
    if (-not $found[$DefaultAgent]) {
        throw "Pediste -DefaultAgent $DefaultAgent pero no esta instalado."
    }
    $defaultId = $DefaultAgent
} else {
    foreach ($id in $prefer) {
        if ($found[$id]) { $defaultId = $id; break }
    }
}

foreach ($id in $prefer) {
    if ($found[$id]) { Write-Host "    $id  $($found[$id])" }
    else { Write-Host "    $id  (no instalado)" }
}

if (-not $defaultId) {
    throw @"
No encontre ningun agente CLI. Instala al menos uno y volve a correr este script:
  Claude Code:  irm https://claude.ai/install.ps1 | iex
  Codex CLI:    npm install -g @openai/codex
  Gemini CLI:   npm install -g @google/gemini-cli
  Grok CLI:     irm https://x.ai/cli/install.ps1 | iex
  Hermes:       iex (irm https://hermes-agent.nousresearch.com/install.ps1)
  aider:        pipx install aider-chat
"@
}
Write-Host "    default: $defaultId"

# --- Helper -----------------------------------------------------------------
Step "Compilando helper en $installDir"
New-Item -ItemType Directory -Force -Path $installDir | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $installDir 'logs') | Out-Null
$csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path $csc)) { throw "No encuentro csc.exe de .NET Framework 4 ($csc)." }

& $csc /nologo /target:winexe /codepage:65001 /optimize+ `
    "/out:$installDir\Pluma.exe" `
    /r:System.Web.Extensions.dll /r:Microsoft.CSharp.dll /r:System.Core.dll `
    /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:System.Xml.dll `
    "$root\src\helper\Pluma.cs" `
    "$root\src\helper\PlumaConfig.cs"
if ($LASTEXITCODE -ne 0) { throw 'Fallo la compilacion del helper.' }

Copy-Item "$root\src\helper\word-rules.md" $installDir -Force

$config = [ordered]@{
    claudePath        = $(if ($found['claude']) { $found['claude'] } else { '' })
    workspace         = $Workspace
    model             = $Model
    claudeConfigDir   = $ClaudeConfigDir
    defaultAgent      = $defaultId
    timeoutSeconds    = 600
    charsPerSecond    = 20
    startDelayMs      = 800
    sentencePauseMs   = 450
    paragraphPauseMs  = 900
    catchUpAfterChars = 600
}
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
$freshWorkspace = -not (Test-Path $Workspace)
if ($freshWorkspace) {
    Copy-Item "$root\workspace-template" $Workspace -Recurse
    Write-Host '    Creado desde workspace-template. Edita PLUMA.md con el rol.'
} else {
    Write-Host '    Ya existe: no se pisan PLUMA.md ni el contexto.'
}

$agentsPath = Join-Path $Workspace 'agents.json'
if ($freshWorkspace -or -not (Test-Path $agentsPath)) {
    Write-DetectedAgents "$root\workspace-template\agents.json" $agentsPath $found $defaultId
    Write-Host "    agents.json con default $defaultId"
} else {
    Write-Host '    agents.json ya existe: no se toca.'
}
foreach ($extra in @('tasks.json', 'PLUMA.md')) {
    $dest = Join-Path $Workspace $extra
    $src = Join-Path "$root\workspace-template" $extra
    if (-not (Test-Path $dest) -and (Test-Path $src)) {
        Copy-Item $src $dest
        Write-Host "    Sumado $extra"
    }
}

# --- Plantilla de Word --------------------------------------------------------
Step 'Generando Pluma.dotm'

$bas = Join-Path $env:TEMP 'Pluma.bas'
$srcBas = [IO.File]::ReadAllText("$root\src\word\Pluma.bas") -replace "`r?`n", "`r`n"
[IO.File]::WriteAllText($bas, $srcBas, [Text.Encoding]::GetEncoding(1252))
$tmpDotm = Join-Path $env:TEMP 'Pluma.dotm'
Remove-Item $tmpDotm -ErrorAction SilentlyContinue

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

    $word.CustomizationContext = $doc
    $keyCode = $word.BuildKeyCode(13, [ref]1024)   # wdKeyReturn + wdKeyAlt
    $word.KeyBindings.Add(1, 'EnviarAPluma', $keyCode) | Out-Null

    $menu = $word.CommandBars.Item('Text')
    $popup = $menu.Controls.Add(10)   # msoControlPopup
    $popup.Caption = 'Pluma'
    $popup.BeginGroup = $true

    Add-PlumaButton $popup 'Enviar' 'EnviarAPluma' $false
    Add-PlumaButton $popup 'Corregir' 'TareaCorregir' $true
    Add-PlumaButton $popup ('M' + [char]0x00E1 + 's formal') 'TareaFormal' $false
    Add-PlumaButton $popup 'Resumir' 'TareaResumir' $false
    Add-PlumaButton $popup 'Expandir' 'TareaExpandir' $false
    Add-PlumaButton $popup 'Traducir' 'TareaTraducir' $false
    Add-PlumaButton $popup ('Configuraci' + [char]0x00F3 + 'n...') 'ConfigurarPluma' $true
    Add-PlumaButton $popup 'Detener' 'DetenerPluma' $true

    $saveArgs = [object[]]@([string]$tmpDotm, 15)   # 15 = wdFormatXMLTemplateMacroEnabled
    [void]$doc.GetType().InvokeMember('SaveAs2', [Reflection.BindingFlags]::InvokeMethod, $null, $doc, $saveArgs)
    $doc.Close([ref]0)
}
finally {
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
Copy-Item $tmpDotm (Join-Path $startup 'Pluma.dotm') -Force
Remove-Item $tmpDotm -ErrorAction SilentlyContinue
$oldDotm = Join-Path $startup 'ClaudeWord.dotm'
if (Test-Path $oldDotm) {
    Remove-Item $oldDotm -Force
    Write-Host '    Saque ClaudeWord.dotm viejo para que no pise Alt+Enter.'
}
Write-Host "    $startup\Pluma.dotm"

$programs = [Environment]::GetFolderPath('Programs')
if ($programs) {
    $shell = New-Object -ComObject WScript.Shell
    $sc = $shell.CreateShortcut((Join-Path $programs 'Pluma.lnk'))
    $sc.TargetPath = Join-Path $installDir 'Pluma.exe'
    $sc.Arguments = '--config'
    $sc.WorkingDirectory = $installDir
    $sc.Description = 'Configuracion de Pluma'
    $sc.Save()
    Write-Host "    Menu Inicio: $programs\Pluma.lnk"
}

Write-Host ''
Write-Host 'Listo. Abri Word, deja el cursor en un parrafo y presiona Alt+Enter' -ForegroundColor Green
Write-Host '(o clic derecho > Pluma > Enviar). Log del ultimo pedido:'
Write-Host "    $installDir\logs\last.log"
