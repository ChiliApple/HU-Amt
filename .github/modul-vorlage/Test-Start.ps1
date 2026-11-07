#Requires -Version 5.1
<#
.SYNOPSIS
    Starttest der App mit einem Modul in einer frischen Kopie der Test-App (PowerShell 5.1 und 7).
.DESCRIPTION
    Zielmaschine: Entwicklungs-PC. Vorlage aus dem Kern (.github\modul-vorlage), Module kopieren sie unverändert nach .github\dev.
    - Kopiert die Test-App ohne Data, Addons, AddonsDev und Logs in einen Temp-Ordner, legt nur dieses Modul nach
      AddonsDev\<Modul-Id> und ruft Main.ps1 -SmokeTest auf. So stört kein anderes Modul (z. B. ein zweites, das
      dieselbe Dateiart annimmt), und die Daten der Test-App bleiben unberührt. Die Test-App selbst wird nur gelesen.
    - Erwartet im Ergebnis "OK", "Module=1" und jedes Muster aus -Erwartet.
    - Einstellungen je Modul ohne Parameter: .github\dev\modul-dev.json neben diesem Skript (siehe LIESMICH.md).
      Angegebene Parameter gehen vor.
.PARAMETER ModulId
    Ordnername unter AddonsDev. Leer = "id" aus addon.json der Quelle.
.PARAMETER Quelle
    Modulordner. Leer = Hauptordner des Repos (zwei Ebenen über diesem Skript).
.PARAMETER App
    Test-App, aus der kopiert wird.
.PARAMETER Erwartet
    Zusätzliche Muster (Regex), die im Ergebnis des Starttests stehen müssen, z. B. ein Schritt des Moduls.
.PARAMETER MinAppAnpassen
    Nur Entwicklung: Ist die Test-App älter als minApp des Moduls, bekommt die Kopie die Version der Test-App als minApp.
#>
[CmdletBinding()]
param(
    [string]$ModulId = '',
    [string]$Quelle = '',
    [string]$App = 'C:\Tools\HU-Amt',
    [ValidateSet('beide', '5', '7')][string]$PowerShell = 'beide',
    [string[]]$Erwartet = @(),
    [switch]$MinAppAnpassen,
    [ValidateRange(10, 3600)][int]$TimeoutSec = 180
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 1.0

function Get-DevWert($Cfg, [string]$Name, $Standard) {
    if ($Cfg -and $Cfg.PSObject.Properties[$Name]) { return $Cfg.PSObject.Properties[$Name].Value }
    return $Standard
}
function Read-JsonDatei([string]$Pfad) {
    return ([IO.File]::ReadAllText($Pfad, [Text.Encoding]::UTF8) | ConvertFrom-Json)
}
# Nur Entwicklung: minApp der Kopie auf die Version der Test-App setzen, wenn die Test-App älter ist (nie die Quelle)
function Set-DevMinApp([string]$AddonDir, [string]$AppDir) {
    $vf = Join-Path $AppDir 'Config\version.json'
    $af = Join-Path $AddonDir 'addon.json'
    if (-not (Test-Path -LiteralPath $vf) -or -not (Test-Path -LiteralPath $af)) { return }
    $appV = [version]"$((Read-JsonDatei $vf).version)"
    $t = [IO.File]::ReadAllText($af, [Text.Encoding]::UTF8)
    $min = [version]"$(($t | ConvertFrom-Json).minApp)"
    if ($appV -ge $min) { return }
    $neu = $t -replace '"minApp"\s*:\s*"[^"]*"', ('"minApp": "' + $appV.ToString() + '"')
    [IO.File]::WriteAllText($af, $neu, (New-Object System.Text.UTF8Encoding $true))
    Write-Warning "Test-App ist $appV, das Modul verlangt $($min): Kopie mit minApp $appV (nur Entwicklung)."
}

if (-not $Quelle) { $Quelle = Split-Path -Parent (Split-Path -Parent $PSScriptRoot) }
$Quelle = (Resolve-Path -LiteralPath $Quelle).ProviderPath
$cfgPfad = Join-Path $PSScriptRoot 'modul-dev.json'
$cfg = $null
if (Test-Path -LiteralPath $cfgPfad) { $cfg = Read-JsonDatei $cfgPfad }
if (-not $PSBoundParameters.ContainsKey('Erwartet')) { $Erwartet = @(Get-DevWert $cfg 'erwartet' @()) }
if (-not $PSBoundParameters.ContainsKey('MinAppAnpassen')) { $MinAppAnpassen = [bool](Get-DevWert $cfg 'minAppAnpassen' $false) }
if (-not $ModulId) {
    $af = Join-Path $Quelle 'addon.json'
    if (-not (Test-Path -LiteralPath $af)) { throw "Keine addon.json in $Quelle. Bitte -ModulId und -Quelle angeben." }
    $ModulId = "$((Read-JsonDatei $af).id)"
}
if ($ModulId -notmatch '^[A-Za-z0-9][A-Za-z0-9_\-]*$') { throw "Ungültige Modul-Id '$ModulId'." }
if (-not (Test-Path -LiteralPath (Join-Path $App 'Main.ps1'))) { throw "Keine Test-App in $App (Main.ps1 fehlt)." }

$work = Join-Path ([IO.Path]::GetTempPath()) ("HUAmtStart-$ModulId-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $work -Force | Out-Null
try {
    foreach ($e in @(Get-ChildItem -LiteralPath $App -Force | Where-Object { $_.Name -notin 'Data', 'Addons', 'AddonsDev', 'Logs' })) {
        Copy-Item -LiteralPath $e.FullName -Destination $work -Recurse -Force
    }
    $dst = Join-Path $work "AddonsDev\$ModulId"
    New-Item -ItemType Directory -Path $dst -Force | Out-Null
    foreach ($e in @(Get-ChildItem -LiteralPath $Quelle -Force | Where-Object { -not $_.Name.StartsWith('.') })) {
        Copy-Item -LiteralPath $e.FullName -Destination $dst -Recurse -Force
    }
    if ($MinAppAnpassen) { Set-DevMinApp -AddonDir $dst -AppDir $work }
    $runs = @()
    if ($PowerShell -in 'beide', '5') { $runs += [pscustomobject]@{ Name = 'PS 5.1'; Exe = (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe') } }
    if ($PowerShell -in 'beide', '7') {
        $pw = Get-Command pwsh.exe -ErrorAction SilentlyContinue
        if ($pw) { $runs += [pscustomobject]@{ Name = 'PS 7'; Exe = $pw.Source } } else { Write-Warning 'pwsh.exe nicht gefunden, PS 7 übersprungen' }
    }
    $fail = 0
    foreach ($r in $runs) {
        $out = Join-Path $work ("smoke-{0}.txt" -f ($r.Name -replace '\W', ''))
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $p = Start-Process -FilePath $r.Exe -ArgumentList @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', "`"$(Join-Path $work 'Main.ps1')`"", '-SmokeTest', "`"$out`"") -PassThru -WindowStyle Hidden
        if (-not $p.WaitForExit($TimeoutSec * 1000)) { try { $p.Kill() } catch { }; Write-Host "$($r.Name): FEHLER Zeitüberschreitung"; $fail++; continue }
        $res = if (Test-Path -LiteralPath $out) { (Get-Content -LiteralPath $out -Raw -Encoding UTF8).Trim() } else { 'FEHLER keine Ergebnisdatei' }
        $ok = $res -like 'OK *' -and $res -match 'Module=1'
        foreach ($m in $Erwartet) { if ("$m" -and $res -notmatch "$m") { $ok = $false } }
        if (-not $ok) { $fail++ }
        Write-Host ("{0}: {1} ({2:N0} s)`n  {3}" -f $r.Name, $(if ($ok) { 'OK' } else { 'FEHLER' }), $sw.Elapsed.TotalSeconds, $res)
        if (-not $ok) {
            $log = Get-ChildItem -LiteralPath (Join-Path $work 'Data') -Recurse -Filter '*.log' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime | Select-Object -Last 1
            if ($log) { Get-Content -LiteralPath $log.FullName -Tail 15 | ForEach-Object { Write-Host "  | $_" } }
        }
    }
    if (-not $runs.Count) { throw 'Keine PowerShell für den Starttest gefunden.' }
    if ($fail) { throw "$fail Starttest(s) fehlgeschlagen" }
} finally {
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}
