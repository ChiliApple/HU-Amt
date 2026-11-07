#Requires -Version 5.1
<#
.SYNOPSIS
    Kopiert ein Modul zum Entwickeln in die Test-App (AddonsDev, ohne Signatur).
.DESCRIPTION
    Zielmaschine: nb001. Vorlage aus dem Kern (.github\modul-vorlage), Module kopieren sie unverändert nach .github\dev.
    - Eigene Textdateien des Moduls werden vorher auf UTF-8 mit BOM und CRLF gebracht (.yml und .py bleiben unberührt,
      dafür gibt es Format-Dateien.ps1).
    - Quelle: Hauptordner des Repos (das Modul), ohne Einträge, die mit "." beginnen (.git, .github, .gitignore ...).
    - Ziel: <App>\AddonsDev\<Modul-Id> (wird ersetzt).
    - Bricht ab, wenn die Test-App läuft (Prozess mit <App>\Main.ps1 in der Befehlszeile, ohne Starttests) oder eine
      Datei im Ziel gesperrt ist (z. B. DLLs, die eine laufende App geladen hat). Dann bleibt das Ziel unverändert.
    - Bricht ab, wenn im Modul Dateien liegen, die nach Data gehören (Daten, Logs, Belege) oder nicht ins Repo (Programme).
    - Einstellungen je Modul ohne Parameter: .github\dev\modul-dev.json neben diesem Skript (siehe LIESMICH.md).
      Angegebene Parameter gehen vor.
.PARAMETER ModulId
    Ordnername unter AddonsDev. Leer = "id" aus addon.json der Quelle.
.PARAMETER Quelle
    Modulordner. Leer = Hauptordner des Repos (zwei Ebenen über diesem Skript).
.PARAMETER App
    Test-App. Muss Main.ps1 enthalten.
.PARAMETER VerboteneEndungen
    Dateiendungen, die nicht im Modul liegen dürfen.
.PARAMETER VerboteneNamen
    Datei- oder Ordnernamen, die nicht im Modul liegen dürfen (z. B. Datendateien des Moduls).
.PARAMETER ErlaubteNamen
    Dateinamen, die trotz verbotener Endung erlaubt sind (z. B. eine Testseite, die zum Modul gehört).
.PARAMETER FormatAusschluss
    Regex auf den relativen Pfad: diese Textdateien werden nicht umformatiert (z. B. fremde Bibliotheken).
.PARAMETER OhneFormat
    Keine Textdatei umformatieren.
.PARAMETER MinAppAnpassen
    Nur Entwicklung: Ist die Test-App älter als minApp des Moduls, bekommt die Kopie die Version der Test-App als minApp.
#>
[CmdletBinding()]
param(
    [string]$ModulId = '',
    [string]$Quelle = '',
    [string]$App = 'C:\Tools\HU-Amt',
    [string[]]$VerboteneEndungen = @('.log', '.exe', '.dll', '.zip', '.pdf', '.png', '.jpg', '.jpeg', '.csv', '.pyc'),
    [string[]]$VerboteneNamen = @('einstellungen.json', '__pycache__'),
    [string[]]$ErlaubteNamen = @(),
    [string]$FormatAusschluss = '^\.github\\dev\\Werkzeuge\\',
    [switch]$OhneFormat,
    [switch]$MinAppAnpassen
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

# --- Einstellungen ---
if (-not $Quelle) { $Quelle = Split-Path -Parent (Split-Path -Parent $PSScriptRoot) }
$Quelle = (Resolve-Path -LiteralPath $Quelle).ProviderPath.TrimEnd('\')
$cfgPfad = Join-Path $PSScriptRoot 'modul-dev.json'
$cfg = $null
if (Test-Path -LiteralPath $cfgPfad) { $cfg = Read-JsonDatei $cfgPfad }
if (-not $PSBoundParameters.ContainsKey('VerboteneEndungen')) { $VerboteneEndungen = @(Get-DevWert $cfg 'verboteneEndungen' $VerboteneEndungen) }
if (-not $PSBoundParameters.ContainsKey('VerboteneNamen')) { $VerboteneNamen = @(Get-DevWert $cfg 'verboteneNamen' $VerboteneNamen) }
if (-not $PSBoundParameters.ContainsKey('ErlaubteNamen')) { $ErlaubteNamen = @(Get-DevWert $cfg 'erlaubteNamen' $ErlaubteNamen) }
if (-not $PSBoundParameters.ContainsKey('FormatAusschluss')) { $FormatAusschluss = "$(Get-DevWert $cfg 'formatAusschluss' $FormatAusschluss)" }
if (-not $PSBoundParameters.ContainsKey('OhneFormat')) { $OhneFormat = [bool](Get-DevWert $cfg 'ohneFormat' $false) }
if (-not $PSBoundParameters.ContainsKey('MinAppAnpassen')) { $MinAppAnpassen = [bool](Get-DevWert $cfg 'minAppAnpassen' $false) }
if (-not $ModulId) {
    $af = Join-Path $Quelle 'addon.json'
    if (-not (Test-Path -LiteralPath $af)) { throw "Keine addon.json in $Quelle. Bitte -ModulId und -Quelle angeben." }
    $ModulId = "$((Read-JsonDatei $af).id)"
}
if ($ModulId -notmatch '^[A-Za-z0-9][A-Za-z0-9_\-]*$') { throw "Ungültige Modul-Id '$ModulId'." }
if (-not (Test-Path -LiteralPath (Join-Path $App 'Main.ps1'))) { throw "Keine Test-App in $App (Main.ps1 fehlt)." }
$App = (Resolve-Path -LiteralPath $App).ProviderPath.TrimEnd('\')
$verbEnd = @($VerboteneEndungen | ForEach-Object { "$_".ToLowerInvariant() })

# --- Textdateien auf UTF-8 mit BOM und CRLF ---
if (-not $OhneFormat) {
    $bom = New-Object System.Text.UTF8Encoding $true
    # Achtung: -Include wirkt in PS 5.1 nicht zusammen mit -LiteralPath, deshalb nach Endung filtern
    $textExt = '.ps1', '.psm1', '.xaml', '.json', '.md', '.html', '.txt'
    foreach ($f in @(Get-ChildItem -LiteralPath $Quelle -Recurse -File -Force)) {
        $rel = $f.FullName.Substring($Quelle.Length + 1)
        if ($textExt -notcontains $f.Extension.ToLowerInvariant() -or $rel -match '^\.git\\') { continue }
        if ($FormatAusschluss -and $rel -match $FormatAusschluss) { continue }
        $t = [IO.File]::ReadAllText($f.FullName, [Text.Encoding]::UTF8)
        $n = ($t -replace "`r`n", "`n") -replace "`n", "`r`n"
        $bytes = [IO.File]::ReadAllBytes($f.FullName)
        $hasBom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
        if ($n -ne $t -or -not $hasBom) { [IO.File]::WriteAllText($f.FullName, $n, $bom); Write-Host "BOM/CRLF: $rel" }
    }
}

# --- Dateien, die nicht ins Modul gehören ---
$items = @(Get-ChildItem -LiteralPath $Quelle -Force | Where-Object { -not $_.Name.StartsWith('.') })
$bad = New-Object System.Collections.Generic.List[string]
foreach ($it in $items) {
    $all = @($it)
    if ($it.PSIsContainer) { $all += @(Get-ChildItem -LiteralPath $it.FullName -Recurse -Force) }
    foreach ($x in $all) {
        $rel = $x.FullName.Substring($Quelle.Length + 1)
        if ($VerboteneNamen -contains $x.Name) { $bad.Add($rel); continue }
        if (-not $x.PSIsContainer -and $verbEnd -contains $x.Extension.ToLowerInvariant() -and $ErlaubteNamen -notcontains $x.Name) { $bad.Add($rel) }
    }
}
if ($bad.Count) { throw "Im Modul liegen Dateien, die nicht hineingehören: $($bad -join ', ')" }

# --- läuft die Test-App? ---
$main = Join-Path $App 'Main.ps1'
$run = @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='pwsh.exe'" -ErrorAction SilentlyContinue |
        Where-Object { "$($_.CommandLine)".IndexOf($main, [StringComparison]::OrdinalIgnoreCase) -ge 0 -and "$($_.CommandLine)" -notmatch '-SmokeTest' })
if ($run.Count) { throw "Die Test-App läuft ($App, Prozess $(($run | ForEach-Object ProcessId) -join ', ')). Bitte zuerst schließen." }
$dst = Join-Path $App "AddonsDev\$ModulId"
if (Test-Path -LiteralPath $dst) {
    # gesperrte Dateien (z. B. geladene DLLs): nichts löschen, sonst bleibt ein halber Modulordner
    foreach ($f in @(Get-ChildItem -LiteralPath $dst -Recurse -File -Force)) {
        try { $fs = [IO.File]::Open($f.FullName, 'Open', 'ReadWrite', 'None'); $fs.Dispose() }
        catch { throw "Die Test-App läuft noch ($($f.Name) ist gesperrt). Bitte die Test-App schließen." }
    }
    Remove-Item -LiteralPath $dst -Recurse -Force
}

# --- kopieren ---
New-Item -ItemType Directory -Path $dst -Force | Out-Null
foreach ($e in $items) { Copy-Item -LiteralPath $e.FullName -Destination $dst -Recurse -Force }
if ($MinAppAnpassen) { Set-DevMinApp -AddonDir $dst -AppDir $App }
Write-Host "Kopiert nach $dst"
