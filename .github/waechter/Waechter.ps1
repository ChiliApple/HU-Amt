#Requires -Version 5.1
<#
.SYNOPSIS
    Wächter für externe Werkzeuge und Bibliotheken eines Moduls: neuere Versionen, bekannte Sicherheitslücken, Ende der Pflege.
.DESCRIPTION
    Zielmaschine: GitHub-Runner (wöchentlich über den Workflow des Kerns .github\workflows\waechter.yml) oder Entwicklungs-PC zum Ausprobieren.
    Läuft unter Windows PowerShell 5.1 und PowerShell 7. Liegt nur im Kern, gilt für alle Module.
    Liest <Modul>\.github\waechter.json, dazu werkzeuge.json und requirements.txt, wenn dort angegeben. Ändert nichts am Modul.
    Schreibt in -Ausgabe:
      waechter-ergebnis.json   maschinenlesbar (für die Übersicht in der Modulverwaltung)
      waechter-bericht.md      Bericht für das Issue
    In der CI zusätzlich: GITHUB_OUTPUT (befunde, signatur) und die Zusammenfassung des Laufs.
    Optional: NVD_API_KEY (schnellere Abfragen bei der NVD), GITHUB_TOKEN (mehr Abfragen bei GitHub).
    -Pruefen: nur waechter.json des Moduls prüfen (ohne Netz), Exit 1 bei Problemen.
.EXAMPLE
    powershell.exe -File .github\waechter\Waechter.ps1 -Wurzel ..\HU-Amt-OCR -Ausgabe $env:TEMP
.EXAMPLE
    powershell.exe -File .github\waechter\Waechter.ps1 -Wurzel ..\HU-Amt-OCR -Pruefen
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Wurzel,
    [string]$Ausgabe = '',
    [int]$NvdPause = -1,
    [switch]$Pruefen
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$Wurzel = (Resolve-Path -LiteralPath $Wurzel).Path
if (-not $Ausgabe) { $Ausgabe = (Get-Location).Path }
if ($NvdPause -lt 0) { $NvdPause = $(if ($env:NVD_API_KEY) { 1 } else { 7 }) }
. (Join-Path $PSScriptRoot 'Waechter.Funktionen.ps1')

$probleme = @(Test-HUWKonfig $Wurzel)
if ($probleme.Count) {
    $probleme | ForEach-Object { Write-Host "waechter.json: $_" -ForegroundColor Red }
    exit 1
}
if ($Pruefen) { Write-Host 'waechter.json in Ordnung'; exit 0 }

$konfig = Get-Content -LiteralPath (Join-Path $Wurzel '.github/waechter.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$erg = Invoke-HUWaechter -Wurzel $Wurzel -Konfig $konfig -NvdPause $NvdPause
$bericht = ConvertTo-HUWBericht $erg

New-Item -ItemType Directory -Path $Ausgabe -Force | Out-Null
$utf8 = New-Object System.Text.UTF8Encoding $false
[IO.File]::WriteAllText((Join-Path $Ausgabe 'waechter-ergebnis.json'), ($erg | ConvertTo-Json -Depth 8), $utf8)
[IO.File]::WriteAllText((Join-Path $Ausgabe 'waechter-bericht.md'), $bericht, $utf8)

foreach ($e in @($erg.eintraege)) {
    $l = @($e.luecken).Count
    $zus = @(); if ($l) { $zus += "$l Lücke(n)" }; $zus += @($e.hinweise); if ($e.fehler) { $zus += "Fehler: $($e.fehler)" }
    Write-Host ('{0,-14} {1,-16} {2,-16} {3,-8} {4}' -f $e.id, $e.version, $e.neueste, $e.status, ($zus -join '; '))
}
Write-Host "Befunde: $($erg.befunde), Signatur $($erg.signatur)"
if ($env:GITHUB_OUTPUT) { "befunde=$($erg.befunde)", "signatur=$($erg.signatur)" | Add-Content -Path $env:GITHUB_OUTPUT -Encoding UTF8 }
if ($env:GITHUB_STEP_SUMMARY) { $bericht | Add-Content -Path $env:GITHUB_STEP_SUMMARY -Encoding UTF8 }
