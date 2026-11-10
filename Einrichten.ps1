#Requires -Version 5.1
<#
.SYNOPSIS
    HU-Amt aus einem Paket fuer eine Aussenstelle einrichten (ohne Administratorrechte).
.DESCRIPTION
    Wird aus dem ausgepackten Paket ueber Einrichten.cmd gestartet (dieses Skript liegt im Ordner Programm des Pakets).
    - Ziel: %LOCALAPPDATA%\HU-Amt-<Name> des angemeldeten Benutzers. Dort darf nur dieser Benutzer schreiben, andere
      Benutzer des PCs koennen weder Programm noch Verwalter-Liste aendern. Kein OneDrive, keine Administratorrechte noetig.
    - Erstes Einrichten: alles kopieren. Schon eingerichtet: Programm und Module erneuern, Data\ bleibt unveraendert.
    - Entfernt die Internet-Markierung (Zone.Identifier) aller Dateien, legt die Desktop-Verknuepfung "HU-Amt <Name>" an
      und startet HU-Amt. Beim ersten Start prueft HU-Amt das Programm und liest den beigelegten Zugang ein.
    Jeder Windows-Benutzer, der HU-Amt braucht, richtet es fuer sich ein (eigene Daten).
.NOTES
    Zielmaschine: der PC der Aussenstelle. Windows PowerShell 5.1 (startet auch unter PowerShell 7).
#>
param([switch]$NoStart, [switch]$Quiet, [string]$TargetRoot = '')
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 1.0
try { Add-Type -AssemblyName PresentationFramework -ErrorAction Stop } catch { }

$title = 'HU-Amt einrichten'
function Show-EinMsg([string]$Text, [string]$Icon = 'Information', [string]$Buttons = 'OK') {
    if ($Quiet) { Write-Host $Text; return 'Yes' }
    return "$([System.Windows.MessageBox]::Show($Text, $title, $Buttons, $Icon))"
}

$src = $PSScriptRoot
$fail = $null
try {
    # 1. Quelle pruefen: ausgepacktes Paket, nicht aus dem ZIP heraus gestartet
    foreach ($f in 'Main.ps1', 'Start-HU-Amt.cmd', 'Config\app.json', 'Config\paket.json', 'Config\installation.json') {
        if (-not (Test-Path -LiteralPath (Join-Path $src $f))) {
            throw "Das Paket ist nicht vollständig ausgepackt.`n`nBitte die ZIP-Datei mit der rechten Maustaste anklicken, 'Alle extrahieren ...' wählen und danach im ausgepackten Ordner Einrichten doppelklicken."
        }
    }
    $inst = [IO.File]::ReadAllText((Join-Path $src 'Config\installation.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
    $name = "$($inst.name)".Trim()
    if (-not $name) { throw 'Im Paket fehlt der Name der Außenstelle.' }
    $folder = 'HU-Amt-' + (($name -replace '[^\p{L}\p{Nd}]+', '-').Trim('-'))
    $root = $(if ($TargetRoot) { $TargetRoot } else { $env:LOCALAPPDATA })
    if (-not $root) { throw 'Der persönliche Programmordner (LOCALAPPDATA) ist nicht bekannt.' }
    $target = Join-Path $root $folder
    if ([IO.Path]::GetFullPath($target).TrimEnd('\') -ieq [IO.Path]::GetFullPath($src).TrimEnd('\')) { throw 'HU-Amt ist hier schon eingerichtet. Bitte über das Symbol am Desktop starten.' }

    # 2. laeuft HU-Amt aus dem Ziel noch?
    $mainT = [regex]::Escape((Join-Path $target 'Main.ps1'))
    $running = @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe' or Name='pwsh.exe'" -ErrorAction SilentlyContinue | Where-Object { "$($_.CommandLine)" -match $mainT })
    if ($running.Count) { throw "HU-Amt $name ist gerade geöffnet.`n`nBitte HU-Amt schließen und Einrichten noch einmal doppelklicken." }

    # 3. kopieren (schon eingerichtet: Data\ bleibt)
    $update = Test-Path -LiteralPath (Join-Path $target 'Main.ps1')
    if ($update -and (Show-EinMsg "HU-Amt $name ist auf diesem PC schon eingerichtet.`n`nProgramm und Module erneuern? Module, die nicht im neuen Paket sind, werden entfernt. Die eingetragenen Daten bleiben erhalten." 'Question' 'YesNo') -ne 'Yes') { return }
    if (-not (Test-Path -LiteralPath $target)) { [void](New-Item -ItemType Directory -Path $target -Force) }
    $srcFull = [IO.Path]::GetFullPath($src).TrimEnd('\')
    foreach ($f in @(Get-ChildItem -LiteralPath $src -Recurse -File -Force)) {
        $rel = $f.FullName.Substring($srcFull.Length + 1)
        if ($update -and $rel -like 'Data\*') { continue }
        $dst = Join-Path $target $rel
        $dd = Split-Path $dst -Parent
        if (-not (Test-Path -LiteralPath $dd)) { [void](New-Item -ItemType Directory -Path $dd -Force) }
        Copy-Item -LiteralPath $f.FullName -Destination $dst -Force
    }
    if ($update) {
        # Module, die das neue Paket nicht mehr enthaelt, entfernen (ihre Daten in Data\ bleiben).
        # Den Zugang zu ihnen loescht HU-Amt beim ersten Start, sonst laedt das Modul-Update sie wieder.
        $srcAdd = Join-Path $src 'Addons'; $tgtAdd = Join-Path $target 'Addons'
        $keep = @(); if (Test-Path -LiteralPath $srcAdd) { $keep = @(Get-ChildItem -LiteralPath $srcAdd -Directory -Force | ForEach-Object { $_.Name }) }
        if (Test-Path -LiteralPath $tgtAdd) {
            foreach ($d in @(Get-ChildItem -LiteralPath $tgtAdd -Directory -Force | Where-Object { $keep -notcontains $_.Name })) {
                Remove-Item -LiteralPath $d.FullName -Recurse -Force
            }
        }
        # Paket neu: noch einmal pruefen lassen und den neuen Zugang einlesen
        $pj = Join-Path $target 'Config\paket.json'
        try { $p = [IO.File]::ReadAllText($pj, [Text.Encoding]::UTF8) | ConvertFrom-Json; $p.geprueft = $false; [IO.File]::WriteAllText($pj, (ConvertTo-Json -InputObject $p -Depth 6), (New-Object Text.UTF8Encoding $true)) } catch { }
    }
    Get-ChildItem -LiteralPath $target -Recurse -File -Force | ForEach-Object { try { Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue } catch { } }

    # 4. Desktop-Verknuepfung "HU-Amt <Name>" (gleicher Name wie in der App)
    $lnkName = ((("HU-Amt $name") -replace '[\\/:*?"<>|]', ' ') -replace '\s+', ' ').Trim()
    $desk = [Environment]::GetFolderPath('Desktop')
    if ($desk -and -not $Quiet) {
        $sh = New-Object -ComObject WScript.Shell
        try {
            $s = $sh.CreateShortcut((Join-Path $desk "$lnkName.lnk"))
            $s.TargetPath = Join-Path $target 'Start-HU-Amt.cmd'
            $s.WorkingDirectory = $target
            $s.WindowStyle = 7
            $s.IconLocation = "$(Join-Path $target 'Assets\icon.ico'),0"
            $s.Description = "HU-Amt $name"
            $s.Save()
        } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sh) }
    }

    # 5. starten
    if (-not $NoStart) { Start-Process -FilePath (Join-Path $target 'Start-HU-Amt.cmd') -WorkingDirectory $target -WindowStyle Hidden }
    [void](Show-EinMsg ("HU-Amt $name ist eingerichtet und startet jetzt.`n`n" +
            "Ab jetzt startest du es mit dem Symbol '$lnkName' am Desktop.`n`n" +
            "Den ausgepackten Ordner und die ZIP-Datei kannst du jetzt löschen."))
} catch { $fail = $_.Exception.Message }
if ($fail) { [void](Show-EinMsg "HU-Amt wurde nicht eingerichtet:`n`n$fail" 'Warning'); exit 1 }
exit 0
