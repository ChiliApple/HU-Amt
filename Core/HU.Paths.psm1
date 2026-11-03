#Requires -Version 5.1
<#
.SYNOPSIS
    HU.Paths - Datenordner und Sperren fuer Einzel-PC und gemeinsamen Netzordner.
.DESCRIPTION
    Drei Regeln, die immer gelten (am Einzel-PC greifen sie einfach nie):
      1. Gemeinsames liegt in Data\<Modul-Id>\, Persoenliches in Data\Benutzer\<Name>\.
      2. Jede laufende App haelt Data\Laufend\<Name>_<PC>.lock offen. Stuerzt sie ab, gibt Windows die Datei frei.
         Eine Lock-Datei, die sich loeschen laesst, ist verwaist. Das Update startet nur, wenn keine fremde mehr offen ist.
      3. Wer eine gemeinsame Datei bearbeitet, sperrt sie (<Datei>.lock mit Name). Der Naechste sieht
         "gerade in Bearbeitung von ..." und kann nur lesen.
    Die Sperren nutzen die Freigabe-Regeln von Windows (FileShare ohne Delete): Das gilt lokal und ueber SMB.
.NOTES
    Zielmaschine: der PC, auf dem die App laeuft. PowerShell 5.1 und 7.
#>

Set-StrictMode -Version 1.0

# Ordnernamen in Data\, die kein Modul belegen darf
$script:ReservedDataNames = @('Benutzer', 'Laufend', 'Docs')

function Get-HUSafeName([string]$Name) {
    # fuer Dateinamen: nur Buchstaben, Ziffern, Punkt, Minus, Unterstrich
    $n = ($Name -replace '[^\p{L}\p{Nd}\.\-_]', '_').Trim('.', ' ')
    if (-not $n) { $n = 'unbekannt' }
    return $n
}

<#
.SYNOPSIS
    Liefert die Datenpfade der App und legt die Ordner an.
#>
function Initialize-HUDataPaths {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AppRoot,
        [string]$UserName = $env:USERNAME,
        [string]$ComputerName = $env:COMPUTERNAME
    )
    $data = Join-Path $AppRoot 'Data'
    $user = Join-Path (Join-Path $data 'Benutzer') (Get-HUSafeName $UserName)
    $run = Join-Path $data 'Laufend'
    foreach ($d in $data, $user, $run) {
        if (-not (Test-Path -LiteralPath $d)) { [void](New-Item -ItemType Directory -Path $d -Force) }
    }
    return [pscustomobject]@{
        DataRoot     = $data
        UserRoot     = $user
        RunDir       = $run
        DocsDir      = Join-Path $data 'Docs'
        UserName     = "$UserName"
        ComputerName = "$ComputerName"
    }
}

<#
.SYNOPSIS
    Prueft, ob der Ordner beschreibbar ist (Datei anlegen und loeschen). Wirft bei Fehler.
#>
function Assert-HUWritable([Parameter(Mandatory)][string]$Path) {
    $probe = Join-Path $Path ('.schreibtest-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    [IO.File]::WriteAllText($probe, 'x')
    Remove-Item -LiteralPath $probe -Force
}

<#
.SYNOPSIS
    Gemeinsamer Datenordner eines Moduls: Data\<Id>\ (wird angelegt).
#>
function Get-HUModuleDataPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DataRoot,
        [Parameter(Mandatory)][string]$ModuleId
    )
    if ($ModuleId -notmatch '^[A-Za-z0-9][A-Za-z0-9_\-]{0,63}$') { throw "Ungueltige Modul-Id: '$ModuleId'" }
    if ($script:ReservedDataNames -contains $ModuleId) { throw "Modul-Id '$ModuleId' ist reserviert" }
    $p = Join-Path $DataRoot $ModuleId
    if (-not (Test-Path -LiteralPath $p)) { [void](New-Item -ItemType Directory -Path $p -Force) }
    return $p
}

# Inhalt einer Sperrdatei: Name|PC|Zeit (ISO)|PID
function ConvertTo-HULockText([string]$UserName, [string]$ComputerName) {
    return ('{0}|{1}|{2}|{3}' -f $UserName, $ComputerName, (Get-Date).ToString('s'), $PID)
}
function ConvertFrom-HULockText([string]$Text) {
    $p = @("$Text".Trim() -split '\|')
    $since = $null
    if ($p.Count -ge 3) { try { $since = [datetime]::Parse($p[2], [Globalization.CultureInfo]::InvariantCulture) } catch { } }
    return [pscustomobject]@{
        UserName     = $(if ($p.Count -ge 1) { $p[0] } else { '' })
        ComputerName = $(if ($p.Count -ge 2) { $p[1] } else { '' })
        Since        = $since
    }
}
# Sperrdatei lesen, waehrend der Besitzer sie offen haelt (er erlaubt FileShare.Read)
function Read-HULockText([string]$Path) {
    try {
        $fs = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        try { $sr = New-Object IO.StreamReader($fs, [Text.Encoding]::UTF8); return $sr.ReadToEnd() } finally { $fs.Dispose() }
    } catch { return '' }
}
# Datei neu anlegen und offen halten: andere duerfen lesen, aber weder schreiben noch loeschen
function New-HUHeldFile([string]$Path, [string]$Text) {
    $fs = [IO.File]::Open($Path, [IO.FileMode]::CreateNew, [IO.FileAccess]::ReadWrite, [IO.FileShare]::Read)
    $b = [Text.Encoding]::UTF8.GetBytes($Text)
    $fs.Write($b, 0, $b.Length); $fs.Flush()
    return $fs
}
# Verwaiste Sperrdatei entfernen. $true = weg (war verwaist oder fehlte), $false = noch in Gebrauch
function Remove-HUStaleLock([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return $true }
    try { [IO.File]::Delete($Path); return (-not (Test-Path -LiteralPath $Path)) } catch { return $false }
}

# ============================================================================
# Regel 2: laufende App
# ============================================================================

<#
.SYNOPSIS
    Meldet diese App als laufend an (Data\Laufend\<Name>_<PC>.lock bleibt offen bis Close-HURunLock).
    Rueckgabe: Objekt mit Stream und Pfad, oder $null, wenn dieselbe Datei schon offen ist (zweiter Start am selben PC).
#>
function Open-HURunLock {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RunDir,
        [string]$UserName = $env:USERNAME,
        [string]$ComputerName = $env:COMPUTERNAME
    )
    $path = Join-Path $RunDir ('{0}_{1}.lock' -f (Get-HUSafeName $UserName), (Get-HUSafeName $ComputerName))
    if (-not (Remove-HUStaleLock $path)) { return $null }
    $fs = New-HUHeldFile $path (ConvertTo-HULockText $UserName $ComputerName)
    return [pscustomobject]@{ Path = $path; Stream = $fs }
}

function Close-HURunLock($Lock) {
    if (-not $Lock) { return }
    try { if ($Lock.Stream) { $Lock.Stream.Dispose() } } catch { }
    try { if (Test-Path -LiteralPath $Lock.Path) { Remove-Item -LiteralPath $Lock.Path -Force } } catch { }
}

<#
.SYNOPSIS
    Wer hat die App gerade offen? Verwaiste Lock-Dateien werden dabei entfernt.
    -Self: eigene Sperre (Objekt aus Open-HURunLock), wird nicht mitgezaehlt.
    Rueckgabe: Liste mit UserName, ComputerName, Since, Path (leer = niemand ausser mir).
#>
function Get-HURunningInstances {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RunDir,
        $Self = $null
    )
    $out = New-Object System.Collections.Generic.List[object]
    if (-not (Test-Path -LiteralPath $RunDir)) { return @() }
    foreach ($f in @(Get-ChildItem -LiteralPath $RunDir -Filter '*.lock' -File -ErrorAction SilentlyContinue)) {
        if ($Self -and $f.FullName -eq $Self.Path) { continue }
        $txt = Read-HULockText $f.FullName
        if (Remove-HUStaleLock $f.FullName) { continue }
        $i = ConvertFrom-HULockText $txt
        $out.Add([pscustomobject]@{ UserName = $i.UserName; ComputerName = $i.ComputerName; Since = $i.Since; Path = $f.FullName })
    }
    return @($out.ToArray())
}

# ============================================================================
# Regel 3: gemeinsame Datei bearbeiten
# ============================================================================

<#
.SYNOPSIS
    Sperrt eine gemeinsame Datei zum Bearbeiten (<Datei>.lock).
    Rueckgabe: Ok = $true mit Stream (bis Unlock-HUDataFile), oder Ok = $false mit UserName, ComputerName, Since des anderen.
#>
function Lock-HUDataFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$UserName = $env:USERNAME,
        [string]$ComputerName = $env:COMPUTERNAME
    )
    $lp = "$Path.lock"
    for ($try = 0; $try -lt 2; $try++) {
        if (-not (Remove-HUStaleLock $lp)) { break }
        try {
            $fs = New-HUHeldFile $lp (ConvertTo-HULockText $UserName $ComputerName)
            return [pscustomobject]@{ Ok = $true; Path = $Path; LockPath = $lp; Stream = $fs; UserName = $UserName; ComputerName = $ComputerName; Since = Get-Date }
        } catch [System.IO.IOException] {
            # ein anderer war gleichzeitig schneller: noch einmal pruefen
        }
    }
    $i = ConvertFrom-HULockText (Read-HULockText $lp)
    return [pscustomobject]@{ Ok = $false; Path = $Path; LockPath = $lp; Stream = $null; UserName = $i.UserName; ComputerName = $i.ComputerName; Since = $i.Since }
}

function Unlock-HUDataFile($Lock) {
    if (-not $Lock -or -not $Lock.Ok) { return }
    try { if ($Lock.Stream) { $Lock.Stream.Dispose() } } catch { }
    try { if (Test-Path -LiteralPath $Lock.LockPath) { Remove-Item -LiteralPath $Lock.LockPath -Force } } catch { }
}

# Text fuer die Oberflaeche: "gerade in Bearbeitung von Maria (PC-12) seit 10:15"
function Get-HULockInfoText($Lock) {
    if (-not $Lock -or $Lock.Ok) { return '' }
    $who = $(if ($Lock.UserName) { $Lock.UserName } else { 'jemand anderem' })
    $pc = $(if ($Lock.ComputerName) { " ($($Lock.ComputerName))" } else { '' })
    $since = $(if ($Lock.Since) { " seit $($Lock.Since.ToString('dd.MM. HH:mm'))" } else { '' })
    return "Gerade in Bearbeitung von $who$pc$since. Du kannst die Daten nur ansehen."
}

<#
.SYNOPSIS
    Kopie einer Datei mit Zeitstempel vor dem Speichern: <Ordner>\_Kopien\<Name>_JJJJMMTT-HHMMSS<Ext>.
    Behaelt die letzten -Keep Kopien dieser Datei. Rueckgabe: Pfad der Kopie oder '' (Datei fehlte).
#>
function Backup-HUDataFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$Keep = 20
    )
    if (-not (Test-Path -LiteralPath $Path)) { return '' }
    $dir = Join-Path (Split-Path $Path -Parent) '_Kopien'
    if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
    $base = [IO.Path]::GetFileNameWithoutExtension($Path); $ext = [IO.Path]::GetExtension($Path)
    $dest = Join-Path $dir ('{0}_{1}{2}' -f $base, (Get-Date).ToString('yyyyMMdd-HHmmss'), $ext)
    $n = 2
    while (Test-Path -LiteralPath $dest) { $dest = Join-Path $dir ('{0}_{1}_{2}{3}' -f $base, (Get-Date).ToString('yyyyMMdd-HHmmss'), $n, $ext); $n++ }
    Copy-Item -LiteralPath $Path -Destination $dest
    if ($Keep -gt 0) {
        $old = @(Get-ChildItem -LiteralPath $dir -File | Where-Object { $_.Name -like "$($base)_*$ext" } | Sort-Object Name -Descending | Select-Object -Skip $Keep)
        foreach ($o in $old) { Remove-Item -LiteralPath $o.FullName -Force -ErrorAction SilentlyContinue }
    }
    return $dest
}

Export-ModuleMember -Function Initialize-HUDataPaths, Assert-HUWritable, Get-HUModuleDataPath, Get-HUSafeName,
    Open-HURunLock, Close-HURunLock, Get-HURunningInstances,
    Lock-HUDataFile, Unlock-HUDataFile, Get-HULockInfoText, Backup-HUDataFile
