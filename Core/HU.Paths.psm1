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
$script:ReservedDataNames = @('Benutzer', 'Laufend', 'Docs', 'Postfach', 'Wiederherstellung')
# Breite Gruppen (ueber SIDs, nicht Namen: deutsche und englische Windows heissen verschieden)
$script:HUAclBroad = [ordered]@{ 'S-1-1-0' = 'Jeder'; 'S-1-5-11' = 'Authentifizierte Benutzer'; 'S-1-5-32-545' = 'Benutzer' }
# Schreibende Rechte an der Datei: WriteData 0x2, AppendData 0x4, Delete 0x10000, ChangePermissions 0x40000,
# TakeOwnership 0x80000, dazu GENERIC_ALL 0x10000000 und GENERIC_WRITE 0x40000000 (roh in manchen Eintraegen)
$script:HUAclFileWrite = 0x2 -bor 0x4 -bor 0x10000 -bor 0x40000 -bor 0x80000 -bor 0x10000000 -bor 0x40000000
# Am Ordner: DeleteSubdirectoriesAndFiles 0x40 (loescht jede Datei darin, egal welche Rechte die Datei hat),
# ChangePermissions, TakeOwnership, GENERIC_ALL. "Aendern" (Modify) enthaelt 0x40 nicht, "Vollzugriff" schon.
$script:HUAclFolderWrite = 0x40 -bor 0x40000 -bor 0x80000 -bor 0x10000000

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
        [string]$PcName = $env:COMPUTERNAME,
        [switch]$TestMode   # Datenstamm Data\_Test\ (Testbestand); laufende Apps bleiben in Data\Laufend
    )
    $base = Join-Path $AppRoot 'Data'
    $data = $(if ($TestMode) { Join-Path $base '_Test' } else { $base })
    $user = Join-Path (Join-Path $data 'Benutzer') (Get-HUSafeName $UserName)
    $run = Join-Path $base 'Laufend'
    foreach ($d in $data, $user, $run) {
        if (-not (Test-Path -LiteralPath $d)) { [void](New-Item -ItemType Directory -Path $d -Force) }
    }
    return [pscustomobject]@{
        DataRoot     = $data
        UserRoot     = $user
        RunDir       = $run
        DocsDir      = Join-Path $data 'Docs'
        UserName     = "$UserName"
        ComputerName = "$PcName"
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
function ConvertTo-HULockText([string]$UserName, [string]$PcName) {
    return ('{0}|{1}|{2}|{3}' -f $UserName, $PcName, (Get-Date).ToString('s'), $PID)
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
        [string]$PcName = $env:COMPUTERNAME
    )
    $path = Join-Path $RunDir ('{0}_{1}.lock' -f (Get-HUSafeName $UserName), (Get-HUSafeName $PcName))
    if (-not (Remove-HUStaleLock $path)) { return $null }
    $fs = New-HUHeldFile $path (ConvertTo-HULockText $UserName $PcName)
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
        [string]$PcName = $env:COMPUTERNAME
    )
    $lp = "$Path.lock"
    for ($try = 0; $try -lt 2; $try++) {
        if (-not (Remove-HUStaleLock $lp)) { break }
        try {
            $fs = New-HUHeldFile $lp (ConvertTo-HULockText $UserName $PcName)
            return [pscustomobject]@{ Ok = $true; Path = $Path; LockPath = $lp; Stream = $fs; UserName = $UserName; ComputerName = $PcName; Since = Get-Date }
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
        # nur eigene Kopien dieser Datei (Name_JJJJMMTT-HHMMSS[_n]Ext), keine fremden Dateien mit gleichem Anfang
        $rx = '^' + [regex]::Escape($base) + '_\d{8}-\d{6}(_\d+)?' + [regex]::Escape($ext) + '$'
        $old = @(Get-ChildItem -LiteralPath $dir -File | Where-Object { $_.Name -match $rx } | Sort-Object Name -Descending | Select-Object -Skip $Keep)
        foreach ($o in $old) { Remove-Item -LiteralPath $o.FullName -Force -ErrorAction SilentlyContinue }
    }
    return $dest
}

# ---------------------------------------------------------------- Rechte der Verwalter-Liste ----------------------------------------------------------------
# Belege (Microsoft Learn): FileSystemRights-Werte (System.Security.AccessControl.FileSystemRights), Get-Acl liest nur die
# NTFS-ACL der Datei bzw. des Ordners; Freigabeberechtigungen liegen am SMB-Share und kommen dort nicht vor.

function Get-HUAclBroadName([string]$Sid) {
    if ($script:HUAclBroad.Contains($Sid)) { return $script:HUAclBroad[$Sid] }
    if ($Sid -match '^S-1-5-21-\d+-\d+-\d+-513$') { return 'Domänen-Benutzer' }
    return ''
}

# Zulassen-Eintraege breiter Gruppen mit einem der Rechte aus -Mask. -ForContainer: nur Eintraege, die fuer den Ordner
# selbst gelten (nicht "nur vererben"). Verweigern-Eintraege werden nicht verrechnet.
function Get-HUAclBroadWriters($Acl, [int]$Mask, [switch]$ForContainer) {
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($r in @($Acl.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]))) {
        if ($r.AccessControlType -ne [System.Security.AccessControl.AccessControlType]::Allow) { continue }
        if ($ForContainer -and ($r.PropagationFlags -band [System.Security.AccessControl.PropagationFlags]::InheritOnly)) { continue }
        $name = Get-HUAclBroadName "$($r.IdentityReference.Value)"
        if (-not $name) { continue }
        if (([int64][int]$r.FileSystemRights -band [int64]$Mask) -eq 0) { continue }
        $how = $(if ($r.IsInherited) { 'geerbt' } else { 'direkt' })
        $txt = "$name ($($r.IdentityReference.Value), $how)"
        if (-not $out.Contains($txt)) { $out.Add($txt) }
    }
    return @($out.ToArray())
}

<#
.SYNOPSIS
    Prueft die NTFS-Rechte von Data\Verwalter.txt und des Ordners Data (auch ueber UNC, PowerShell 5.1 und 7).
.DESCRIPTION
    Warnung, wenn breite Gruppen (Jeder, Authentifizierte Benutzer, Benutzer, Domaenen-Benutzer; ueber SIDs) die Datei
    aendern, loeschen oder ihre Rechte aendern duerfen (auch geerbt), oder am Ordner Data Unterordner und Dateien
    loeschen duerfen (Vollzugriff). Freigaberechte (SMB) sind ueber Get-Acl nicht sichtbar und werden nicht geprueft.
    Laesst sich eine ACL nicht lesen: Status 'NichtPruefbar' (ohne Fehler).
.OUTPUTS
    @{ Status = 'Ok' | 'Warnung' | 'NichtPruefbar'; Text; ToolTip; Funde (string[]) }
#>
function Test-HUVerwalterAcl {
    param([Parameter(Mandatory)][string]$DataRoot)
    $smb = 'Freigaberechte (SMB) sind von hier aus nicht sichtbar und wurden nicht geprüft.'
    $funde = New-Object System.Collections.Generic.List[string]
    $file = [IO.Path]::Combine($DataRoot, 'Verwalter.txt')   # Join-Path wirft unter 5.1 bei fehlendem Laufwerk
    try {
        $dirAcl = Get-Acl -LiteralPath $DataRoot -ErrorAction Stop
        if (Test-Path -LiteralPath $file -PathType Leaf) {
            foreach ($x in @(Get-HUAclBroadWriters (Get-Acl -LiteralPath $file -ErrorAction Stop) $script:HUAclFileWrite)) { $funde.Add("Verwalter.txt ändern oder löschen: $x") }
        } else {
            # ohne Liste gilt der Rueckfall der App; wer im Ordner Dateien anlegen darf (WriteData 0x2 = CreateFiles), legt sie an
            foreach ($x in @(Get-HUAclBroadWriters $dirAcl 0x2 -ForContainer)) { $funde.Add("Verwalter.txt fehlt und darf neu angelegt werden: $x") }
        }
        foreach ($x in @(Get-HUAclBroadWriters $dirAcl $script:HUAclFolderWrite -ForContainer)) { $funde.Add("Ordner Data mit Vollzugriff (darf jede Datei löschen): $x") }
    } catch {
        return [pscustomobject]@{ Status = 'NichtPruefbar'; Text = 'Rechte der Verwalter-Liste: nicht prüfbar.'
            ToolTip = "Die Rechte von Data\Verwalter.txt ließen sich nicht lesen.`n$smb"; Funde = @()
        }
    }
    if ($funde.Count) {
        return [pscustomobject]@{ Status = 'Warnung'
            Text = 'Achtung: Breite Gruppen dürfen die Verwalter-Liste ändern oder löschen. Details im Tooltip.'
            ToolTip = ("Wer Data\Verwalter.txt ändern darf, kann sich selbst zum Verwalter machen.`n`n" + ($funde -join "`n") +
                "`n`nEmpfehlung: Data mit ""Ändern"" für die Benutzer (ohne Vollzugriff); Verwalter.txt nur Lesen für Benutzer, " +
                "Ändern nur für Verwalter, Vererbung an der Datei unterbrochen.`n$smb")
            Funde = @($funde.ToArray())
        }
    }
    return [pscustomobject]@{ Status = 'Ok'; Text = 'Rechte der Verwalter-Liste: keine breiten Schreibrechte (NTFS).'
        ToolTip = "Weder Datei noch Ordner Data geben Jeder, Authentifizierten Benutzern, Benutzern oder Domänen-Benutzern Schreibrechte auf die Liste.`n$smb"; Funde = @()
    }
}

Export-ModuleMember -Function Initialize-HUDataPaths, Assert-HUWritable, Get-HUModuleDataPath, Get-HUSafeName,
    Open-HURunLock, Close-HURunLock, Get-HURunningInstances,
    Lock-HUDataFile, Unlock-HUDataFile, Get-HULockInfoText, Backup-HUDataFile, Test-HUVerwalterAcl
