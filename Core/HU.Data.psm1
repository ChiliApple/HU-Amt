#Requires -Version 5.1
<#
.SYNOPSIS
    HU.Data - gemeinsame Daten-Bausteine fuer Module: JSON-Ablage, Speichern mit Sperre, Kennungen, Nummernkreis,
    Zwischenspeicher, Zahlen/Euro/Datum fuer Oesterreich, Verlauf, Anhaenge, gemeinsame Einstellungen.
.DESCRIPTION
    Regeln:
      - Laeuft ohne Fenster und ohne Functions\ (auch im nackten Hintergrund-Runspace). Keine $script:-Variablen der App.
      - Schreiben nutzt HU.Paths (Sperren, Kopien) und HU.Exchange (Write-HUFileAtomic). Fehlen sie, werden sie aus
        demselben Ordner geladen (nur in dieses Modul, nicht global).
      - JSON wird unter PowerShell 7 ohne Datums-Umwandlung gelesen (Texte bleiben Texte wie unter 5.1).
        -DateKind gibt es erst ab PowerShell 7.5; unter 7.0 bis 7.4 kommen Datumstexte weiter als DateTime an,
        darum Werte immer ueber ConvertTo-HUTimeText / ConvertTo-HUDateText lesen.
      - Zahlen und Euro immer mit fester Schreibweise (Komma, Tausenderpunkt). de-AT liefert unter PowerShell 7 (ICU)
        ein geschuetztes Leerzeichen als Tausendertrennung, unter 5.1 einen Punkt; darum ein eigenes Zahlenformat.
      - Ein Datensatz je Datei, geschrieben nur ueber Save-HUDataFile (Sperre, Warten, Stand-Pruefung, Kopie, atomar).
      - Geloescht wird nichts: Anhaenge kommen beim Entfernen nach _entfernt\.
.NOTES
    Zielmaschine: der PC, auf dem die App laeuft. PowerShell 5.1 und 7.
#>

Set-StrictMode -Version 1.0

# ---------------------------------------------------------------- Startwerte ----------------------------------------------------------------

$script:HUDataInv = [Globalization.CultureInfo]::InvariantCulture
# Datum lesen: zweistellige Jahre fest bis 2049 (KERN-B6), unabhaengig von der Windows-Einstellung des PCs
$script:HUDataDateCulture = [Globalization.CultureInfo]([Globalization.CultureInfo]::InvariantCulture.Clone())
$script:HUDataDateCal = New-Object Globalization.GregorianCalendar
$script:HUDataDateCal.TwoDigitYearMax = 2049
$script:HUDataDateCulture.DateTimeFormat.Calendar = $script:HUDataDateCal
$script:HUDataNum = [Globalization.NumberFormatInfo]([Globalization.CultureInfo]::InvariantCulture.NumberFormat.Clone())
$script:HUDataNum.NumberDecimalSeparator = ','
$script:HUDataNum.NumberGroupSeparator = '.'
$script:HUDataEuro = [string][char]0x20AC
$script:HUDataChanged = 'Die Daten wurden inzwischen an einem anderen PC geändert. Bitte neu laden, es wurde nichts überschrieben.'
$script:HUDataJsonDateKind = [bool]((Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind'))
$script:HUDataDisplayName = $null
$script:HUDataDateFormats = [string[]]@(
    'd.M.yyyy', 'd.M.yy', 'yyyy-MM-dd', 'yyyy-MM-ddTHH:mm', 'yyyy-MM-ddTHH:mm:ss', 'yyyy-MM-dd HH:mm', 'yyyy-MM-dd HH:mm:ss',
    'd.M.yyyy H:mm', 'd.M.yyyy H:mm:ss', 'd.M.yy H:mm')

# Kern-Module fuer das Schreiben: nur laden, wenn sie nicht schon da sind (App, Test oder Runspace)
foreach ($m in @(@{ Name = 'HU.Paths'; Probe = 'Lock-HUDataFile' }, @{ Name = 'HU.Exchange'; Probe = 'Write-HUFileAtomic' })) {
    if (-not (Get-Command $m.Probe -ErrorAction SilentlyContinue)) {
        $mp = Join-Path $PSScriptRoot ($m.Name + '.psm1')
        if (Test-Path -LiteralPath $mp) { Import-Module $mp -DisableNameChecking }
    }
}

# ---------------------------------------------------------------- Hilfen (intern) ----------------------------------------------------------------

function Assert-HUDataWriter {
    foreach ($c in 'Lock-HUDataFile', 'Unlock-HUDataFile', 'Get-HULockInfoText', 'Backup-HUDataFile', 'Write-HUFileAtomic') {
        if (-not (Get-Command $c -ErrorAction SilentlyContinue)) { throw "Kern-Funktion $c fehlt (HU.Paths/HU.Exchange nicht geladen)" }
    }
}
function Get-HUDataFileBase([string]$Name) {
    # wie Get-HUSafeName, dazu Geraetenamen von Windows (CON, NUL, COM1 ...) entschaerfen
    $n = ("$Name" -replace '[^\p{L}\p{Nd}\.\-_]', '_').Trim('.', ' ')
    if (-not $n) { $n = 'unbekannt' }
    if ($n -match '^(?i)(CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9])(\..*)?$') { $n = '_' + $n }
    return $n
}
function New-HUDataResult([bool]$Ok, [string]$Message, $Stamp, [bool]$Skipped = $false) {
    return [pscustomobject]@{ Ok = $Ok; Message = $Message; Stamp = $Stamp; Skipped = $Skipped }
}
function Test-HUDataIOException($Err) {
    $ex = $Err.Exception
    while ($ex) {
        if ($ex -is [System.IO.IOException] -or $ex -is [System.UnauthorizedAccessException]) { return $true }
        $ex = $ex.InnerException
    }
    return $false
}

# ================================================================ JSON und Dateien ================================================================

<#
.SYNOPSIS
    Jetzt als Text JJJJ-MM-TTTHH:MM:SS (fuer zeit, geaendert, angelegt).
#>
function Get-HUNowText([datetime]$Now = (Get-Date)) { return $Now.ToString('yyyy-MM-ddTHH:mm:ss', $script:HUDataInv) }

<#
.SYNOPSIS
    Zeitstempel als Text JJJJ-MM-TTTHH:MM:SS. Nimmt DateTime (PowerShell 7 bis 7.4 wandelt beim JSON-Lesen um) oder Text (bleibt).
#>
function ConvertTo-HUTimeText($Value) {
    if ($null -eq $Value) { return '' }
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-ddTHH:mm:ss', $script:HUDataInv) }
    return "$Value"
}

<#
.SYNOPSIS
    Stand einer Datei: '<LastWriteTimeUtc.Ticks>|<Laenge>' oder '' (fehlt).
    FileInfo wirft nicht, wenn die Datei gerade von einem anderen PC atomar ersetzt wird (kurz nicht vorhanden).
#>
function Get-HUFileStamp([string]$Path) {
    if (-not $Path) { return '' }
    $i = New-Object IO.FileInfo $Path
    if (-not $i.Exists) { return '' }
    return ('{0}|{1}' -f $i.LastWriteTimeUtc.Ticks, $i.Length)
}

<#
.SYNOPSIS
    JSON-Text lesen. Unter PowerShell ab 7.5 ohne Datums-Umwandlung (-DateKind String), Texte bleiben Texte wie unter 5.1.
#>
function ConvertFrom-HUJsonText([string]$Text) {
    if ($script:HUDataJsonDateKind) { return ($Text | ConvertFrom-Json -DateKind String) }
    return ($Text | ConvertFrom-Json)
}

<#
.SYNOPSIS
    Objekt als JSON-Text (Standard Tiefe 10).
#>
function ConvertTo-HUJsonText($Object, [int]$Depth = 10, [switch]$Compress) {
    return (ConvertTo-Json -InputObject $Object -Depth $Depth -Compress:$Compress)
}

<#
.SYNOPSIS
    Tiefe Kopie ueber JSON (Datumstexte bleiben Texte).
#>
function Copy-HUDataObject($Object, [int]$Depth = 12) {
    if ($null -eq $Object) { return $null }
    return (ConvertFrom-HUJsonText (ConvertTo-Json -InputObject $Object -Depth $Depth))
}

<#
.SYNOPSIS
    JSON-Datei lesen. Rueckgabe: @{ Data; Error } (fehlt: Data $null, Error '').
    Wird die Datei gerade ersetzt, fehlt sie kurz oder ist gesperrt: bis zu 6 Versuche.
#>
function Read-HUJson([string]$Path) {
    if (-not $Path) { return @{ Data = $null; Error = '' } }
    for ($try = 1; $try -le 6; $try++) {
        if (-not (Test-Path -LiteralPath $Path)) {
            if ($try -ge 3) { return @{ Data = $null; Error = '' } }
            Start-Sleep -Milliseconds 40; continue
        }
        try {
            $t = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
            if (-not $t.Trim()) { return @{ Data = $null; Error = 'Datei ist leer' } }
            return @{ Data = (ConvertFrom-HUJsonText $t); Error = '' }
        } catch {
            if (-not (Test-HUDataIOException $_)) { return @{ Data = $null; Error = $_.Exception.Message } }
            if ($try -ge 6) { return @{ Data = $null; Error = $_.Exception.Message } }
            Start-Sleep -Milliseconds (40 * $try)
        }
    }
    return @{ Data = $null; Error = '' }
}

<#
.SYNOPSIS
    Eigenschaft lesen mit Standard. Nimmt PSObject, Hashtable und [ordered]. Fehlt sie oder ist sie $null: -Default.
#>
function Get-HUProp($Object, [string]$Name, $Default = '') {
    if ($null -eq $Object) { return $Default }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name) -and $null -ne $Object[$Name]) { return $Object[$Name] }
        return $Default
    }
    $p = $Object.PSObject.Properties[$Name]
    if ($p -and $null -ne $p.Value) { return $p.Value }
    return $Default
}

# ================================================================ Sperre und Speichern ================================================================

<#
.SYNOPSIS
    Sperre einer gemeinsamen Datei holen und bis -WaitMs warten (ein anderer PC speichert nur Millisekunden).
    Ein kurzes "Zugriff verweigert" beim Freigeben gilt als belegt. Legt den Ordner an.
    Rueckgabe: Sperre aus Lock-HUDataFile (Ok, UserName, ComputerName, Since ...). Freigeben mit Unlock-HUDataFile.
#>
function Wait-HUDataFileLock {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$UserName = $env:USERNAME,
        [string]$PcName = $env:COMPUTERNAME,
        [int]$WaitMs = 10000
    )
    Assert-HUDataWriter
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
    $until = (Get-Date).AddMilliseconds($WaitMs)
    do {
        try { $lock = Lock-HUDataFile -Path $Path -UserName $UserName -PcName $PcName }
        catch { $lock = [pscustomobject]@{ Ok = $false; Path = $Path; LockPath = "$Path.lock"; Stream = $null; UserName = ''; ComputerName = ''; Since = $null } }
        if ($lock.Ok -or (Get-Date) -ge $until) { break }
        Start-Sleep -Milliseconds (Get-Random -Minimum 30 -Maximum 150)
    } while ($true)
    return $lock
}

<#
.SYNOPSIS
    Text in eine Datei schreiben (atomar), bei kurzem "Zugriff verweigert" bis zu 6 Versuche.
    -New: Datei darf noch nicht existieren (wirft sonst). -Backup: vorher Kopie nach _Kopien\ (nur wenn die Datei existiert).
    Ohne Sperre. Fuer gemeinsame Dateien Save-HUDataFile nehmen.
#>
function Write-HUDataFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Text = '',
        [byte[]]$Bytes = $null,
        [switch]$New,
        [switch]$Backup,
        [int]$Keep = 20
    )
    Assert-HUDataWriter
    for ($try = 1; $try -le 6; $try++) {
        try {
            $exists = Test-Path -LiteralPath $Path
            if ($New -and $exists) { throw "Datei gibt es schon: $Path" }
            if ($Backup -and $exists -and $try -eq 1) { [void](Backup-HUDataFile -Path $Path -Keep $Keep) }
            $a = @{ Path = $Path }
            if ($null -ne $Bytes) { $a.Bytes = $Bytes } else { $a.Text = $Text }
            if (-not $New) { $a.Overwrite = $true }
            [void](Write-HUFileAtomic @a)
            return
        } catch {
            if (-not (Test-HUDataIOException $_)) { throw }
            if ($New -and (Test-Path -LiteralPath $Path)) { throw }
            if ($try -ge 6) { throw }
            Start-Sleep -Milliseconds (50 * $try)
        }
    }
}

<#
.SYNOPSIS
    Gemeinsame Datei speichern: Sperre (bis -WaitMs warten), Stand pruefen, Text bauen, Kopie, atomar schreiben.
.DESCRIPTION
    -Build liefert den Inhalt: Text (wird so geschrieben), ein Objekt (wird JSON mit -Depth) oder $null (= nichts zu tun,
    Rueckgabe Ok mit Skipped). Wirft -Build, wird nichts geschrieben und der Text ist die Meldung.
    -Build laeuft unter der Sperre: dort neu lesen und aendern, dann ist kein Stand verloren.
    -Stamp: erwarteter Stand aus Get-HUFileStamp; '' = Datei darf noch nicht existieren; $null = nicht pruefen.
    Neue Datei: ohne Ueberschreiben. Vorhandene: Kopie nach _Kopien\ (ausser -NoBackup, -Keep Kopien), dann ersetzen.
    -Lock: schon gehaltene Sperre (z. B. Verwaltungsfenster); dann wird nicht neu gesperrt und nicht freigegeben.
    Rueckgabe: Ok, Message, Stamp (neuer Stand, bei Fehler der uebergebene), Skipped.
#>
function Save-HUDataFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][scriptblock]$Build,
        $Stamp = $null,
        [string]$UserName = $env:USERNAME,
        [string]$PcName = $env:COMPUTERNAME,
        [string]$Changed = $script:HUDataChanged,
        [int]$WaitMs = 10000,
        [switch]$NoBackup,
        [int]$Keep = 20,
        $Lock = $null,
        [int]$Depth = 10
    )
    Assert-HUDataWriter
    $own = $null
    if (-not ($Lock -and $Lock.Ok)) {
        $own = Wait-HUDataFileLock -Path $Path -UserName $UserName -PcName $PcName -WaitMs $WaitMs
        if (-not $own.Ok) { return (New-HUDataResult $false (Get-HULockInfoText $own) $Stamp) }
    }
    try {
        if ($null -ne $Stamp -and (Get-HUFileStamp $Path) -ne "$Stamp") { return (New-HUDataResult $false $Changed $Stamp) }
        $out = & $Build
        if ($null -eq $out) { return (New-HUDataResult $true '' (Get-HUFileStamp $Path) $true) }
        $text = $(if ($out -is [string]) { $out } else { ConvertTo-Json -InputObject $out -Depth $Depth })
        $exists = Test-Path -LiteralPath $Path
        Write-HUDataFile -Path $Path -Text $text -New:(-not $exists) -Backup:($exists -and -not $NoBackup) -Keep $Keep
        return (New-HUDataResult $true '' (Get-HUFileStamp $Path))
    } catch { return (New-HUDataResult $false $_.Exception.Message $Stamp) }
    finally { if ($own) { Unlock-HUDataFile $own } }
}

# ================================================================ Kennungen und Nummernkreis ================================================================

<#
.SYNOPSIS
    Kennung JJJJMMTT-HHMMSS-<PC>-<8 Zeichen>: sortierbar nach Zeit, eindeutig ueber PCs (PC nur A-Z0-9, hoechstens 15).
#>
function New-HUDataId([string]$PcName = $env:COMPUTERNAME, [datetime]$Now = (Get-Date)) {
    $pc = ("$PcName" -replace '[^A-Za-z0-9]', '')
    if (-not $pc) { $pc = 'PC' }
    if ($pc.Length -gt 15) { $pc = $pc.Substring(0, 15) }
    return ('{0}-{1}-{2}' -f $Now.ToString('yyyyMMdd-HHmmss', $script:HUDataInv), $pc.ToUpperInvariant(), [guid]::NewGuid().ToString('N').Substring(0, 8))
}
function Test-HUDataId([string]$Id) { return ("$Id" -cmatch '^\d{8}-\d{6}-[A-Z0-9]{1,15}-[0-9a-f]{8}$') }

<#
.SYNOPSIS
    Kurze Kennung <Praefix>-<8 Zeichen>, z. B. f-1a2b3c4d.
#>
function New-HUShortId([Parameter(Mandatory)][string]$Prefix) {
    if ($Prefix -notmatch '^[a-z][a-z0-9]{0,15}$') { throw "Ungültiges Präfix: '$Prefix'" }
    return ($Prefix + '-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
}
<#
.SYNOPSIS
    Prueft eine kurze Kennung. -Prefix leer = jedes Praefix aus Kleinbuchstaben. -Loose: 1 bis 16 Zeichen nach dem Strich
    (aeltere Kennungen), sonst genau 8.
#>
function Test-HUShortId([string]$Id, [string]$Prefix = '', [switch]$Loose) {
    $p = $(if ($Prefix) { [regex]::Escape($Prefix) } else { '[a-z][a-z0-9]{0,15}' })
    $n = $(if ($Loose) { '{1,16}' } else { '{8}' })
    return ("$Id" -cmatch ('^' + $p + '-[a-z0-9]' + $n + '$'))
}

<#
.SYNOPSIS
    Naechste Nummer aus einem Zaehler (JSON { "<Schluessel>": <letzte Nummer> }) unter Sperre.
    -Format: {0} = Schluessel, {1} = Nummer (Standard 2026-0001). -Exists: bekommt die Nummer als Text, $true = schon
    vergeben (z. B. Datei da, Zaehler nach einer Wiederherstellung zu klein), dann wird weitergezaehlt.
    Rueckgabe: Nummer als Text. Wirft, wenn die Sperre nicht frei wird oder der Zaehler nicht lesbar ist.
#>
function Get-HUNextNumber {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Key = "$((Get-Date).Year)",
        [string]$Format = '{0}-{1:D4}',
        [scriptblock]$Exists = $null,
        [string]$UserName = $env:USERNAME,
        [string]$PcName = $env:COMPUTERNAME,
        [int]$WaitMs = 10000
    )
    $lock = Wait-HUDataFileLock -Path $Path -UserName $UserName -PcName $PcName -WaitMs $WaitMs
    if (-not $lock.Ok) { throw (Get-HULockInfoText $lock) }
    try {
        $r = Read-HUJson $Path
        if ($r.Error) { throw "Zähler nicht lesbar: $($r.Error)" }
        $h = [ordered]@{}
        if ($r.Data) { foreach ($q in $r.Data.PSObject.Properties) { $h[$q.Name] = [int]$q.Value } }
        $n = 0
        if ($h.Contains($Key)) { $n = [int]$h[$Key] }
        $n++
        $text = $Format -f $Key, $n
        if ($Exists) { while (& $Exists $text) { $n++; $text = $Format -f $Key, $n } }
        $h[$Key] = $n
        Write-HUDataFile -Path $Path -Text (ConvertTo-Json -InputObject ([pscustomobject]$h) -Depth 3)
        return $text
    } finally { Unlock-HUDataFile $lock }
}

# ================================================================ Datensaetze mit Zwischenspeicher ================================================================

# Fehlt ein Ordner wirklich (dann leer) oder ist er nur nicht erreichbar (KERN-B4)?
#   - Wurzel (Laufwerk oder \\server\freigabe) nicht da: nicht erreichbar.
#   - Elternordner da: der Ordner fehlt wirklich.
#   - Elternordner fehlt: nicht erreichbar, wenn der Zwischenspeicher dort schon Datensaetze kennt (-Known).
#     Ohne bekannte Datensaetze fehlt er wirklich (z. B. Unterordner, die erst beim ersten Speichern entstehen).
function Test-HUDataFolderMissing([string]$Dir, [bool]$Known = $false) {
    $full = [IO.Path]::GetFullPath($Dir)
    $root = [IO.Path]::GetPathRoot($full)
    if ($root -and -not [IO.Directory]::Exists($root)) { return $false }
    $parent = [IO.Path]::GetDirectoryName($full.TrimEnd('\', '/'))
    if (-not $parent -or [IO.Directory]::Exists($parent)) { return $true }
    return (-not $Known)
}

<#
.SYNOPSIS
    Alle Datensaetze (eine JSON-Datei je Datensatz) aus einem oder mehreren Ordnern lesen, nur Geaendertes neu.
.DESCRIPTION
    -Cache: Hashtable, die der Aufrufer zwischen den Aufrufen behaelt (wird hier fortgeschrieben und auch zurueckgegeben).
      Schluessel = voller Pfad, Wert @{ Stamp; Item }. Neu gelesen wird nur, wenn sich der Stand (Zeit|Groesse) geaendert hat.
      -Immutable: Datensaetze aendern sich nie (z. B. Vorgaenge), schon gelesene werden nie neu gelesen.
    -Convert: bekommt ($Data, $FileInfo) und liefert den Datensatz; $null = ungueltig (Fehler), Werfen = Fehler mit Text.
      Ohne -Convert ist der Datensatz das gelesene Objekt.
    -NamePattern: Regex fuer den Dateinamen ohne .json (z. B. '^f-[a-z0-9]{8}$'); andere Dateien werden uebergangen.
    Dateien mit Punkt am Anfang (Temp-Dateien) werden uebergangen. Ist eine Datei gerade nicht lesbar, bleibt der alte
    Stand aus dem Zwischenspeicher. Verschwundene Dateien fallen aus dem Zwischenspeicher (nur in den gelesenen Ordnern).
    Ist ein Ordner nicht erreichbar (Wurzel weg, oder Elternordner weg und der Zwischenspeicher kennt dort Datensaetze),
    kommt in Errors "Ordner nicht erreichbar: <Pfad>", der Zwischenspeicher bleibt und seine Datensaetze kommen als Items.
    Leer (ohne Fehler) nur, wenn der Ordner wirklich fehlt (Elternordner da, oder noch nie Datensaetze dort).
    Rueckgabe: Items, Stamps (Pfad -> Stand), Errors, Cache, Gelesen, Neu, Weg, Ms.
#>
function Get-HUDataRecords {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$Folder,
        [string]$Filter = '*.json',
        [hashtable]$Cache = $null,
        [scriptblock]$Convert = $null,
        [string]$NamePattern = '',
        [switch]$Immutable
    )
    $sw = [Diagnostics.Stopwatch]::StartNew()
    if ($null -eq $Cache) { $Cache = @{} }
    $items = New-Object System.Collections.Generic.List[object]
    $stamps = @{}
    $errors = New-Object System.Collections.Generic.List[string]
    $read = 0; $neu = 0; $weg = 0
    foreach ($dir in @($Folder | Where-Object { $_ } | Select-Object -Unique)) {
        $prefix = [IO.Path]::GetFullPath($dir).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
        $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        $files = @()
        if ([IO.Directory]::Exists($dir)) {
            try { $files = @((New-Object IO.DirectoryInfo $dir).GetFiles($Filter) | Sort-Object Name) }
            catch { $errors.Add("Ordner nicht lesbar: $($_.Exception.Message)"); continue }
        }
        elseif (-not (Test-HUDataFolderMissing $dir (@($Cache.Keys | Where-Object { "$_".StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) }).Count -gt 0))) {
            # KERN-B4: Freigabe kurz weg o. Ae. -> melden, Zwischenspeicher behalten und dessen Stand liefern
            $errors.Add("Ordner nicht erreichbar: $dir")
            foreach ($k in @($Cache.Keys | Where-Object { "$_".StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) } | Sort-Object)) {
                $stamps[$k] = $Cache[$k].Stamp
                $items.Add($Cache[$k].Item)
            }
            continue
        }
        foreach ($fi in $files) {
            if ($fi.Name.StartsWith('.')) { continue }
            if ($NamePattern -and [IO.Path]::GetFileNameWithoutExtension($fi.Name) -notmatch $NamePattern) { continue }
            $key = $fi.FullName
            [void]$seen.Add($key)
            $st = '{0}|{1}' -f $fi.LastWriteTimeUtc.Ticks, $fi.Length
            $hit = $(if ($Cache.ContainsKey($key)) { $Cache[$key] } else { $null })
            $item = $null
            if ($hit -and ($Immutable -or $hit.Stamp -eq $st)) { $item = $hit.Item; $st = $hit.Stamp }
            else {
                $read++
                $r = Read-HUJson $fi.FullName
                if (-not $r.Data) {
                    if ($r.Error) { $errors.Add("$($fi.Name): $($r.Error)") }
                    if ($hit) { $item = $hit.Item; $st = $hit.Stamp } else { continue }
                } else {
                    try {
                        $item = $(if ($Convert) { & $Convert $r.Data $fi } else { $r.Data })
                        if ($null -eq $item) { throw 'ungültiger Datensatz' }
                    } catch {
                        $errors.Add("$($fi.Name): $($_.Exception.Message)")
                        if ($hit) { $item = $hit.Item; $st = $hit.Stamp } else { continue }
                    }
                    if (-not $hit) { $neu++ }
                    $Cache[$key] = @{ Stamp = $st; Item = $item }
                }
            }
            $stamps[$key] = $st
            $items.Add($item)
        }
        foreach ($k in @($Cache.Keys)) {
            if ("$k".StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) -and -not $seen.Contains($k)) { $Cache.Remove($k); $weg++ }
        }
    }
    $sw.Stop()
    return [pscustomobject]@{ Items = $items.ToArray(); Stamps = $stamps; Errors = $errors.ToArray(); Cache = $Cache; Gelesen = $read; Neu = $neu; Weg = $weg; Ms = $sw.ElapsedMilliseconds }
}

# ================================================================ Zahlen, Euro, Datum (Oesterreich) ================================================================

<#
.SYNOPSIS
    Zahl aus Eingabe oder Datei: 1.234,56 / 1234,56 / 1234.56 / 12,50 EUR-Zeichen / Zahlentypen. Sonst -Default.
    Mit Komma: Komma = Dezimalzeichen, Punkte davor = Tausender. Ohne Komma: Punkt = Dezimalzeichen.
#>
function ConvertTo-HUDecimal($Value, [decimal]$Default = 0) {
    if ($null -eq $Value) { return $Default }
    if ($Value -is [decimal]) { return $Value }
    if ($Value -is [double] -or $Value -is [single] -or $Value -is [int] -or $Value -is [long] -or $Value -is [int16] -or $Value -is [byte]) {
        try { return [decimal]$Value } catch { return $Default }
    }
    $t = ("$Value" -replace ('[\s' + $script:HUDataEuro + ']'), '').Trim()
    if (-not $t) { return $Default }
    if ($t -match ',') { $t = $t.Replace('.', '').Replace(',', '.') }
    $d = [decimal]0
    if ([decimal]::TryParse($t, [Globalization.NumberStyles]::Number, $script:HUDataInv, [ref]$d)) { return $d }
    return $Default
}

<#
.SYNOPSIS
    Ist die Eingabe zweideutig? Genau ein Punkt mit genau drei Ziffern danach, kein Komma (1.500, 12.250, -3.000):
    ConvertTo-HUDecimal liest das als 1,5, gemeint sein kann aber 1500. Fuer Eingabefelder (Rueckfrage mit
    Read-HUNumberInput im Fenster); Daten aus Dateien liest ConvertTo-HUDecimal unveraendert.
#>
function Test-HUAmbiguousNumber([string]$Text) {
    $t = ("$Text" -replace ('[\s' + $script:HUDataEuro + ']'), '').Trim()
    return ($t -match '^[+-]?\d{1,3}\.\d{3}$')
}

<#
.SYNOPSIS
    Ist der Text leer oder eine Zahl im Sinn von ConvertTo-HUDecimal?
#>
function Test-HUNumberText([string]$Text) {
    $t = ("$Text" -replace ('[\s' + $script:HUDataEuro + ']'), '').Trim()
    if (-not $t) { return $true }
    if ($t -match ',') { $t = $t.Replace('.', '').Replace(',', '.') }
    $d = [decimal]0
    return [decimal]::TryParse($t, [Globalization.NumberStyles]::Number, $script:HUDataInv, [ref]$d)
}

<#
.SYNOPSIS
    Kaufmaennisch runden (0,5 immer weg von null).
#>
function Get-HURound([decimal]$Value, [int]$Digits = 2) { return [Math]::Round($Value, $Digits, [MidpointRounding]::AwayFromZero) }

function Get-HUNumberPattern([int]$Digits, [bool]$Fixed, [bool]$Grouping) {
    $p = $(if ($Grouping) { '#,##0' } else { '0' })
    if ($Digits -gt 0) { $p += '.' + ($(if ($Fixed) { '0' } else { '#' }) * $Digits) }
    return $p
}

<#
.SYNOPSIS
    Euro-Betrag: 1.234,56 EUR-Zeichen (immer 2 Nachkommastellen). -NoGrouping: ohne Tausenderpunkt.
#>
function Format-HUEuro([decimal]$Value, [switch]$NoGrouping) {
    return ($Value.ToString((Get-HUNumberPattern 2 $true (-not $NoGrouping)), $script:HUDataNum) + ' ' + $script:HUDataEuro)
}

<#
.SYNOPSIS
    Zahl: 1.234,5 (bis -Digits Nachkommastellen, ohne Nullen am Ende). -Fixed: immer -Digits Stellen. -NoGrouping: ohne Tausenderpunkt.
#>
function Format-HUNumber([decimal]$Value, [int]$Digits = 2, [switch]$Fixed, [switch]$NoGrouping) {
    return $Value.ToString((Get-HUNumberPattern $Digits ([bool]$Fixed) (-not $NoGrouping)), $script:HUDataNum)
}

<#
.SYNOPSIS
    Kilometer: 1.234,5 km (eine Nachkommastelle, ohne Null am Ende). -NoGrouping: ohne Tausenderpunkt.
#>
function Format-HUKm([decimal]$Value, [switch]$NoGrouping) {
    return ($Value.ToString((Get-HUNumberPattern 1 $false (-not $NoGrouping)), $script:HUDataNum) + ' km')
}

<#
.SYNOPSIS
    Datum aus Eingabe oder Datei. Nimmt DateTime, 6.11.2026, 06.11.26, 2026-11-06, mit Uhrzeit (2026-11-06T14:30,
    6.11.2026 14:30), ISO mit Sekundenbruchteil/Zone und Excel-Seriennummern (20000 bis 80000, z. B. 46332).
    Die Uhrzeit bleibt erhalten; -DateOnly schneidet sie ab. Sonst $null.
#>
function ConvertFrom-HUDateText($Text, [switch]$DateOnly) {
    $d = $null
    if ($Text -is [datetime]) { $d = [datetime]$Text }
    else {
        $t = "$Text".Trim()
        if (-not $t) { return $null }
        $x = [datetime]::MinValue
        if ([datetime]::TryParseExact($t, $script:HUDataDateFormats, $script:HUDataDateCulture, [Globalization.DateTimeStyles]::None, [ref]$x)) { $d = $x }
        elseif ($t -match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}' -and [datetime]::TryParse($t, $script:HUDataInv, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$x)) {
            # ISO mit Bruchteil oder Zone (z. B. von ConvertTo-Json unter PowerShell 7): als Ortszeit
            $d = $(if ($x.Kind -eq [DateTimeKind]::Utc) { $x.ToLocalTime() } else { $x })
            if ($d.Kind -ne [DateTimeKind]::Unspecified) { $d = [datetime]::SpecifyKind($d, [DateTimeKind]::Unspecified) }
        } else {
            $n = 0.0
            if ($t -match '^\d+([.,]\d+)?$' -and [double]::TryParse($t.Replace(',', '.'), [Globalization.NumberStyles]::Float, $script:HUDataInv, [ref]$n) -and $n -ge 20000 -and $n -le 80000) {
                $d = [datetime]::FromOADate($n)
            }
        }
    }
    if ($null -eq $d) { return $null }
    if ($DateOnly) { return $d.Date }
    return $d
}

<#
.SYNOPSIS
    Datum als Text JJJJ-MM-TT (fuer Dateien) oder '' (kein Datum).
#>
function ConvertTo-HUDateText($Value) {
    $d = ConvertFrom-HUDateText $Value
    if ($null -eq $d) { return '' }
    return $d.ToString('yyyy-MM-dd', $script:HUDataInv)
}

<#
.SYNOPSIS
    Uhrzeit aus Eingabe: 7:30, 07:30, 730, 7.30, 7,30, 7 -> TimeSpan. Sonst $null.
#>
function ConvertFrom-HUClockText([string]$Text) {
    $t = "$Text".Trim().Replace('.', ':').Replace(',', ':')
    if (-not $t) { return $null }
    if ($t -match '^(\d{1,2})$') { $t = "$($Matches[1]):00" }
    elseif ($t -match '^(\d{1,2})(\d{2})$') { $t = "$($Matches[1]):$($Matches[2])" }
    if ($t -notmatch '^(\d{1,2}):(\d{2})$') { return $null }
    $h = [int]$Matches[1]; $m = [int]$Matches[2]
    if ($h -gt 23 -or $m -gt 59) { return $null }
    return (New-Object TimeSpan $h, $m, 0)
}

<#
.SYNOPSIS
    Datum zur Anzeige TT.MM.JJJJ. Kein Datum: der Text unveraendert (-EmptyIfInvalid: '').
#>
function Format-HUDate($Value, [switch]$EmptyIfInvalid) {
    $d = ConvertFrom-HUDateText $Value
    if ($d) { return $d.ToString('dd.MM.yyyy', $script:HUDataInv) }
    if ($EmptyIfInvalid -or $null -eq $Value) { return '' }
    return "$Value"
}

<#
.SYNOPSIS
    Zeitpunkt zur Anzeige TT.MM.JJJJ HH:MM. Kein Datum: der Text unveraendert (-EmptyIfInvalid: '').
#>
function Format-HUTime($Value, [switch]$EmptyIfInvalid) {
    $d = ConvertFrom-HUDateText $Value
    if ($d) { return $d.ToString('dd.MM.yyyy HH:mm', $script:HUDataInv) }
    if ($EmptyIfInvalid -or $null -eq $Value) { return '' }
    return (ConvertTo-HUTimeText $Value)
}

# ================================================================ Verlauf ================================================================

<#
.SYNOPSIS
    Verlauf-Eintrag: zeit, benutzer, pc, aktion, text[, felder{feld, alt, neu}] und wahlweise weitere Werte (-Extra).
#>
function New-HUHistoryEntry {
    [CmdletBinding()]
    param(
        [string]$Action = '',
        [string]$Text = '',
        [object[]]$Fields = @(),
        [System.Collections.IDictionary]$Extra = $null,
        [string]$UserName = $env:USERNAME,
        [string]$PcName = $env:COMPUTERNAME,
        [datetime]$Now = (Get-Date)
    )
    $o = [ordered]@{ zeit = (Get-HUNowText $Now); benutzer = "$UserName"; pc = "$PcName"; aktion = "$Action"; text = "$Text" }
    $f = @($Fields | Where-Object { $null -ne $_ })
    if ($f.Count) { $o.felder = @(foreach ($z in $f) { [pscustomobject][ordered]@{ feld = "$(Get-HUProp $z 'feld')"; alt = "$(Get-HUProp $z 'alt')"; neu = "$(Get-HUProp $z 'neu')" } }) }
    if ($Extra) { foreach ($k in @($Extra.Keys)) { if (-not $o.Contains("$k")) { $o["$k"] = $Extra[$k] } } }
    return [pscustomobject]$o
}

<#
.SYNOPSIS
    Eintrag an den Verlauf eines Datensatzes haengen (Eigenschaft -Property, Standard verlauf; fehlt sie, wird sie angelegt).
    Rueckgabe: der neue Eintrag.
#>
function Add-HUHistoryEntry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Record,
        [string]$Action = '',
        [string]$Text = '',
        [object[]]$Fields = @(),
        [System.Collections.IDictionary]$Extra = $null,
        [string]$UserName = $env:USERNAME,
        [string]$PcName = $env:COMPUTERNAME,
        [datetime]$Now = (Get-Date),
        [string]$Property = 'verlauf'
    )
    $e = New-HUHistoryEntry -Action $Action -Text $Text -Fields $Fields -Extra $Extra -UserName $UserName -PcName $PcName -Now $Now
    $old = @(Get-HUProp $Record $Property @() | Where-Object { $null -ne $_ })
    $new = @($old) + @($e)
    if ($Record -is [System.Collections.IDictionary]) { $Record[$Property] = $new }
    elseif ($Record.PSObject.Properties[$Property]) { $Record.$Property = $new }
    else { $Record | Add-Member -NotePropertyName $Property -NotePropertyValue $new }
    return $e
}

<#
.SYNOPSIS
    Gelesenen Verlauf-Eintrag in die feste Form bringen. Liest auch alte Felder: von -> benutzer, art -> aktion.
    Weitere Werte (z. B. faellig, anhang) bleiben erhalten.
#>
function ConvertTo-HUHistoryEntry($Entry) {
    $o = [ordered]@{
        zeit     = (ConvertTo-HUTimeText (Get-HUProp $Entry 'zeit' ''))
        benutzer = "$(Get-HUProp $Entry 'benutzer' (Get-HUProp $Entry 'von' ''))"
        pc       = "$(Get-HUProp $Entry 'pc' '')"
        aktion   = "$(Get-HUProp $Entry 'aktion' (Get-HUProp $Entry 'art' ''))"
        text     = "$(Get-HUProp $Entry 'text' '')"
    }
    $f = @(Get-HUProp $Entry 'felder' @() | Where-Object { $null -ne $_ })
    if ($f.Count) { $o.felder = @(foreach ($z in $f) { [pscustomobject][ordered]@{ feld = "$(Get-HUProp $z 'feld')"; alt = "$(Get-HUProp $z 'alt')"; neu = "$(Get-HUProp $z 'neu')" } }) }
    $names = @()
    if ($Entry -is [System.Collections.IDictionary]) { $names = @($Entry.Keys | ForEach-Object { "$_" }) }
    elseif ($null -ne $Entry) { $names = @($Entry.PSObject.Properties | ForEach-Object { $_.Name }) }
    foreach ($n in $names) {
        if ($n -in 'zeit', 'benutzer', 'von', 'pc', 'aktion', 'art', 'text', 'felder') { continue }
        $o[$n] = Get-HUProp $Entry $n $null
    }
    return [pscustomobject]$o
}

function ConvertTo-HUFieldText($Value) {
    if ($null -eq $Value) { return '' }
    if ($Value -is [datetime]) { return (Format-HUTime $Value) }
    if ($Value -is [bool]) { if ($Value) { return 'ja' } else { return 'nein' } }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string] -and $Value -isnot [System.Collections.IDictionary]) { return (@($Value | ForEach-Object { "$_" }) -join ', ') }
    return "$Value"
}

<#
.SYNOPSIS
    Felder zweier Staende vergleichen. -Fields: Liste von Namen oder [ordered]@{ name = 'Anzeige' }.
    -Format: Scriptblock ($Name, $Wert) -> Anzeigetext (Standard: Listen mit Komma, ja/nein, Datum TT.MM.JJJJ HH:MM).
    Verglichen wird der Anzeigetext (Gross/klein zaehlt). Rueckgabe: Liste { feld (Anzeige), alt, neu }.
#>
function Compare-HURecordFields {
    [CmdletBinding()]
    param(
        $Old,
        $New,
        [Parameter(Mandatory)]$Fields,
        [scriptblock]$Format = $null
    )
    $map = [ordered]@{}
    if ($Fields -is [System.Collections.IDictionary]) { foreach ($k in @($Fields.Keys)) { $map["$k"] = "$($Fields[$k])" } }
    else { foreach ($k in @($Fields)) { $map["$k"] = "$k" } }
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($k in @($map.Keys)) {
        $a = Get-HUProp $Old $k $null; $b = Get-HUProp $New $k $null
        if ($Format) { $at = "$(& $Format $k $a)"; $bt = "$(& $Format $k $b)" }
        else { $at = ConvertTo-HUFieldText $a; $bt = ConvertTo-HUFieldText $b }
        if ($at -cne $bt) { $out.Add([pscustomobject][ordered]@{ feld = $map[$k]; alt = $at; neu = $bt }) }
    }
    return @($out.ToArray())
}

<#
.SYNOPSIS
    Aenderungen als ein Text: "Feld: alt -> neu; Feld: alt -> neu". -Separator aenderbar.
#>
function Get-HUFieldChangeText([object[]]$Changes, [string]$Separator = '; ') {
    $arrow = ' ' + [string][char]0x2192 + ' '
    return (@(foreach ($c in @($Changes | Where-Object { $null -ne $_ })) { "$(Get-HUProp $c 'feld'): $(Get-HUProp $c 'alt')$arrow$(Get-HUProp $c 'neu')" }) -join $Separator)
}

# ================================================================ Anhaenge ================================================================

<#
.SYNOPSIS
    Anhang pruefen: Datei da, Endung erlaubt (-Extensions, leer = alle), hoechstens -MaxMB (0 = Anhaenge aus).
    Rueckgabe: Ok, Message.
#>
function Test-HUAttachment {
    [CmdletBinding()]
    param(
        [string]$Path,
        [string[]]$Extensions = @(),
        [int]$MaxMB = 20
    )
    if (-not $Path -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { return [pscustomobject]@{ Ok = $false; Message = "Die Datei gibt es nicht: $Path" } }
    $fi = New-Object IO.FileInfo $Path
    if ($MaxMB -le 0) { return [pscustomobject]@{ Ok = $false; Message = 'Anhänge sind ausgeschaltet.' } }
    $ext = @($Extensions | Where-Object { $_ } | ForEach-Object { $e = "$_".ToLowerInvariant(); if (-not $e.StartsWith('.')) { $e = '.' + $e }; $e })
    if ($ext.Count -and $ext -notcontains $fi.Extension.ToLowerInvariant()) {
        $list = (@($ext | ForEach-Object { $_.TrimStart('.').ToUpperInvariant() }) -join ', ')
        return [pscustomobject]@{ Ok = $false; Message = "$($fi.Name): erlaubt sind nur $list." }
    }
    if ($fi.Length -gt ([long]$MaxMB * 1MB)) {
        return [pscustomobject]@{ Ok = $false; Message = ('{0}: {1} MB groß, erlaubt sind {2} MB.' -f $fi.Name, ($fi.Length / 1MB).ToString('0.0', $script:HUDataNum), $MaxMB) }
    }
    return [pscustomobject]@{ Ok = $true; Message = '' }
}

<#
.SYNOPSIS
    Anhang als Kopie in -Folder ablegen, nie ueberschreiben.
    -TimePrefix: Name JJJJMMTT-HHMMSS_<Name>.<ext>, gleicher Name in derselben Sekunde -> JJJJMMTT-HHMMSS_2_<Name>.
    Ohne: <Name>.<ext>, dann <Name>_2.<ext>, <Name>_3.<ext>. Name wird sicher gemacht und auf -MaxNameLength gekuerzt,
    Endung klein. Wirft, wenn die Pruefung (Test-HUAttachment) fehlschlaegt. Rueckgabe: Name, Path.
#>
function Add-HUAttachment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Folder,
        [Parameter(Mandatory)][string]$Path,
        [string[]]$Extensions = @(),
        [int]$MaxMB = 20,
        [switch]$TimePrefix,
        [datetime]$Now = (Get-Date),
        [int]$MaxNameLength = 80
    )
    Assert-HUDataWriter
    $chk = Test-HUAttachment -Path $Path -Extensions $Extensions -MaxMB $MaxMB
    if (-not $chk.Ok) { throw $chk.Message }
    if (-not (Test-Path -LiteralPath $Folder)) { [void](New-Item -ItemType Directory -Path $Folder -Force) }
    $fi = New-Object IO.FileInfo $Path
    $base = Get-HUDataFileBase ([IO.Path]::GetFileNameWithoutExtension($fi.Name))
    if ($MaxNameLength -gt 0 -and $base.Length -gt $MaxNameLength) { $base = $base.Substring(0, $MaxNameLength) }
    $ext = $fi.Extension.ToLowerInvariant()
    $ts = $Now.ToString('yyyyMMdd-HHmmss', $script:HUDataInv)
    $bytes = [IO.File]::ReadAllBytes($fi.FullName)
    for ($n = 1; $n -le 1000; $n++) {
        if ($TimePrefix) { $name = $(if ($n -eq 1) { "${ts}_$base$ext" } else { "${ts}_${n}_$base$ext" }) }
        else { $name = $(if ($n -eq 1) { "$base$ext" } else { "${base}_$n$ext" }) }
        $target = Join-Path $Folder $name
        if (Test-Path -LiteralPath $target) { continue }
        try { [void](Write-HUFileAtomic -Path $target -Bytes $bytes); return [pscustomobject]@{ Name = $name; Path = $target } }
        catch { if (Test-Path -LiteralPath $target) { continue }; throw }   # ein anderer PC war gleichzeitig schneller
    }
    throw 'Kein freier Dateiname für den Anhang.'
}

<#
.SYNOPSIS
    Anhaenge eines Ordners (ohne Unterordner wie _entfernt), nach Namen sortiert. -Extensions leer = alle.
    Rueckgabe: FileInfo-Liste.
#>
function Get-HUAttachments([string]$Folder, [string[]]$Extensions = @()) {
    if (-not $Folder -or -not [IO.Directory]::Exists($Folder)) { return @() }
    $ext = @($Extensions | Where-Object { $_ } | ForEach-Object { $e = "$_".ToLowerInvariant(); if (-not $e.StartsWith('.')) { $e = '.' + $e }; $e })
    return @((New-Object IO.DirectoryInfo $Folder).GetFiles() | Where-Object { -not $_.Name.StartsWith('.') -and (-not $ext.Count -or $ext -contains $_.Extension.ToLowerInvariant()) } | Sort-Object Name)
}

<#
.SYNOPSIS
    Anzeigename ohne Zeitstempel: 20261107-143000_Bericht.pdf -> Bericht.pdf (auch mit _2_).
#>
function Get-HUAttachmentDisplayName([string]$FileName) { return ("$FileName" -replace '^\d{8}-\d{6}(_\d+)?_', '') }

<#
.SYNOPSIS
    Anhang "entfernen": nach _entfernt\ verschieben (Standard <Folder>\_entfernt, sonst -BinFolder), nicht loeschen.
    -Name: Dateiname im Ordner oder voller Pfad (muss im Ordner liegen). Gibt es den Namen dort schon, kommt ein Zeitstempel davor.
    Rueckgabe: neuer Pfad, oder '' wenn die Datei fehlte.
#>
function Remove-HUAttachment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Folder,
        [Parameter(Mandatory)][string]$Name,
        [string]$BinFolder = '',
        [datetime]$Now = (Get-Date)
    )
    $root = [IO.Path]::GetFullPath($Folder).TrimEnd('\', '/')
    if ([IO.Path]::IsPathRooted($Name)) {
        $src = [IO.Path]::GetFullPath($Name)
        if (-not $src.StartsWith($root + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw 'Der Anhang gehört nicht zu diesem Ordner.' }
    } else { $src = Join-Path $root ([IO.Path]::GetFileName($Name)) }
    if (-not (Test-Path -LiteralPath $src -PathType Leaf)) { return '' }
    $bin = $(if ($BinFolder) { $BinFolder } else { Join-Path $root '_entfernt' })
    if (-not (Test-Path -LiteralPath $bin)) { [void](New-Item -ItemType Directory -Path $bin -Force) }
    $leaf = [IO.Path]::GetFileName($src)
    $dst = Join-Path $bin $leaf
    $ts = $Now.ToString('yyyyMMdd-HHmmss', $script:HUDataInv)
    $n = 1
    while (Test-Path -LiteralPath $dst) {
        $dst = Join-Path $bin $(if ($n -eq 1) { "${ts}_$leaf" } else { "${ts}_${n}_$leaf" })
        $n++
    }
    Move-Item -LiteralPath $src -Destination $dst
    return $dst
}

# ================================================================ Gemeinsame Einstellungen ================================================================

function ConvertTo-HUDataDict($Object) {
    $h = [ordered]@{}
    if ($null -eq $Object) { return $h }
    if ($Object -is [System.Collections.IDictionary]) { foreach ($k in @($Object.Keys)) { $h["$k"] = $Object[$k] }; return $h }
    foreach ($p in $Object.PSObject.Properties) { $h[$p.Name] = $p.Value }
    return $h
}
function Test-HUDataIsRecord($Value) {
    return ($Value -is [System.Collections.IDictionary] -or $Value -is [System.Management.Automation.PSCustomObject])
}

<#
.SYNOPSIS
    Vorgaben mit gelesenen Werten mischen: Ergebnis hat genau die Schluessel der Vorgaben, Werte aus -Data, wo vorhanden
    und nicht $null. Verschachtelte Objekte werden ebenso gemischt. -KeepExtra: unbekannte Schluessel aus -Data behalten.
    Rueckgabe: PSCustomObject (Vorgaben werden nicht veraendert).
#>
function Merge-HUSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Defaults,
        $Data = $null,
        [switch]$KeepExtra
    )
    $def = ConvertTo-HUDataDict $Defaults
    $dat = ConvertTo-HUDataDict $Data
    $o = [ordered]@{}
    foreach ($k in @($def.Keys)) {
        $dv = $def[$k]
        $v = $(if ($dat.Contains($k)) { $dat[$k] } else { $null })
        if ($null -eq $v) {
            if (Test-HUDataIsRecord $dv) { $o[$k] = Merge-HUSettings -Defaults $dv } else { $o[$k] = $dv }
        } elseif ((Test-HUDataIsRecord $dv) -and (Test-HUDataIsRecord $v)) {
            $o[$k] = Merge-HUSettings -Defaults $dv -Data $v -KeepExtra:$KeepExtra
        } else { $o[$k] = $v }
    }
    if ($KeepExtra) { foreach ($k in @($dat.Keys)) { if (-not $o.Contains($k)) { $o[$k] = $dat[$k] } } }
    return [pscustomobject]$o
}

<#
.SYNOPSIS
    Gemeinsame Einstellungen eines Moduls lesen und mit -Defaults mischen. geaendert/geaendertVon kommen immer mit
    (geaendert als Text). Rueckgabe: Settings, Stamp (fuer Save-HUSettings), Error.
#>
function Read-HUSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        $Defaults = $null,
        [switch]$KeepExtra
    )
    $stamp = Get-HUFileStamp $Path
    $r = Read-HUJson $Path
    $def = ConvertTo-HUDataDict $Defaults
    if (-not $def.Contains('geaendert')) { $def.geaendert = '' }
    if (-not $def.Contains('geaendertVon')) { $def.geaendertVon = '' }
    $s = Merge-HUSettings -Defaults $def -Data $r.Data -KeepExtra:$KeepExtra
    $s.geaendert = ConvertTo-HUTimeText $s.geaendert
    $s.geaendertVon = "$($s.geaendertVon)"
    return [pscustomobject]@{ Settings = $s; Stamp = $stamp; Error = $r.Error }
}

<#
.SYNOPSIS
    Gemeinsame Einstellungen speichern: setzt geaendert/geaendertVon (auch im uebergebenen Objekt) und speichert mit
    Save-HUDataFile (Sperre, Stand-Pruefung, Kopie). Rueckgabe: Ok, Message, Stamp, Skipped.
#>
function Save-HUSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)]$Settings,
        $Stamp = $null,
        [string]$UserName = $env:USERNAME,
        [string]$PcName = $env:COMPUTERNAME,
        [string]$Changed = 'Die Einstellungen wurden inzwischen an einem anderen PC geändert. Bitte neu öffnen, es wurde nichts überschrieben.',
        [int]$Depth = 10,
        [int]$WaitMs = 10000
    )
    $now = Get-HUNowText
    foreach ($kv in @(@('geaendert', $now), @('geaendertVon', "$UserName"))) {
        if ($Settings -is [System.Collections.IDictionary]) { $Settings[$kv[0]] = $kv[1] }
        elseif ($Settings.PSObject.Properties[$kv[0]]) { $Settings.($kv[0]) = $kv[1] }
        else { $Settings | Add-Member -NotePropertyName $kv[0] -NotePropertyValue $kv[1] }
    }
    $text = ConvertTo-Json -InputObject $Settings -Depth $Depth
    return (Save-HUDataFile -Path $Path -Stamp $Stamp -UserName $UserName -PcName $PcName -Changed $Changed -WaitMs $WaitMs -Build { $text })
}

<#
.SYNOPSIS
    Darf der Benutzer? Verwalter immer; sonst steht er in -List (Gross/klein egal), als Name oder DOMAENE\Name.
    -UserName darf selbst DOMAENE\Name sein. -Domain Standard $env:USERDOMAIN.
#>
function Test-HUPermittedUser {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)][AllowNull()][AllowEmptyCollection()][object[]]$List = @(),
        [Parameter(Position = 1)][string]$UserName = $env:USERNAME,
        [switch]$IsVerwalter,
        [string]$Domain = $env:USERDOMAIN
    )
    if ($IsVerwalter) { return $true }
    $u = "$UserName".Trim()
    if (-not $u) { return $false }
    $dom = "$Domain".Trim()
    if ($u -match '^([^\\]+)\\(.+)$') { $dom = $Matches[1]; $u = $Matches[2] }
    foreach ($e in @($List)) {
        $t = "$e".Trim()
        if (-not $t) { continue }
        if ($t -ieq $u) { return $true }
        if ($dom -and $t -ieq "$dom\$u") { return $true }
    }
    return $false
}

<#
.SYNOPSIS
    Anzeigename des angemeldeten Windows-Benutzers aus der Domaene (UserPrincipal.Current.DisplayName).
    Ohne Domaene, bei Fehler oder fuer einen anderen -UserName: der Benutzername. Wird je Sitzung einmal abgefragt.
#>
function Get-HUUserDisplayName([string]$UserName = $env:USERNAME) {
    if ("$UserName" -ine "$env:USERNAME") { return "$UserName" }
    if ($null -ne $script:HUDataDisplayName) { return $(if ($script:HUDataDisplayName) { $script:HUDataDisplayName } else { "$UserName" }) }
    $n = ''
    try {
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement -ErrorAction Stop
        $n = "$([System.DirectoryServices.AccountManagement.UserPrincipal]::Current.DisplayName)".Trim()
    } catch { $n = '' }
    $script:HUDataDisplayName = $n
    if ($n) { return $n }
    return "$UserName"
}

Export-ModuleMember -Function Get-HUNowText, ConvertTo-HUTimeText, Get-HUFileStamp, ConvertFrom-HUJsonText, ConvertTo-HUJsonText,
    Copy-HUDataObject, Read-HUJson, Get-HUProp,
    Wait-HUDataFileLock, Write-HUDataFile, Save-HUDataFile,
    New-HUDataId, Test-HUDataId, New-HUShortId, Test-HUShortId, Get-HUNextNumber,
    Get-HUDataRecords,
    ConvertTo-HUDecimal, Test-HUNumberText, Test-HUAmbiguousNumber, Get-HURound, Format-HUEuro, Format-HUNumber, Format-HUKm,
    ConvertFrom-HUDateText, ConvertTo-HUDateText, ConvertFrom-HUClockText, Format-HUDate, Format-HUTime,
    New-HUHistoryEntry, Add-HUHistoryEntry, ConvertTo-HUHistoryEntry, Compare-HURecordFields, Get-HUFieldChangeText,
    Test-HUAttachment, Add-HUAttachment, Get-HUAttachments, Get-HUAttachmentDisplayName, Remove-HUAttachment,
    Merge-HUSettings, Read-HUSettings, Save-HUSettings, Test-HUPermittedUser, Get-HUUserDisplayName
