#Requires -Version 5.1
<#
.SYNOPSIS
    HU.Exchange - Datei-Austausch mit anderen Programmen: Tabellen (CSV) nach Profil und Uebergabeordner je Modul.
.DESCRIPTION
    Tabellen:
      Ein Profil beschreibt das Format, das ein anderes Programm erwartet: Trennzeichen, Kodierung, Dezimalzeichen,
      Datumsformat und die Spalten (Name in der Datei, Eigenschaft im Programm, Typ text | number | date).
      Vorgaben fuer Oesterreich: Trennzeichen ';', Dezimalkomma, Datum dd.MM.yyyy, UTF-8 mit BOM.
      Kodierungen: utf8bom, utf8, windows-1252 (viele aeltere Programme brauchen 1252).
      Profile liegen als JSON in Data\<Modul-Id>\Profile\<Name>.json. Ein neues Zielformat ist nur eine Datei, kein Code.
      Die Kodierung wird immer ausdruecklich gesetzt (nie die Vorgabe von Export-Csv, die sich zwischen 5.1 und 7 unterscheidet).
    Uebergabeordner:
      Jedes Modul, das etwas weitergibt, schreibt in seinen Uebergabeordner (Standard Data\<Modul-Id>\Ausgang,
      einstellbar in Data\<Modul-Id>\uebergabe.json, auch eine Freigabe).
      Dateien entstehen unter einem Temp-Namen und werden erst fertig umbenannt (ein abholendes Programm liest nie
      eine halbe Datei). Dateinamen tragen einen Zeitstempel, vorhandene Dateien werden nie ueberschrieben.
.NOTES
    Zielmaschine: der PC, auf dem die App laeuft. PowerShell 5.1 und 7.
#>

Set-StrictMode -Version 1.0

$script:HUExEncodings = @('utf8bom', 'utf8', 'windows-1252')
$script:HUExTypes = @('text', 'number', 'date')
# Datumsformate fuer Format-HUExValue: oesterreichisch (d.M.yyyy, mit Uhrzeit) und ISO (mit Uhrzeit, Bruchteil, Zone)
$script:HUExDateFormats = [string[]]@(
    'd.M.yyyy', 'd.M.yyyy H:mm', 'd.M.yyyy H:mm:ss',
    'yyyy-MM-dd', 'yyyy-MM-dd HH:mm', 'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-ddTHH:mm', 'yyyy-MM-ddTHH:mm:ss',
    'yyyy-MM-ddTHH:mm:ss.FFFFFFF', 'yyyy-MM-ddTHH:mm:ssK', 'yyyy-MM-ddTHH:mm:ss.FFFFFFFK')

# ---------------------------------------------------------------- Kodierung und Profil ----------------------------------------------------------------

function Get-HUTableEncoding([string]$Name) {
    switch ("$Name".ToLowerInvariant()) {
        'utf8bom' { return (New-Object System.Text.UTF8Encoding $true) }
        'utf8' { return (New-Object System.Text.UTF8Encoding $false) }
        'windows-1252' { return [System.Text.Encoding]::GetEncoding(1252) }
        default { throw "Unbekannte Kodierung '$Name' (erlaubt: $($script:HUExEncodings -join ', '))." }
    }
}

<#
.SYNOPSIS
    Neues Profil mit den Vorgaben fuer Oesterreich. -Columns: @{ Name; Source; Type } (Source leer = wie Name, Type leer = text).
#>
function New-HUTableProfile {
    [CmdletBinding()]
    param(
        [string]$Name = 'Standard',
        [string]$Delimiter = ';',
        [string]$Encoding = 'utf8bom',
        [string]$DecimalSeparator = ',',
        [string]$DateFormat = 'dd.MM.yyyy',
        [object[]]$Columns = @()
    )
    return (ConvertTo-HUTableProfile ([pscustomobject]@{ Name = $Name; Delimiter = $Delimiter; Encoding = $Encoding; DecimalSeparator = $DecimalSeparator; DateFormat = $DateFormat; Columns = @($Columns) }))
}

<#
.SYNOPSIS
    Profil pruefen und vervollstaendigen (fehlende Werte = Vorgaben). Wirft bei ungueltigen Angaben.
#>
function ConvertTo-HUTableProfile($InputObject) {
    $g = { param($n, $d) if ($InputObject -and $InputObject.PSObject.Properties[$n] -and $null -ne $InputObject.$n -and "$($InputObject.$n)" -ne '') { $InputObject.$n } else { $d } }
    $delim = "$(& $g 'Delimiter' ';')"
    if ($delim -eq '\t' -or $delim -ieq 'tab') { $delim = "`t" }
    if ($delim.Length -ne 1 -or $delim -match '["\r\n]') { throw "Trennzeichen '$delim' ungueltig (genau ein Zeichen, kein Anfuehrungszeichen)." }
    $enc = "$(& $g 'Encoding' 'utf8bom')".ToLowerInvariant()
    if ($script:HUExEncodings -notcontains $enc) { throw "Unbekannte Kodierung '$enc' (erlaubt: $($script:HUExEncodings -join ', '))." }
    $dec = "$(& $g 'DecimalSeparator' ',')"
    if ($dec -notin ',', '.') { throw "Dezimalzeichen '$dec' ungueltig (erlaubt: Komma oder Punkt)." }
    if ($dec -eq $delim) { throw 'Dezimalzeichen und Trennzeichen duerfen nicht gleich sein.' }
    $df = "$(& $g 'DateFormat' 'dd.MM.yyyy')"
    try { [void]([datetime]'2026-01-31').ToString($df, [Globalization.CultureInfo]::InvariantCulture) } catch { throw "Datumsformat '$df' ungueltig." }
    $cols = New-Object System.Collections.Generic.List[object]
    $raw = @(); if ($InputObject -and $InputObject.PSObject.Properties['Columns'] -and $InputObject.Columns) { $raw = @($InputObject.Columns) }
    foreach ($c in $raw) {
        if ($null -eq $c) { continue }
        $cn = if ($c -is [string]) { $c } elseif ($c -is [hashtable]) { "$($c['Name'])" } else { "$($c.Name)" }
        $cs = if ($c -is [string]) { '' } elseif ($c -is [hashtable]) { "$($c['Source'])" } else { "$(if ($c.PSObject.Properties['Source']) { $c.Source })" }
        $ct = if ($c -is [string]) { '' } elseif ($c -is [hashtable]) { "$($c['Type'])" } else { "$(if ($c.PSObject.Properties['Type']) { $c.Type })" }
        if (-not $cn.Trim()) { throw 'Spalte ohne Namen im Profil.' }
        if (-not $cs) { $cs = $cn }
        $ct = $(if ($ct) { $ct.ToLowerInvariant() } else { 'text' })
        if ($script:HUExTypes -notcontains $ct) { throw "Spalte '$cn': Typ '$ct' unbekannt (erlaubt: $($script:HUExTypes -join ', '))." }
        if (@($cols | Where-Object { $_.Name -eq $cn }).Count) { throw "Spalte '$cn' doppelt im Profil." }
        $cols.Add([pscustomobject][ordered]@{ Name = $cn; Source = $cs; Type = $ct })
    }
    return [pscustomobject][ordered]@{
        Name = "$(& $g 'Name' 'Standard')"; Delimiter = $delim; Encoding = $enc; DecimalSeparator = $dec; DateFormat = $df; Columns = @($cols.ToArray())
    }
}

function Read-HUTableProfile([Parameter(Mandatory)][string]$Path) {
    $t = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
    $o = $t | ConvertFrom-Json
    if (-not ($o.PSObject.Properties['Name'] -and "$($o.Name)")) { $o | Add-Member -NotePropertyName Name -NotePropertyValue ([IO.Path]::GetFileNameWithoutExtension($Path)) -Force }
    return (ConvertTo-HUTableProfile $o)
}

function Save-HUTableProfile {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$TableProfile, [Parameter(Mandatory)][string]$Path)
    $p = ConvertTo-HUTableProfile $TableProfile
    $o = [pscustomobject][ordered]@{ Name = $p.Name; Delimiter = $(if ($p.Delimiter -eq "`t") { '\t' } else { $p.Delimiter }); Encoding = $p.Encoding; DecimalSeparator = $p.DecimalSeparator; DateFormat = $p.DateFormat; Columns = @($p.Columns) }
    $json = $o | ConvertTo-Json -Depth 5
    Write-HUFileAtomic -Path $Path -Text $json -Encoding (New-Object System.Text.UTF8Encoding $true) -Overwrite
}

# Ordner der Profile eines Moduls: Data\<Id>\Profile (wird angelegt)
function Get-HUTableProfileFolder([Parameter(Mandatory)][string]$DataRoot, [Parameter(Mandatory)][string]$ModuleId) {
    $p = Join-Path (Get-HUExModuleDir $DataRoot $ModuleId) 'Profile'
    if (-not (Test-Path -LiteralPath $p)) { [void](New-Item -ItemType Directory -Path $p -Force) }
    return $p
}

# alle Profile eines Ordners (fehlerhafte werden mit Grund gemeldet, nicht verschluckt)
function Get-HUTableProfiles([Parameter(Mandatory)][string]$Folder) {
    $out = New-Object System.Collections.Generic.List[object]
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { return @() }
    foreach ($f in @(Get-ChildItem -LiteralPath $Folder -Filter '*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
        try { $out.Add([pscustomobject]@{ Path = $f.FullName; Profile = (Read-HUTableProfile $f.FullName); Error = '' }) }
        catch { $out.Add([pscustomobject]@{ Path = $f.FullName; Profile = $null; Error = "$($_.Exception.Message)" }) }
    }
    return @($out.ToArray())
}

# ---------------------------------------------------------------- Werte umwandeln ----------------------------------------------------------------

function Get-HUExNumberFormat([string]$Decimal) {
    $nf = [Globalization.NumberFormatInfo]::InvariantInfo.Clone()
    $nf.NumberDecimalSeparator = $Decimal
    $nf.NumberGroupSeparator = ''
    return $nf
}

# Zahl aus Text (KERN-B1): Schreibweise wie in Oesterreich zuerst.
#   - Mit Komma: Komma = Dezimalzeichen, Punkt = Tausender, nur in Dreiergruppen (1.234,56 / 12,50 / -1,5).
#   - Ohne Komma, hoechstens ein Punkt: Punkt = Dezimalzeichen (1.5 / 1234.56 aus Programmdaten, wie ConvertTo-HUDecimal).
#   - Ohne Komma, mehrere Punkte in Dreiergruppen: Tausender (1.234.567).
#   - Alles andere (z. B. 1,234.56 oder 1.2,5) ist keine Zahl.
function ConvertFrom-HUExNumberText([string]$Text) {
    $t = "$Text".Trim()
    $d = [decimal]0
    if ($t -match '^[+-]?\d+(\.\d+)?$') {
        if ([decimal]::TryParse($t, [Globalization.NumberStyles]::AllowLeadingSign -bor [Globalization.NumberStyles]::AllowDecimalPoint, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $d }
    }
    elseif ($t -match '^[+-]?(\d+|\d{1,3}(\.\d{3})+)(,\d+)?$') {
        $inv = $t.Replace('.', '').Replace(',', '.')
        if ([decimal]::TryParse($inv, [Globalization.NumberStyles]::AllowLeadingSign -bor [Globalization.NumberStyles]::AllowDecimalPoint, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $d }
    }
    throw "'$Text' ist keine Zahl."
}

# Datum aus Text (KERN-B1): nur feste Formate aus $script:HUExDateFormats, kein freies TryParse.
function ConvertFrom-HUExDateText([string]$Text) {
    $t = "$Text".Trim()
    $x = [datetime]::MinValue
    # [string[]] ausdruecklich: ein Array-Literal bindet sonst an die Ueberladung mit einem einzelnen Format
    if ([datetime]::TryParseExact($t, $script:HUExDateFormats, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$x)) { return $x }
    throw "'$Text' ist kein Datum."
}

function Format-HUExValue($Value, [string]$Type, $TableProfile) {
    if ($null -eq $Value) { return '' }
    switch ($Type) {
        'number' {
            if ("$Value" -eq '') { return '' }
            $nf = Get-HUExNumberFormat $TableProfile.DecimalSeparator
            $d = [decimal]0
            if ($Value -is [string]) { $d = ConvertFrom-HUExNumberText $Value }
            else { $d = [decimal]$Value }
            return $d.ToString('0.############################', $nf)
        }
        'date' {
            if ("$Value" -eq '') { return '' }
            $dt = if ($Value -is [datetime]) { $Value } else { ConvertFrom-HUExDateText "$Value" }
            return $dt.ToString($TableProfile.DateFormat, [Globalization.CultureInfo]::InvariantCulture)
        }
        default { return "$Value" }
    }
}

function ConvertFrom-HUExValue([string]$Text, [string]$Type, $TableProfile) {
    switch ($Type) {
        'number' {
            $t = "$Text".Trim(); if (-not $t) { return $null }
            $d = [decimal]0
            if (-not [decimal]::TryParse($t, [Globalization.NumberStyles]::Number, (Get-HUExNumberFormat $TableProfile.DecimalSeparator), [ref]$d)) { throw "'$Text' ist keine Zahl (Dezimalzeichen '$($TableProfile.DecimalSeparator)')." }
            return $d
        }
        'date' {
            $t = "$Text".Trim(); if (-not $t) { return $null }
            $x = [datetime]::MinValue
            if (-not [datetime]::TryParseExact($t, $TableProfile.DateFormat, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$x)) { throw "'$Text' passt nicht zum Datumsformat '$($TableProfile.DateFormat)'." }
            return $x
        }
        default { return $Text }
    }
}

function ConvertTo-HUCsvField([string]$Text, [string]$Delimiter) {
    if ($Text.IndexOfAny(([char[]]@($Delimiter[0], '"', "`r", "`n"))) -ge 0 -or $Text -ne $Text.Trim()) { return '"' + $Text.Replace('"', '""') + '"' }
    return $Text
}

# ---------------------------------------------------------------- Lesen und Schreiben ----------------------------------------------------------------

<#
.SYNOPSIS
    Datei unter einem Temp-Namen im Zielordner schreiben und dann umbenennen. Ohne -Overwrite wirft es, wenn das Ziel existiert.
#>
function Write-HUFileAtomic {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [string]$Text = '', [byte[]]$Bytes = $null, [System.Text.Encoding]$Encoding = $null, [switch]$Overwrite)
    $full = [IO.Path]::GetFullPath($Path)
    $dir = Split-Path -Parent $full
    if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
    if (-not $Overwrite -and (Test-Path -LiteralPath $full)) { throw "Datei gibt es schon: $full" }
    if ($null -eq $Bytes) {
        if (-not $Encoding) { $Encoding = New-Object System.Text.UTF8Encoding $true }
        $pre = $Encoding.GetPreamble(); $body = $Encoding.GetBytes($Text)
        $Bytes = New-Object byte[] ($pre.Length + $body.Length)
        [Array]::Copy($pre, 0, $Bytes, 0, $pre.Length); [Array]::Copy($body, 0, $Bytes, $pre.Length, $body.Length)
    }
    $tmp = Join-Path $dir ('.~' + [IO.Path]::GetFileName($full) + '.' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.tmp')
    try {
        $fs = New-Object IO.FileStream($tmp, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        try { $fs.Write($Bytes, 0, $Bytes.Length); $fs.Flush($true) } finally { $fs.Dispose() }
        if ($Overwrite -and (Test-Path -LiteralPath $full)) {
            # Replace braucht WRITE_DAC (Microsoft Learn, ReplaceFileW). Manche Netzfreigaben verweigern das trotz NTFS-Vollzugriff:
            # dann die Datei normal ueberschreiben (Inhalt kommt aus der fertig geschriebenen Temp-Datei).
            try { [IO.File]::Replace($tmp, $full, [NullString]::Value) }   # $null kaeme als '' an
            catch {
                $ex = $_.Exception
                if ($ex -is [System.Management.Automation.MethodInvocationException] -and $ex.InnerException) { $ex = $ex.InnerException }
                if ($ex -isnot [System.UnauthorizedAccessException]) { throw }
                # KERN-B10: bewusst so gelassen. Copy ist nicht atomar (ein Abbruch mitten im Kopieren kann eine halbe Datei
                # hinterlassen), aber ohne WRITE_DAC gibt es keinen atomaren Weg; die vollstaendige Temp-Datei ist die Quelle.
                [IO.File]::Copy($tmp, $full, $true)
            }
        }
        else { [IO.File]::Move($tmp, $full) }
    } finally { if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue } }
    return $full
}

<#
.SYNOPSIS
    Zeilen nach Profil als CSV schreiben (atomar). Ohne Spalten im Profil: alle Eigenschaften der ersten Zeile,
    Zahlen und Datumswerte werden erkannt. Rueckgabe: Pfad.
#>
function Export-HUTable {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows, [Parameter(Mandatory)]$TableProfile, [Parameter(Mandatory)][string]$Path, [switch]$NoHeader, [switch]$Overwrite)
    $p = ConvertTo-HUTableProfile $TableProfile
    $cols = @($p.Columns)
    if (-not $cols.Count) {
        $first = @($Rows | Where-Object { $null -ne $_ })[0]
        if ($first) {
            $props = if ($first -is [hashtable]) { @($first.Keys) } else { @($first.PSObject.Properties | ForEach-Object Name) }
            $cols = @(foreach ($n in $props) {
                    $v = if ($first -is [hashtable]) { $first[$n] } else { $first.$n }
                    $t = if ($v -is [datetime]) { 'date' } elseif ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal] -or $v -is [single]) { 'number' } else { 'text' }
                    [pscustomobject]@{ Name = "$n"; Source = "$n"; Type = $t }
                })
        }
    }
    $sb = New-Object System.Text.StringBuilder
    $d = $p.Delimiter
    if (-not $NoHeader -and $cols.Count) { [void]$sb.Append((@($cols | ForEach-Object { ConvertTo-HUCsvField $_.Name $d }) -join $d)).Append("`r`n") }
    $i = 0
    foreach ($r in $Rows) {
        $i++
        if ($null -eq $r) { continue }
        $fields = foreach ($c in $cols) {
            $v = if ($r -is [hashtable]) { $r[$c.Source] } elseif ($r.PSObject.Properties[$c.Source]) { $r.($c.Source) } else { $null }
            try { ConvertTo-HUCsvField (Format-HUExValue $v $c.Type $p) $d } catch { throw "Zeile $($i), Spalte '$($c.Name)': $($_.Exception.Message)" }
        }
        [void]$sb.Append((@($fields) -join $d)).Append("`r`n")
    }
    return (Write-HUFileAtomic -Path $Path -Text $sb.ToString() -Encoding (Get-HUTableEncoding $p.Encoding) -Overwrite:$Overwrite)
}

<#
.SYNOPSIS
    CSV nach Profil lesen. Mit Spalten im Profil: Rueckgabe-Eigenschaften = Source, Werte nach Typ umgewandelt
    (fehlende Spalten in der Datei sind ein Fehler, weitere Spalten werden ignoriert). Ohne Spalten: alles als Text.
#>
function Import-HUTable {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)]$TableProfile)
    $p = ConvertTo-HUTableProfile $TableProfile
    $enc = Get-HUTableEncoding $p.Encoding
    $bytes = [IO.File]::ReadAllBytes($Path)
    $start = 0
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { $start = 3; $enc = New-Object System.Text.UTF8Encoding $false }
    $text = $enc.GetString($bytes, $start, $bytes.Length - $start)
    if (-not $text.Trim()) { return @() }
    $rows = @($text | ConvertFrom-Csv -Delimiter $p.Delimiter)
    $cols = @($p.Columns)
    if (-not $cols.Count) { return $rows }
    $have = if ($rows.Count) { @($rows[0].PSObject.Properties | ForEach-Object Name) } else { @() }
    $miss = @($cols | Where-Object { $have -notcontains $_.Name } | ForEach-Object Name)
    if ($rows.Count -and $miss.Count) { throw "Spalten fehlen in der Datei: $($miss -join ', ')" }
    $out = New-Object System.Collections.Generic.List[object]
    $i = 1
    foreach ($r in $rows) {
        $i++
        $o = [ordered]@{}
        foreach ($c in $cols) {
            try { $o[$c.Source] = ConvertFrom-HUExValue "$($r.($c.Name))" $c.Type $p } catch { throw "Zeile $($i), Spalte '$($c.Name)': $($_.Exception.Message)" }
        }
        $out.Add([pscustomobject]$o)
    }
    return @($out.ToArray())
}

# ---------------------------------------------------------------- Uebergabeordner ----------------------------------------------------------------

function Get-HUExModuleDir([string]$DataRoot, [string]$ModuleId) {
    if ($ModuleId -notmatch '^[A-Za-z0-9][A-Za-z0-9_\-]{0,63}$') { throw "Ungueltige Modul-Id: '$ModuleId'" }
    if (@('Benutzer', 'Laufend', 'Docs', 'Postfach', 'Wiederherstellung') -contains $ModuleId) { throw "Modul-Id '$ModuleId' ist reserviert" }
    $p = Join-Path $DataRoot $ModuleId
    if (-not (Test-Path -LiteralPath $p)) { [void](New-Item -ItemType Directory -Path $p -Force) }
    return $p
}

<#
.SYNOPSIS
    Uebergabeordner eines Moduls: eingestellter Ordner (Data\<Id>\uebergabe.json, Folder) oder Data\<Id>\Ausgang. Wird angelegt.
#>
function Get-HUHandoverFolder([Parameter(Mandatory)][string]$DataRoot, [Parameter(Mandatory)][string]$ModuleId) {
    $mod = Get-HUExModuleDir $DataRoot $ModuleId
    $f = ''
    $cfg = Join-Path $mod 'uebergabe.json'
    if (Test-Path -LiteralPath $cfg) { try { $f = "$(([IO.File]::ReadAllText($cfg, [Text.Encoding]::UTF8) | ConvertFrom-Json).Folder)".Trim() } catch { $f = '' } }
    $dir = if ($f) { [Environment]::ExpandEnvironmentVariables($f) } else { Join-Path $mod 'Ausgang' }
    if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
    return $dir
}

# Uebergabeordner einstellen ('' = Standard Data\<Id>\Ausgang)
function Set-HUHandoverFolder([Parameter(Mandatory)][string]$DataRoot, [Parameter(Mandatory)][string]$ModuleId, [string]$Folder = '') {
    $mod = Get-HUExModuleDir $DataRoot $ModuleId
    $json = ([pscustomobject]@{ Folder = "$Folder".Trim() } | ConvertTo-Json)
    [void](Write-HUFileAtomic -Path (Join-Path $mod 'uebergabe.json') -Text $json -Overwrite)
}

<#
.SYNOPSIS
    Freier Dateiname im Ordner: <Name>_<JJJJ-MM-TT_HHmmss>.<Endung>, bei Gleichstand _2, _3 ... (nie ueberschreiben).
#>
function New-HUHandoverPath([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string]$BaseName, [string]$Extension = 'csv', [datetime]$Now = (Get-Date)) {
    $b = ($BaseName -replace '[^\p{L}\p{Nd}\.\-_]', '_').Trim('.', ' ', '_'); if (-not $b) { $b = 'Datei' }
    $e = $Extension.TrimStart('.')
    $stem = "{0}_{1}" -f $b, $Now.ToString('yyyy-MM-dd_HHmmss')
    $p = Join-Path $Folder "$stem.$e"
    $n = 1
    while (Test-Path -LiteralPath $p) { $n++; $p = Join-Path $Folder "${stem}_$n.$e" }
    return $p
}

<#
.SYNOPSIS
    Tabelle nach Profil in den Uebergabeordner des Moduls schreiben (Temp-Name, dann umbenennen; Zeitstempel im Namen).
    Rueckgabe: Pfad der fertigen Datei.
#>
function Export-HUHandoverTable {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows, [Parameter(Mandatory)]$TableProfile, [Parameter(Mandatory)][string]$DataRoot,
        [Parameter(Mandatory)][string]$ModuleId, [Parameter(Mandatory)][string]$BaseName, [datetime]$Now = (Get-Date))
    $dir = Get-HUHandoverFolder $DataRoot $ModuleId
    for ($try = 0; $try -lt 5; $try++) {
        $path = New-HUHandoverPath -Folder $dir -BaseName $BaseName -Extension 'csv' -Now $Now
        try { return (Export-HUTable -Rows $Rows -TableProfile $TableProfile -Path $path) }
        catch { if ($try -lt 4 -and (Test-Path -LiteralPath $path)) { continue }; throw }
    }
}

Export-ModuleMember -Function Get-HUTableEncoding, New-HUTableProfile, ConvertTo-HUTableProfile, Read-HUTableProfile, Save-HUTableProfile,
    Get-HUTableProfileFolder, Get-HUTableProfiles, Write-HUFileAtomic, Export-HUTable, Import-HUTable,
    Get-HUHandoverFolder, Set-HUHandoverFolder, New-HUHandoverPath, Export-HUHandoverTable
