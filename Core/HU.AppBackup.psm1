#Requires -Version 5.1
<#
.SYNOPSIS
    HU.AppBackup - verschluesselte App-Sicherung (Format .husicherung).
.DESCRIPTION
    Zielmaschine: der PC, auf dem die App laeuft (PowerShell 5.1 / .NET 4.8 und PowerShell 7).

    Format (Version 1):
        'HUSICH1' (7 Byte ASCII) | Kopflaenge (4 Byte, little endian) | Kopf (UTF-8 JSON) | Chiffrat | HMAC (32 Byte)
    Kopf: Format, App, Created (UTC), Kdf (PBKDF2-SHA256), Iter, Salt (Base64, 32 Byte), IV (Base64, 16 Byte), Items (nur Namen).
    Schluessel: PBKDF2-HMAC-SHA256(Kennwort, Salt, Iter) -> 64 Byte = 32 AES-256 + 32 HMAC-SHA256.
    Verschluesselung: AES-256-CBC/PKCS7, Encrypt-then-MAC. Der HMAC deckt Magic, Kopflaenge, Kopf und Chiffrat
    und wird in konstanter Zeit VOR dem Entschluesseln geprueft (AES-GCM gibt es in .NET Framework 4.8 nicht).
    Falsches Kennwort und veraenderte Datei ergeben dieselbe Meldung.
#>

Set-StrictMode -Version 1

$script:HUAbMagic   = [Text.Encoding]::ASCII.GetBytes('HUSICH1')
$script:HUAbIter    = 600000
$script:HUAbMinPw   = 10
$script:HUAbMaxHead = 64KB
$script:HUAbIterMin  = 600000
$script:HUAbIterMax  = 5000000
$script:HUAbFailMsg = "Kennwort falsch oder Datei veraendert/beschaedigt - nichts entschluesselt.`n`nBitte pruefen: Gross-/Kleinschreibung und Tastaturbelegung (z/y, Feststelltaste), die richtige Datei (PC-Name und Datum stehen im Dateinamen), Datei vollstaendig kopiert (Groesse wie am Quell-PC). Ohne das richtige Kennwort ist die Sicherung nicht zu oeffnen."

function Get-HUAbPlain([object]$Password) {
    if ($Password -is [securestring]) {
        $b = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
        try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($b) } finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($b) }
    }
    return [string]$Password
}

function Test-HUAbPassword {
    <# Gibt $null zurueck, wenn das Kennwort taugt, sonst den Grund. #>
    param([object]$Password, [object]$Repeat)
    $p = Get-HUAbPlain $Password
    if ($p.Length -lt $script:HUAbMinPw) { return "Kennwort zu kurz (mindestens $($script:HUAbMinPw) Zeichen)." }
    if ($PSBoundParameters.ContainsKey('Repeat') -and $p -cne (Get-HUAbPlain $Repeat)) { return 'Die beiden Kennwoerter stimmen nicht ueberein.' }
    return $null
}

# 072 G12: PBKDF2-SHA256. PS 7 (.NET 6+): statisches Pbkdf2 (der Konstruktor ist dort veraltet, SYSLIB0060);
# PS 5.1: Konstruktor mit HashAlgorithmName - braucht .NET Framework 4.7.2 oder neuer, sonst klare Meldung.
function Get-HUAbPbkdf2([byte[]]$Pw, [byte[]]$Salt, [int]$Iterations) {
    $t = [Security.Cryptography.Rfc2898DeriveBytes]
    if ($t.GetMethod('Pbkdf2', [type[]]@([byte[]], [byte[]], [int], [Security.Cryptography.HashAlgorithmName], [int]))) {
        return [Security.Cryptography.Rfc2898DeriveBytes]::Pbkdf2($Pw, $Salt, $Iterations, [Security.Cryptography.HashAlgorithmName]::SHA256, 32)
    }
    if (-not $t.GetConstructor([type[]]@([byte[]], [byte[]], [int], [Security.Cryptography.HashAlgorithmName]))) {
        throw 'App-Sicherung braucht .NET Framework 4.7.2 oder neuer (Windows 10 ab 1803 hat es) - bitte .NET aktualisieren.'
    }
    $kdf = New-Object Security.Cryptography.Rfc2898DeriveBytes -ArgumentList $Pw, $Salt, $Iterations, ([Security.Cryptography.HashAlgorithmName]::SHA256)
    try { return $kdf.GetBytes(32) } finally { $kdf.Dispose() }
}

# 072 B4: Temp-ZIP (Klartext, DeleteOnClose) kann nach einem Absturz liegen bleiben - alte Reste beim naechsten Lauf entfernen
function Clear-HUAbTempZips {
    try {
        $lim = (Get-Date).AddHours(-1)
        foreach ($f in @(Get-ChildItem -LiteralPath ([IO.Path]::GetTempPath()) -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^HU-(Sicherung|Wiederherstellung|Quelle)-[0-9a-f]{32}\.(zip|husicherung)$' -and $_.LastWriteTime -lt $lim })) {
            try { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop } catch { }
        }
    } catch { }
}

# 072 G13: liegt die Sicherung nicht auf einer lokalen Platte (Freigabe, USB), erst lokal kopieren - Pruefung (MAC)
# und Entschluesseln lesen dann sicher dieselben Bytes. Rueckgabe: lesbarer Stream (Temp-Kopie verschwindet beim Schliessen).
function Open-HUAbSourceStream([string]$Path, [switch]$ForceCopy) {
    $local = -not $ForceCopy
    try {
        if (-not $local) { }
        elseif ($Path -match '^(\\\\|//)') { $local = $false }
        elseif ([Environment]::OSVersion.Platform -eq [PlatformID]::Win32NT) { $local = ([IO.DriveInfo]::new([IO.Path]::GetPathRoot([IO.Path]::GetFullPath($Path))).DriveType -eq [IO.DriveType]::Fixed) }
    } catch { $local = $false }
    $src = New-Object IO.FileStream($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    if ($local) { return $src }
    try {
        $tmp = New-Object IO.FileStream((Join-Path ([IO.Path]::GetTempPath()) ("HU-Quelle-" + [guid]::NewGuid().ToString('N') + '.husicherung')), [IO.FileMode]::CreateNew, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None, 65536, [IO.FileOptions]::DeleteOnClose)
        $src.CopyTo($tmp, 1MB); [void]$tmp.Seek(0, 'Begin')
        return $tmp
    } finally { $src.Dispose() }
}

function Get-HUAbKeys {
    # 070 B2: PBKDF2 liefert nur 32 Byte (ein Block - sonst rechnet die App doppelt, ein Angreifer aber nur einmal);
    # AES- und MAC-Schluessel werden daraus per HMAC-SHA256 mit festen Bezeichnern abgeleitet.
    param([object]$Password, [byte[]]$Salt, [int]$Iterations = $script:HUAbIter)
    $pw = [Text.Encoding]::UTF8.GetBytes((Get-HUAbPlain $Password))
    try {
        $master = Get-HUAbPbkdf2 $pw $Salt $Iterations
    } finally { [Array]::Clear($pw, 0, $pw.Length) }
    $hm = New-Object Security.Cryptography.HMACSHA256 -ArgumentList (, $master)
    try {
        $enc = $hm.ComputeHash([Text.Encoding]::ASCII.GetBytes('HU-Sicherung enc'))
        $mac = $hm.ComputeHash([Text.Encoding]::ASCII.GetBytes('HU-Sicherung mac'))
    } finally { $hm.Dispose(); [Array]::Clear($master, 0, 32) }
    return @{ Enc = $enc; Mac = $mac }
}

function Get-HUAbRandom([int]$Count) {
    $b = New-Object byte[] $Count
    $r = [Security.Cryptography.RandomNumberGenerator]::Create()
    try { $r.GetBytes($b) } finally { $r.Dispose() }
    return , $b
}

function Test-HUAbEqual([byte[]]$A, [byte[]]$B) {
    if ($null -eq $A -or $null -eq $B -or $A.Length -ne $B.Length) { return $false }
    $d = 0
    for ($i = 0; $i -lt $A.Length; $i++) { $d = $d -bor ($A[$i] -bxor $B[$i]) }
    return ($d -eq 0)
}

function New-HUAbHeadBytes([int]$Iterations, [byte[]]$Salt, [byte[]]$IV, [hashtable]$Header) {
    $h = [ordered]@{ Format = 1; Created = (Get-Date).ToUniversalTime().ToString('o'); Kdf = 'PBKDF2-SHA256'; Iter = $Iterations; Salt = [Convert]::ToBase64String($Salt); IV = [Convert]::ToBase64String($IV) }
    foreach ($k in $Header.Keys) { if (-not $h.Contains($k)) { $h[$k] = $Header[$k] } }
    $head = [Text.Encoding]::UTF8.GetBytes(($h | ConvertTo-Json -Depth 6 -Compress))
    $len = [BitConverter]::GetBytes([int]$head.Length); if (-not [BitConverter]::IsLittleEndian) { [Array]::Reverse($len) }
    $all = New-Object byte[] ($script:HUAbMagic.Length + 4 + $head.Length)
    [Array]::Copy($script:HUAbMagic, 0, $all, 0, $script:HUAbMagic.Length)
    [Array]::Copy($len, 0, $all, $script:HUAbMagic.Length, 4)
    [Array]::Copy($head, 0, $all, $script:HUAbMagic.Length + 4, $head.Length)
    return , $all
}

function New-HUAbAes([byte[]]$Key, [byte[]]$IV) {
    $aes = [Security.Cryptography.Aes]::Create()
    $aes.KeySize = 256; $aes.Mode = [Security.Cryptography.CipherMode]::CBC; $aes.Padding = [Security.Cryptography.PaddingMode]::PKCS7
    $aes.Key = $Key; $aes.IV = $IV
    return $aes
}

function Protect-HUAppBackupStream {
    <# Verschluesselt $InStream (ZIP) blockweise nach $OutStream (Format .husicherung). Speicherbedarf unabhaengig von der Groesse. #>
    param([Parameter(Mandatory)][IO.Stream]$InStream, [Parameter(Mandatory)][IO.Stream]$OutStream, [Parameter(Mandatory)][object]$Password,
        [hashtable]$Header = @{}, [int]$Iterations = $script:HUAbIter)
    $why = Test-HUAbPassword $Password
    if ($why) { throw $why }
    $salt = Get-HUAbRandom 32; $iv = Get-HUAbRandom 16
    $keys = Get-HUAbKeys -Password $Password -Salt $salt -Iterations $Iterations
    $aes = $null; $t = $null; $hm = $null
    try {
        $head = New-HUAbHeadBytes $Iterations $salt $iv $Header
        $hm = New-Object Security.Cryptography.HMACSHA256 -ArgumentList (, $keys.Mac)
        $OutStream.Write($head, 0, $head.Length); [void]$hm.TransformBlock($head, 0, $head.Length, $null, 0)
        $aes = New-HUAbAes $keys.Enc $iv; $t = $aes.CreateEncryptor()
        $buf = New-Object byte[] 65536; $ob = New-Object byte[] (65536 + 32)
        while ($true) {
            # immer ganze Bloecke lesen, den Rest in TransformFinalBlock
            $n = 0
            while ($n -lt $buf.Length) { $r = $InStream.Read($buf, $n, $buf.Length - $n); if ($r -le 0) { break }; $n += $r }
            if ($n -lt $buf.Length) {
                $last = $t.TransformFinalBlock($buf, 0, $n)
                $OutStream.Write($last, 0, $last.Length); [void]$hm.TransformBlock($last, 0, $last.Length, $null, 0)
                break
            }
            $w = $t.TransformBlock($buf, 0, $n, $ob, 0)
            $OutStream.Write($ob, 0, $w); [void]$hm.TransformBlock($ob, 0, $w, $null, 0)
        }
        [void]$hm.TransformFinalBlock((New-Object byte[] 0), 0, 0)
        $OutStream.Write($hm.Hash, 0, 32)
    } finally {
        if ($t) { $t.Dispose() }; if ($aes) { $aes.Dispose() }; if ($hm) { $hm.Dispose() }
        [Array]::Clear($keys.Enc, 0, 32); [Array]::Clear($keys.Mac, 0, 32)
    }
}

function Read-HUAppBackupHeaderStream {
    <# Liest nur den Kopf (NICHT authentisiert - nur zur Anzeige und fuer die Schluesselableitung). #>
    param([Parameter(Mandatory)][IO.Stream]$Stream)
    $m = $script:HUAbMagic.Length
    $total = $Stream.Length
    if ($total -lt $m + 4 + 2 + 16 + 32) { throw 'Keine HU-Sicherung (Datei zu kurz).' }
    [void]$Stream.Seek(0, 'Begin')
    $pre = New-Object byte[] ($m + 4)
    if ((Read-HUAbFull $Stream $pre ($m + 4)) -ne ($m + 4)) { throw 'Keine HU-Sicherung (Datei zu kurz).' }
    for ($i = 0; $i -lt $m; $i++) { if ($pre[$i] -ne $script:HUAbMagic[$i]) { throw 'Keine HU-Sicherung (Kennung fehlt).' } }
    $lb = New-Object byte[] 4; [Array]::Copy($pre, $m, $lb, 0, 4); if (-not [BitConverter]::IsLittleEndian) { [Array]::Reverse($lb) }
    $hl = [BitConverter]::ToInt32($lb, 0)
    if ($hl -le 1 -or $hl -gt $script:HUAbMaxHead -or ($m + 4 + $hl + 16 + 32) -gt $total) { throw 'Keine HU-Sicherung (Kopf ungueltig).' }
    $hb = New-Object byte[] $hl
    if ((Read-HUAbFull $Stream $hb $hl) -ne $hl) { throw 'Keine HU-Sicherung (Kopf ungueltig).' }
    $txt = [Text.Encoding]::UTF8.GetString($hb)
    try { $h = $txt | ConvertFrom-Json } catch { throw 'Keine HU-Sicherung (Kopf ungueltig).' }
    if ($null -eq $h -or "$($h.Format)" -ne '1') { throw "Format $(if ($h) { $h.Format }) wird von dieser Version nicht unterstuetzt." }
    if ("$($h.Kdf)" -ne 'PBKDF2-SHA256') { throw 'Unbekanntes Schluesselverfahren.' }
    $it = 0
    if (-not [int]::TryParse("$($h.Iter)", [ref]$it) -or $it -lt $script:HUAbIterMin -or $it -gt $script:HUAbIterMax) { throw 'Schluesselparameter ausserhalb des erlaubten Bereichs.' }
    try { $salt = [Convert]::FromBase64String("$($h.Salt)"); $iv = [Convert]::FromBase64String("$($h.IV)") } catch { throw 'Keine HU-Sicherung (Kopf ungueltig).' }
    if ($salt.Length -ne 32 -or $iv.Length -ne 16) { throw 'Keine HU-Sicherung (Kopf ungueltig).' }
    return [pscustomobject]@{ Header = $h; HeadEnd = [int64]($m + 4 + $hl); Iter = $it; Salt = $salt; IV = $iv }
}

function Read-HUAbFull([IO.Stream]$S, [byte[]]$Buf, [int]$Count) {
    $n = 0
    while ($n -lt $Count) { $r = $S.Read($Buf, $n, $Count - $n); if ($r -le 0) { break }; $n += $r }
    return $n
}

function Unprotect-HUAppBackupStream {
    <# 1. Durchgang: HMAC ueber alles (konstante Zeit vergleichen). 2. Durchgang erst danach: entschluesseln nach $OutStream. Gibt den Kopf zurueck. #>
    param([Parameter(Mandatory)][IO.Stream]$InStream, [Parameter(Mandatory)][IO.Stream]$OutStream, [Parameter(Mandatory)][object]$Password, [hashtable]$Progress = $null)
    $r = Read-HUAppBackupHeaderStream $InStream
    $pg = { param($ph, $d, $t) if ($null -ne $Progress) { $Progress['Phase'] = $ph; $Progress['Done'] = $d; $Progress['Total'] = $t } }
    & $pg 'Schluessel' 0 0
    $total = $InStream.Length
    $clen = $total - $r.HeadEnd - 32
    if ($clen -lt 16 -or ($clen % 16) -ne 0) { throw $script:HUAbFailMsg }
    $keys = Get-HUAbKeys -Password $Password -Salt $r.Salt -Iterations $r.Iter
    $aes = $null; $t = $null; $hm = $null
    try {
        $buf = New-Object byte[] 65536
        $hm = New-Object Security.Cryptography.HMACSHA256 -ArgumentList (, $keys.Mac)
        [void]$InStream.Seek(0, 'Begin')
        $left = $total - 32
        while ($left -gt 0) {
            $n = Read-HUAbFull $InStream $buf ([int][Math]::Min([int64]$buf.Length, $left))
            if ($n -le 0) { throw $script:HUAbFailMsg }
            [void]$hm.TransformBlock($buf, 0, $n, $null, 0); $left -= $n
            & $pg 'Pruefsumme' ($total - 32 - $left) ($total - 32)
        }
        [void]$hm.TransformFinalBlock((New-Object byte[] 0), 0, 0)
        $tag = New-Object byte[] 32
        if ((Read-HUAbFull $InStream $tag 32) -ne 32) { throw $script:HUAbFailMsg }
        if (-not (Test-HUAbEqual $tag $hm.Hash)) { throw $script:HUAbFailMsg }
        $aes = New-HUAbAes $keys.Enc $r.IV; $t = $aes.CreateDecryptor()
        [void]$InStream.Seek($r.HeadEnd, 'Begin')
        $left = $clen; $ob = New-Object byte[] (65536 + 32)
        try {
            while ($left -gt 0) {
                $want = [int][Math]::Min([int64]$buf.Length, $left)
                $n = Read-HUAbFull $InStream $buf $want
                if ($n -ne $want) { throw $script:HUAbFailMsg }
                $left -= $n
                & $pg 'Entschluesseln' ($clen - $left) $clen
                if ($left -eq 0) { $last = $t.TransformFinalBlock($buf, 0, $n); $OutStream.Write($last, 0, $last.Length) }
                else { $w = $t.TransformBlock($buf, 0, $n, $ob, 0); $OutStream.Write($ob, 0, $w) }
            }
        } catch { throw $script:HUAbFailMsg }
        return $r.Header
    } finally {
        if ($t) { $t.Dispose() }; if ($aes) { $aes.Dispose() }; if ($hm) { $hm.Dispose() }
        [Array]::Clear($keys.Enc, 0, 32); [Array]::Clear($keys.Mac, 0, 32)
    }
}

# Byte-Varianten (Tests, kleine Daten)
function Protect-HUAppBackupBytes {
    param([Parameter(Mandatory)][byte[]]$Plain, [Parameter(Mandatory)][object]$Password, [hashtable]$Header = @{}, [int]$Iterations = $script:HUAbIter)
    $i = New-Object IO.MemoryStream(, $Plain); $o = New-Object IO.MemoryStream
    try { Protect-HUAppBackupStream -InStream $i -OutStream $o -Password $Password -Header $Header -Iterations $Iterations; return , $o.ToArray() } finally { $i.Dispose(); $o.Dispose() }
}
function Read-HUAppBackupHeader {
    param([Parameter(Mandatory)][byte[]]$Data)
    $i = New-Object IO.MemoryStream(, $Data)
    try { return (Read-HUAppBackupHeaderStream $i) } finally { $i.Dispose() }
}
function Unprotect-HUAppBackupBytes {
    param([Parameter(Mandatory)][byte[]]$Data, [Parameter(Mandatory)][object]$Password)
    $i = New-Object IO.MemoryStream(, $Data); $o = New-Object IO.MemoryStream
    try { $h = Unprotect-HUAppBackupStream -InStream $i -OutStream $o -Password $Password; return @{ Plain = $o.ToArray(); Header = $h } } finally { $i.Dispose(); $o.Dispose() }
}

# ================================================================== Inhalt: Data\ (feste Regeln) ==================================================================
# Gesichert wird alles unter Data\ als "data/<Pfad>", ausser:
#   Docs\ (kommt aus dem Hilfe-Repo), Laufend\ (Sperren), Wiederherstellung\ (Vorstand), Benutzer\<Name>\Logs\, Benutzer\<Name>\Token\,
#   *.lock, *.tmp, sicherung-kennwort.dat (DPAPI, nur fuer diesen Benutzer lesbar - woanders nutzlos).
# Bereiche (beim Wiederherstellen einzeln waehlbar): Daten je Modul (Data\<Id>\), Benutzer (Fenster, Einstellungen), Allgemein.
$script:HUAbSkipTop = @('Docs', 'Laufend', 'Wiederherstellung', '_Test', '_Demo')   # _Test: Testbestand, _Demo: Demo-Bestaende
$script:HUAbSkipUserDirs = @('Logs', 'Token')
$script:HUAbSkipFileRx = '(?i)(\.lock|\.tmp|\.hu-neu|\.husicherung|^sicherung-kennwort\.dat)$'
$script:HUAbFileRx = '^HU-Sicherung_[A-Za-z0-9-]{1,40}_\d{4}-\d{2}-\d{2}_\d{6}\.husicherung$'
# alles wird beim Lesen im Speicher gehalten -> Gesamtgrenze 2 GB
$script:HUAbLimits = @{ Files = 200000; EntryBytes = 512MB; TotalBytes = 2GB }
# Bereiche, die beim Wiederherstellen nicht vorgewaehlt sind (wer Verwalter ist)
$script:HUAbCodeNames = @('data/Verwalter.txt')
function Test-HUAppBackupCodeName([string]$Name) { return ($script:HUAbCodeNames -contains $Name) }

function Get-HUAppBackupArea([string]$Name) {
    if (-not $Name.StartsWith('data/')) { return '' }
    $seg = $Name.Split('/')
    if ($seg.Count -eq 2) { if ($seg[1] -ceq 'Verwalter.txt') { return 'Verwalter' } return 'Allgemein' }
    if ($seg[1] -ceq 'Benutzer') { return 'Benutzer (Fenster und Einstellungen)' }
    return "Daten: $($seg[1])"
}

function Test-HUAbRelPath([string]$Rel) {
    # keine Laufwerke, kein .., keine reservierten Namen, keine Steuerzeichen
    if (-not $Rel -or $Rel -match '[\x00-\x1f]' -or $Rel.Contains(':') -or $Rel.Contains('\') -or $Rel.StartsWith('/')) { return $false }
    foreach ($seg in $Rel.Split('/')) {
        if ($seg -eq '' -or $seg -in '.', '..' -or $seg.EndsWith('.') -or $seg.EndsWith(' ')) { return $false }
        if ($seg -match '^(?i)(CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9])(\..*)?$') { return $false }
    }
    return $true
}

# Gehoert ein relativer Pfad (unter Data, mit /) in die Sicherung?
function Test-HUAbDataRel([string]$Rel) {
    if (-not (Test-HUAbRelPath $Rel)) { return $false }
    $seg = $Rel.Split('/')
    if ($script:HUAbSkipTop -contains $seg[0]) { return $false }
    if ($seg[-1] -match $script:HUAbSkipFileRx) { return $false }
    if ($seg[0] -ceq 'Benutzer' -and $seg.Count -ge 3 -and $script:HUAbSkipUserDirs -contains $seg[2]) { return $false }
    return $true
}

function Resolve-HUAppBackupTarget {
    <# Name in der Sicherung -> Zielpfad, oder '' wenn nicht erlaubt. #>
    param([string]$Name, [string]$AppRoot, [string]$DataRoot)
    if (-not (Test-HUAbRelPath $Name)) { return '' }
    if (-not $Name.StartsWith('data/')) { return '' }
    $rel = $Name.Substring(5)
    if (-not (Test-HUAbDataRel $rel)) { return '' }
    return (Join-Path $DataRoot ($rel.Replace('/', [IO.Path]::DirectorySeparatorChar)))
}

function Get-HUAbFiles([string]$Dir) {
    if (-not $Dir -or -not (Test-Path -LiteralPath $Dir -PathType Container)) { return @() }
    # keine Verknuepfungen (Reparse-Points) verfolgen
    $out = New-Object System.Collections.Generic.List[object]
    $stack = New-Object System.Collections.Generic.Stack[string]; $stack.Push((Get-Item -LiteralPath $Dir).FullName)
    while ($stack.Count) {
        $d = $stack.Pop()
        foreach ($i in @(Get-ChildItem -LiteralPath $d -Force -ErrorAction SilentlyContinue)) {
            if ($i.Attributes -band [IO.FileAttributes]::ReparsePoint) { continue }
            if ($i.PSIsContainer) { $stack.Push($i.FullName) } else { $out.Add($i) }
        }
    }
    return $out.ToArray()
}

function Get-HUAppBackupSources {
    <# Welche Dateien in die Sicherung gehen: @{ Name; Path; Area }. #>
    param([Parameter(Mandatory)][string]$DataRoot)
    $out = New-Object System.Collections.Generic.List[object]
    if (-not (Test-Path -LiteralPath $DataRoot -PathType Container)) { return @() }
    $base = (Get-Item -LiteralPath $DataRoot).FullName.TrimEnd('\', '/')
    foreach ($i in @(Get-HUAbFiles $DataRoot)) {
        $rel = $i.FullName.Substring($base.Length + 1).Replace([IO.Path]::DirectorySeparatorChar, '/')
        if (-not (Test-HUAbDataRel $rel)) { continue }
        $n = "data/$rel"
        $out.Add([pscustomobject]@{ Name = $n; Path = $i.FullName; Area = (Get-HUAppBackupArea $n) })
    }
    return $out.ToArray()
}

function Read-HUAbFileBytes([string]$Path) {
    # die App kann gerade schreiben -> mit Wiederholung und FileShare.ReadWrite lesen
    for ($try = 1; ; $try++) {
        try {
            $fs = New-Object IO.FileStream($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
            try { $ms = New-Object IO.MemoryStream; $fs.CopyTo($ms); return , $ms.ToArray() } finally { $fs.Dispose() }
        } catch { if ($try -ge 3) { throw }; Start-Sleep -Milliseconds 300 }
    }
}

function Get-HUAbSha256([byte[]]$Bytes) {
    $h = [Security.Cryptography.SHA256]::Create()
    try { return (($h.ComputeHash($Bytes) | ForEach-Object { $_.ToString('x2') }) -join '') } finally { $h.Dispose() }
}

function Get-HUAbModuleList([string[]]$Dirs) {
    $l = New-Object System.Collections.Generic.List[object]
    foreach ($AddonsDir in @($Dirs)) {
        if (-not $AddonsDir -or -not (Test-Path -LiteralPath $AddonsDir -PathType Container)) { continue }
        foreach ($d in @(Get-ChildItem -LiteralPath $AddonsDir -Directory -ErrorAction SilentlyContinue)) {
            $f = Join-Path $d.FullName 'addon.json'
            if (-not (Test-Path -LiteralPath $f)) { continue }
            try { $j = [IO.File]::ReadAllText($f, [Text.Encoding]::UTF8) | ConvertFrom-Json; $l.Add([pscustomobject]@{ Id = "$($j.id)"; Title = "$($j.title)"; Version = "$($j.version)" }) } catch { }
        }
    }
    return $l.ToArray()
}

function ConvertTo-HUAbPcName([string]$Pc) {
    $p = ("$Pc" -replace '[^A-Za-z0-9-]', '-').Trim('-'); if (-not $p) { $p = 'PC' }
    if ($p.Length -gt 40) { $p = $p.Substring(0, 40) }
    return $p
}
function Get-HUAppBackupFileName([string]$Pc, [datetime]$Now = (Get-Date)) {
    $p = ConvertTo-HUAbPcName $Pc
    return "HU-Sicherung_$($p)_$($Now.ToString('yyyy-MM-dd_HHmmss')).husicherung"
}

function New-HUAppBackup {
    <#
    .SYNOPSIS
        Sicherung schreiben: ZIP (Manifest mit SHA-256) -> verschluesseln -> .tmp -> Flush -> zuruecklesen und MAC pruefen -> umbenennen -> alte aufraeumen.
        Rueckgabe: @{ Path; Files; Bytes; Removed }.
    #>
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Sources, [Parameter(Mandatory)][object]$Password, [Parameter(Mandatory)][string]$Folder,
        [Parameter(Mandatory)][string]$AppVersion, [string]$Pc = $env:COMPUTERNAME, [int]$Keep = 10, [object[]]$Modules = @(),
        [int]$Iterations = $script:HUAbIter, [datetime]$Now = (Get-Date), [hashtable]$Progress = $null)
    $why = Test-HUAbPassword $Password
    if ($why) { throw $why }
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { New-Item -ItemType Directory -Path $Folder -Force | Out-Null }
    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
    $name = Get-HUAppBackupFileName $Pc $Now
    $final = Join-Path $Folder $name
    $tmp = "$final.tmp"
    Clear-HUAbTempZips
    $total = [int64]0; foreach ($s in $Sources) { try { $total += (Get-Item -LiteralPath $s.Path).Length } catch { } }
    # grosse Datenmengen nicht im Speicher zippen, sondern in einer Temp-Datei, die beim Schliessen verschwindet
    $zs = if ($total -gt 256MB) { New-Object IO.FileStream((Join-Path ([IO.Path]::GetTempPath()) ("HU-Sicherung-" + [guid]::NewGuid().ToString('N') + '.zip')), [IO.FileMode]::CreateNew, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None, 65536, [IO.FileOptions]::DeleteOnClose) } else { New-Object IO.MemoryStream }
    $man = New-Object System.Collections.Generic.List[object]
    $bytes = [int64]0
    $prog = { param($ph, $d) if ($null -ne $Progress) { $Progress['Phase'] = $ph; $Progress['Done'] = $d; $Progress['Total'] = $Sources.Count } }
    & $prog 'Dateien' 0
    try {
        $zip = New-Object IO.Compression.ZipArchive($zs, [IO.Compression.ZipArchiveMode]::Create, $true)
        try {
            foreach ($s in $Sources) {
                $data = Read-HUAbFileBytes $s.Path
                $lw = $null; try { $lw = [IO.File]::GetLastWriteTime($s.Path) } catch { }
                $e = $zip.CreateEntry($s.Name, [IO.Compression.CompressionLevel]::Optimal)
                if ($lw) { try { $e.LastWriteTime = $lw } catch { } }
                $w = $e.Open(); try { $w.Write($data, 0, $data.Length) } finally { $w.Dispose() }
                $man.Add([pscustomobject]@{ Name = $s.Name; Size = $data.Length; Sha256 = (Get-HUAbSha256 $data); Changed = $(if ($lw) { $lw.ToString('s') } else { '' }) })
                $bytes += $data.Length
                if (($man.Count % 25) -eq 0) { & $prog 'Dateien' $man.Count }
            }
            $mj = [Text.Encoding]::UTF8.GetBytes((([pscustomobject]@{ Format = 1; App = $AppVersion; Created = $Now.ToString('s'); Files = $man.ToArray(); Modules = @($Modules) }) | ConvertTo-Json -Depth 6))
            $e = $zip.CreateEntry('meta/manifest.json'); $w = $e.Open(); try { $w.Write($mj, 0, $mj.Length) } finally { $w.Dispose() }
        } finally { $zip.Dispose() }
        [void]$zs.Seek(0, 'Begin')
        & $prog 'Verschluesseln' $Sources.Count
        $areas = @($man | ForEach-Object { Get-HUAppBackupArea $_.Name } | Where-Object { $_ } | Select-Object -Unique)
        $fs = New-Object IO.FileStream($tmp, [IO.FileMode]::Create, [IO.FileAccess]::Write, [IO.FileShare]::None, 65536, [IO.FileOptions]::WriteThrough)
        try {
            Protect-HUAppBackupStream -InStream $zs -OutStream $fs -Password $Password -Iterations $Iterations -Header @{ App = $AppVersion; Files = $man.Count; Areas = $areas }
            $fs.Flush($true)
        } finally { $fs.Dispose() }
    } catch { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue; throw } finally { $zs.Dispose() }
    # zuruecklesen und pruefen (Netzwerkfreigabe), erst dann umbenennen
    & $prog 'Pruefen' $Sources.Count
    try {
        $rs = New-Object IO.FileStream($tmp, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        try { [void](Unprotect-HUAppBackupStream -InStream $rs -OutStream ([IO.Stream]::Null) -Password $Password) } finally { $rs.Dispose() }
    } catch { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue; throw "Sicherung nach dem Schreiben nicht lesbar - verworfen: $($_.Exception.Message)" }
    Move-Item -LiteralPath $tmp -Destination $final -Force
    $removed = @(Remove-HUAppBackupOld -Folder $Folder -Keep $Keep -Pc $Pc)
    return [pscustomobject]@{ Path = $final; Files = $man.Count; Bytes = $bytes; Removed = $removed }
}

function Get-HUAppBackupList([string]$Folder) {
    if (-not $Folder -or -not (Test-Path -LiteralPath $Folder -PathType Container)) { return @() }
    return @(Get-ChildItem -LiteralPath $Folder -File -Filter 'HU-Sicherung_*.husicherung' -ErrorAction SilentlyContinue | Where-Object { $_.Name -match $script:HUAbFileRx } | Sort-Object LastWriteTime -Descending)
}

function Remove-HUAppBackupOld {
    <# Nur eigene Dateinamen DIESES PCs (gemeinsamer Ordner mehrerer PCs); die neuesten $Keep bleiben. Rueckgabe: Namen der geloeschten. #>
    param([Parameter(Mandatory)][string]$Folder, [int]$Keep = 10, [string]$Pc = $env:COMPUTERNAME)
    if ($Keep -lt 1) { $Keep = 1 }
    $pre = "HU-Sicherung_$(ConvertTo-HUAbPcName $Pc)_"
    $old = @(Get-HUAppBackupList $Folder | Where-Object { $_.Name.StartsWith($pre, [StringComparison]::OrdinalIgnoreCase) -and $_.Name.Length -eq ($pre.Length + 29) } | Select-Object -Skip $Keep)
    foreach ($f in $old) { try { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop; $f.Name } catch { } }
}

function Open-HUAppBackup {
    <#
    .SYNOPSIS
        Entschluesseln und pruefen (MAC, Format, App-Version, ZIP-Grenzen, nur erlaubte Namen, SHA-256 laut Manifest).
        Rueckgabe: @{ Header; Manifest; Entries (Name -> byte[]); Rows (Name; Area; Target; Size; Changed; Exists; CurChanged); Errors }
        Schreibt nichts.
    #>
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][object]$Password, [Parameter(Mandatory)][string]$AppVersion,
        [Parameter(Mandatory)][string]$AppRoot, [Parameter(Mandatory)][string]$DataRoot, [hashtable]$Progress = $null)
    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
    $pg = { param($ph, $d, $t) if ($null -ne $Progress) { $Progress['Phase'] = $ph; $Progress['Done'] = $d; $Progress['Total'] = $t } }
    $res = [pscustomobject]@{ Header = $null; Manifest = $null; Entries = @{}; Rows = @(); Errors = (New-Object System.Collections.Generic.List[string]) }
    Clear-HUAbTempZips
    $fs = Open-HUAbSourceStream $Path
    $big = $fs.Length -gt 256MB
    $zs = if ($big) { New-Object IO.FileStream((Join-Path ([IO.Path]::GetTempPath()) ("HU-Wiederherstellung-" + [guid]::NewGuid().ToString('N') + '.zip')), [IO.FileMode]::CreateNew, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None, 65536, [IO.FileOptions]::DeleteOnClose) } else { New-Object IO.MemoryStream }
    try {
        try { $res.Header = Unprotect-HUAppBackupStream -InStream $fs -OutStream $zs -Password $Password -Progress $Progress } finally { $fs.Dispose() }
        # aus neuerer App -> ablehnen
        $hv = $null; $av = $null
        try { $hv = [version]"$($res.Header.App)" } catch { }
        try { $av = [version]"$AppVersion" } catch { }
        if (-not $hv) { $res.Errors.Add('App-Version der Sicherung fehlt.'); return $res }
        if ($av -and $hv -gt $av) { $res.Errors.Add("Sicherung stammt aus Version $hv, laufend ist $av - zuerst aktualisieren."); return $res }
        [void]$zs.Seek(0, 'Begin')
        $zip = New-Object IO.Compression.ZipArchive($zs, [IO.Compression.ZipArchiveMode]::Read, $true)
        try {
            $L = $script:HUAbLimits
            if ($zip.Entries.Count -gt $L.Files) { $res.Errors.Add("Zu viele Dateien ($($zip.Entries.Count))."); return $res }
            $seen = @{}; $total = [int64]0; $cnt = $zip.Entries.Count; $i = 0
            foreach ($ze in $zip.Entries) {
                $i++; if (($i % 25) -eq 0) { & $pg 'Lesen' $i $cnt }
                $n = $ze.FullName
                if ($n.EndsWith('/')) { continue }
                $k = $n.Normalize([Text.NormalizationForm]::FormC).ToLowerInvariant()
                if ($seen.ContainsKey($k)) { $res.Errors.Add("'$n' doppelt."); continue }
                $seen[$k] = $true
                if ($n -cne 'meta/manifest.json' -and -not (Resolve-HUAppBackupTarget $n $AppRoot $DataRoot)) { $res.Errors.Add("'$n' ist kein erlaubter Inhalt."); continue }
                $s = $ze.Open(); $out = New-Object IO.MemoryStream; $buf = New-Object byte[] 65536
                try {
                    while (($r = $s.Read($buf, 0, $buf.Length)) -gt 0) {
                        $out.Write($buf, 0, $r)
                        if ($out.Length -gt $L.EntryBytes -or ($total + $out.Length) -gt $L.TotalBytes) { throw "'$n' zu gross." }
                    }
                } finally { $s.Dispose() }
                $total += $out.Length
                $res.Entries[$n] = $out.ToArray()
            }
        } finally { $zip.Dispose() }
    } catch { $res.Errors.Add("$($_.Exception.Message)"); return $res } finally { $zs.Dispose() }
    if ($res.Errors.Count) { return $res }
    if (-not $res.Entries.ContainsKey('meta/manifest.json')) { $res.Errors.Add('Manifest fehlt.'); return $res }
    try { $res.Manifest = [Text.Encoding]::UTF8.GetString($res.Entries['meta/manifest.json']) | ConvertFrom-Json } catch { $res.Errors.Add('Manifest nicht lesbar.'); return $res }
    if ("$($res.Manifest.Format)" -ne '1') { $res.Errors.Add("Manifest-Format $($res.Manifest.Format) unbekannt."); return $res }
    $mf = @{}; foreach ($m in @($res.Manifest.Files)) { $mf["$($m.Name)"] = $m }
    $rows = New-Object System.Collections.Generic.List[object]
    $keys = @($res.Entries.Keys | Where-Object { $_ -cne 'meta/manifest.json' } | Sort-Object); $i = 0
    foreach ($n in $keys) {
        $i++; if (($i % 25) -eq 0) { & $pg 'Vergleichen' $i $keys.Count }
        $m = $mf[$n]
        if (-not $m) { $res.Errors.Add("'$n' steht nicht im Manifest."); continue }
        if ((Get-HUAbSha256 $res.Entries[$n]) -ne "$($m.Sha256)") { $res.Errors.Add("'$n': Pruefsumme passt nicht."); continue }
        $t = Resolve-HUAppBackupTarget $n $AppRoot $DataRoot
        $cur = Get-Item -LiteralPath $t -ErrorAction SilentlyContinue
        $rows.Add([pscustomobject]@{ Name = $n; Area = (Get-HUAppBackupArea $n); Target = $t; Size = $res.Entries[$n].Length; Changed = "$($m.Changed)"; Exists = [bool]$cur; CurChanged = $(if ($cur) { $cur.LastWriteTime.ToString('s') } else { '' }); CurSize = $(if ($cur) { $cur.Length } else { 0 }) })
    }
    foreach ($k in $mf.Keys) { if (-not $res.Entries.ContainsKey($k)) { $res.Errors.Add("'$k' fehlt in der Sicherung.") } }
    $res.Rows = $rows.ToArray()
    return $res
}

function Restore-HUAppBackup {
    <#
    .SYNOPSIS
        Schreibt die gewaehlten Bereiche aus einer mit Open-HUAppBackup geprueften Sicherung.
        Vorher: aktueller Stand der betroffenen Dateien nach <DataRoot>\Wiederherstellung\<Zeit>\ (die letzten 5 bleiben).
        Rueckgabe: @{ Written; Before (Ordner); Modules (aus der Sicherung) }
    #>
    param([Parameter(Mandatory)]$Opened, [Parameter(Mandatory)][string[]]$Areas, [Parameter(Mandatory)][string]$AppRoot, [Parameter(Mandatory)][string]$DataRoot, [datetime]$Now = (Get-Date))
    if ($Opened.Errors.Count) { throw "Sicherung nicht in Ordnung: $($Opened.Errors -join ' ')" }
    $rows = @($Opened.Rows | Where-Object { $Areas -contains $_.Area })
    if (-not $rows.Count) { throw 'Nichts ausgewaehlt.' }
    # Vorstand
    $vroot = Join-Path $DataRoot 'Wiederherstellung'
    $vdir = Join-Path $vroot $Now.ToString('yyyy-MM-dd_HHmmss')
    foreach ($r in $rows) {
        if (Test-Path -LiteralPath $r.Target -PathType Leaf) {
            $dst = Join-Path $vdir ($r.Name.Replace('/', [IO.Path]::DirectorySeparatorChar))
            New-Item -ItemType Directory -Path (Split-Path -Parent $dst) -Force | Out-Null
            Copy-Item -LiteralPath $r.Target -Destination $dst -Force
        }
    }
    foreach ($old in @(Get-ChildItem -LiteralPath $vroot -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}_\d{6}$' } | Sort-Object Name -Descending | Select-Object -Skip 5)) {
        Remove-Item -LiteralPath $old.FullName -Recurse -Force -ErrorAction SilentlyContinue
    }
    $written = 0
    foreach ($r in $rows) {
        New-Item -ItemType Directory -Path (Split-Path -Parent $r.Target) -Force | Out-Null
        $tmp = "$($r.Target).hu-neu"
        [IO.File]::WriteAllBytes($tmp, $Opened.Entries[$r.Name])
        Move-Item -LiteralPath $tmp -Destination $r.Target -Force
        $written++
    }
    return [pscustomobject]@{ Written = $written; Before = $(if (Test-Path -LiteralPath $vdir) { $vdir } else { '' }); Modules = @($Opened.Manifest.Modules) }
}

# ================================================================== Kennwort fuer die Automatik (DPAPI) ==================================================================
$script:HUAbEntropy = [Text.Encoding]::UTF8.GetBytes('HU-Amt App-Sicherung 1')
function Get-HUAbPwFile([string]$DataRoot) { return (Join-Path $DataRoot 'sicherung-kennwort.dat') }
function Save-HUAppBackupPassword {
    param([Parameter(Mandatory)][string]$DataRoot, [Parameter(Mandatory)][object]$Password)
    $why = Test-HUAbPassword $Password
    if ($why) { throw $why }
    Add-Type -AssemblyName System.Security -ErrorAction SilentlyContinue
    $b = [Security.Cryptography.ProtectedData]::Protect([Text.Encoding]::UTF8.GetBytes((Get-HUAbPlain $Password)), $script:HUAbEntropy, [Security.Cryptography.DataProtectionScope]::CurrentUser)
    if (-not (Test-Path -LiteralPath $DataRoot)) { New-Item -ItemType Directory -Path $DataRoot -Force | Out-Null }
    [IO.File]::WriteAllText((Get-HUAbPwFile $DataRoot), [Convert]::ToBase64String($b))
}
function Get-HUAppBackupPassword {
    <# Gespeichertes Kennwort als SecureString, oder $null. #>
    param([Parameter(Mandatory)][string]$DataRoot)
    $f = Get-HUAbPwFile $DataRoot
    if (-not (Test-Path -LiteralPath $f)) { return $null }
    try {
        Add-Type -AssemblyName System.Security -ErrorAction SilentlyContinue
        $b = [Security.Cryptography.ProtectedData]::Unprotect([Convert]::FromBase64String((Get-Content -LiteralPath $f -Raw).Trim()), $script:HUAbEntropy, [Security.Cryptography.DataProtectionScope]::CurrentUser)
        $s = New-Object Security.SecureString
        foreach ($ch in [Text.Encoding]::UTF8.GetString($b).ToCharArray()) { $s.AppendChar($ch) }
        [Array]::Clear($b, 0, $b.Length); $s.MakeReadOnly()
        return $s
    } catch { return $null }
}
function Remove-HUAppBackupPassword([string]$DataRoot) { Remove-Item -LiteralPath (Get-HUAbPwFile $DataRoot) -Force -ErrorAction SilentlyContinue }

function Get-HUAppBackupWarning {
    <#
        Warnung fuer Verwalter, wenn die Sicherung fehlt oder ueberfaellig ist. '' = alles in Ordnung.
        Ueberfaellig: taeglich seit 2 Kalendertagen, woechentlich seit 9 Kalendertagen keine erfolgreiche Sicherung.
    #>
    param([string]$Mode, [string]$Path, [object]$Last, [datetime]$Now = (Get-Date))
    if (-not "$Path".Trim()) { return 'Die Sicherung ist nicht eingerichtet: kein Sicherungsordner.' }
    if ($Mode -notin 'taeglich', 'woechentlich') { return 'Die automatische Sicherung ist aus.' }
    $l = $null; try { if ("$Last") { $l = [datetime]::Parse("$Last", [Globalization.CultureInfo]::InvariantCulture) } } catch { }
    if (-not $l) { return 'Es gibt noch keine erfolgreiche Sicherung.' }
    $days = ($Now.Date - $l.Date).Days
    if ($days -ge $(if ($Mode -eq 'taeglich') { 2 } else { 9 })) { return "Die letzte erfolgreiche Sicherung ist $days Tage alt." }
    return ''
}

function Test-HUAppBackupDue {
    <#
        Automatik faellig? $Mode 'aus'|'taeglich'|'woechentlich', $Last = Zeitpunkt der letzten erfolgreichen Sicherung.
        Nach Kalendertagen (Festlegung 27.10.): taeglich = heute noch keine Sicherung, woechentlich = seit 7 Kalendertagen keine.
        So bleibt der Zeitpunkt stabil - beim ersten Start des Tages bzw. kurz nach Mitternacht, wenn die App durchlaeuft.
    #>
    param([string]$Mode, [object]$Last, [datetime]$Now = (Get-Date))
    if ($Mode -notin 'taeglich', 'woechentlich') { return $false }
    $l = $null; try { if ("$Last") { $l = [datetime]::Parse("$Last", [Globalization.CultureInfo]::InvariantCulture) } } catch { }
    if (-not $l) { return $true }
    $days = ($Now.Date - $l.Date).Days
    if ($Mode -eq 'taeglich') { return ($days -ge 1) }
    return ($days -ge 7)
}

Export-ModuleMember -Function Test-HUAbPassword, Protect-HUAppBackupStream, Read-HUAppBackupHeaderStream, Unprotect-HUAppBackupStream,
    Protect-HUAppBackupBytes, Read-HUAppBackupHeader, Unprotect-HUAppBackupBytes,
    Get-HUAppBackupArea, Test-HUAppBackupCodeName, Resolve-HUAppBackupTarget, Get-HUAppBackupSources, Get-HUAbModuleList, Get-HUAppBackupFileName, New-HUAppBackup, Get-HUAppBackupList,
    Remove-HUAppBackupOld, Open-HUAppBackup, Restore-HUAppBackup, Save-HUAppBackupPassword, Get-HUAppBackupPassword, Remove-HUAppBackupPassword, Test-HUAppBackupDue, Get-HUAppBackupWarning
