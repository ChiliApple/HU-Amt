#Requires -Version 5.1
<#
.SYNOPSIS
    Module: jeder Reiter ist ein Modul (mitgeliefert, signiert aus einem privaten Repository oder in Entwicklung).
.DESCRIPTION
    Aufbau eines Moduls (Ordner mit dem Namen der Kennung):
      addon.json   { "id": "<Kennung>", "title": "...", "version": "1.0.0", "minApp": "0.0.1", "order": 10,
                     "tab": { "header": "...", "xaml": "Tab.xaml", "script": "Addon.ps1", "init": "Initialize-HUAddon<Kennung>" },
                     "drop": [".pdf", ".xlsx"] }
      Tab.xaml     Wurzel <TabItem>; Theme ueber <TabItem.Resources><!--HU:THEME--></TabItem.Resources>
      Addon.ps1    Funktionen des Moduls (wird in den Programm-Bereich geladen)
    Funktionen im Modul (Konvention statt Schnittstelle):
      Initialize-HUAddon<Id>           Pflicht, wird nach dem Einbau des Reiters aufgerufen
      Stop-HUAddon<Id>                 nur wenn das Modul etwas im Hintergrund laufen hat (beim Beenden)
      Invoke-HUAddonDrop<Id> -Paths    nur wenn der Reiter Dateien annimmt (Hineinziehen ins Fenster); Endungen in "drop"
    Herkunft:
      Modules\<Id>\     mitgeliefert, Teil des Programms (ohne eigene Signatur, kommt mit dem Programm-Update)
      Addons\<Id>\      privat: Dateien + addon.sha256 + addon.sha256.p7s (Herausgeber-Zertifikat); neue Fassungen warten
                        in Addons\.pending\<Id>\ und werden beim naechsten Start uebernommen
      AddonsDev\<Id>\   Entwicklung, ohne Signaturpruefung, ersetzt ein Modul gleicher Kennung
    Daten der Module: Data\<Id>\ (gemeinsam fuer alle Benutzer, mit Sicherung).
    Tokens: Data\Benutzer\<Name>\Token\<owner>_<repo>.xml (DPAPI, nur fuer diesen Windows-Benutzer lesbar).
    Zielmaschine: der PC, auf dem die App laeuft (Windows PowerShell 5.1 und PowerShell 7).
#>

$script:HUAddonManifest = 'addon.sha256'
$script:HUAddonSignature = 'addon.sha256.p7s'

function Get-HUAddonsDir { return (Join-Path $script:AppRoot 'Addons') }
function Get-HUAddonTokenDir {
    $d = Join-Path $script:UserRoot 'Token'
    if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
    return $d
}

# Datenordner je Modul: Data\<Id>\ - gemeinsam fuer alle Benutzer, die Sicherung nimmt ihn mit
function Get-HUAddonDataDir([string]$Id) {
    if (-not (Test-HUAddonId $Id)) { throw "Ungueltige Kennung fuer Modul-Daten: '$Id'" }
    return (Get-HUModuleDataPath -DataRoot $script:DataRoot -ModuleId $Id)
}

# Kennung eines Moduls: Buchstaben/Ziffern, beginnt mit Buchstabe (wird Teil von Funktions- und Ordnernamen)
function Test-HUAddonId([string]$Id) { return ("$Id" -cmatch '^[A-Za-z][A-Za-z0-9]{1,30}$' -and "$Id" -notin 'Benutzer', 'Laufend', 'Docs', 'Postfach', 'Wiederherstellung') }
function Test-HUAddonRepoName([string]$Owner, [string]$Repo) { return ("$Owner" -match '^[A-Za-z0-9-]{1,39}$' -and "$Repo" -match '^[A-Za-z0-9._-]{1,100}$') }

# --- Tokens (DPAPI) --------------------------------------------------------------------------
function Save-HUAddonToken([string]$Owner, [string]$Repo, [Security.SecureString]$Token) {
    if (-not (Test-HUAddonRepoName $Owner $Repo)) { throw 'Repository-Name ungueltig (Besitzer/Name)' }
    $cred = New-Object System.Management.Automation.PSCredential -ArgumentList "$Owner/$Repo", $Token
    $cred | Export-Clixml -LiteralPath (Join-Path (Get-HUAddonTokenDir) "$Owner`_$Repo.xml") -Force
}
function Remove-HUAddonToken([string]$Owner, [string]$Repo) {
    Remove-Item -LiteralPath (Join-Path (Get-HUAddonTokenDir) "$Owner`_$Repo.xml") -Force -ErrorAction SilentlyContinue
}
# Hinterlegte Module: Owner, Repo, Token (Klartext nur im Speicher)
function Get-HUAddonSources {
    $out = @()
    foreach ($f in @(Get-ChildItem -LiteralPath (Get-HUAddonTokenDir) -Filter '*.xml' -File -ErrorAction SilentlyContinue)) {
        try {
            $c = Import-Clixml -LiteralPath $f.FullName
            $parts = "$($c.UserName)" -split '/', 2
            if ($parts.Count -ne 2 -or -not (Test-HUAddonRepoName $parts[0] $parts[1])) { continue }
            $out += [pscustomobject]@{ Owner = $parts[0]; Repo = $parts[1]; Token = $c.GetNetworkCredential().Password }
        } catch { }
    }
    return $out
}

# --- Modul-Datei (*.humodule): mehrere Repositories mit Token auf einmal hinterlegen -------------
# Aufbau (JSON): { "module": [ { "repo": "Besitzer/Name", "token": "..." }, ... ] }
# Ein "token" auf oberster Ebene gilt fuer alle Eintraege ohne eigenen Token.
$script:HUAddonKeyFileExt = '.humodule'
function Test-HUAddonKeyFile([string]$Path) { return ([IO.Path]::GetExtension("$Path").ToLowerInvariant() -eq $script:HUAddonKeyFileExt) }
# Rueckgabe: Liste @{ Owner; Repo; Token; Error } - Error leer, wenn der Eintrag brauchbar ist. Wirft bei unlesbarer Datei.
function Read-HUAddonKeyFile([string]$Path) {
    $raw = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
    try { $j = $raw | ConvertFrom-Json -ErrorAction Stop } catch { throw 'keine gültige Modul-Datei (JSON nicht lesbar)' }
    if (-not $j -or -not $j.PSObject.Properties['module'] -or -not @($j.module).Count) { throw 'keine gültige Modul-Datei (Liste "module" fehlt oder ist leer)' }
    $common = $(if ($j.PSObject.Properties['token']) { "$($j.token)".Trim() } else { '' })
    $out = @()
    foreach ($e in @($j.module)) {
        $repo = $(if ($e -and $e.PSObject.Properties['repo']) { "$($e.repo)".Trim() } else { '' })
        $tok = $(if ($e -and $e.PSObject.Properties['token'] -and "$($e.token)".Trim()) { "$($e.token)".Trim() } else { $common })
        $parts = $repo -split '/', 2
        $r = [pscustomobject]@{ Owner = ''; Repo = ''; Token = $tok; Error = '' }
        if ($parts.Count -eq 2 -and (Test-HUAddonRepoName $parts[0] $parts[1])) { $r.Owner = $parts[0]; $r.Repo = $parts[1] }
        else { $r.Error = "Repository ungültig: '$repo'" }
        if (-not $r.Error -and -not $tok) { $r.Error = 'Token fehlt' }
        $out += $r
    }
    return $out
}
# Kann der Token das Repository lesen? Rueckgabe: '' (ja) oder Grund. Laeuft im Hintergrund (Netz).
function Test-HUAddonTokenAccess([string]$Owner, [string]$Repo, [string]$Token) {
    $ua = $(if ($script:HMUserAgent) { $script:HMUserAgent } else { 'HU-Amt' })
    $h = @{ Accept = 'application/vnd.github+json'; 'User-Agent' = $ua; Authorization = "token $Token" }
    try {
        $r = Invoke-RestMethod "https://api.github.com/repos/$Owner/$Repo" -Headers $h -UseBasicParsing -TimeoutSec 20 -ErrorAction Stop
        if ("$($r.full_name)" -ine "$Owner/$Repo") { return 'Antwort passt nicht zum Repository' }
        return ''
    } catch {
        $code = 0; try { $code = [int]$_.Exception.Response.StatusCode } catch { }
        switch ($code) {
            401 { return 'Token ungültig oder abgelaufen' }
            403 { return 'Token hat keine Leserechte' }
            404 { return 'Repository nicht gefunden oder Token ohne Zugriff' }
            default { return "nicht erreichbar: $($_.Exception.Message)" }
        }
    }
}

# --- Pruefen ---------------------------------------------------------------------------------
# addon.json lesen und pruefen. Rueckgabe: Objekt mit Error = '' (gueltig) oder Grund.
function Read-HUAddonInfo([string]$Dir) {
    $o = [pscustomobject]@{ Id = ''; Title = ''; Version = $null; MinApp = $null; Order = 100; Header = ''; Xaml = ''; Script = ''; Init = ''; Drop = @(); Source = ''; Dir = $Dir; Error = '' }
    $f = Join-Path $Dir 'addon.json'
    if (-not (Test-Path -LiteralPath $f)) { $o.Error = 'addon.json fehlt'; return $o }
    try { $j = [IO.File]::ReadAllText($f, [Text.Encoding]::UTF8) | ConvertFrom-Json } catch { $o.Error = 'addon.json nicht lesbar'; return $o }
    $o.Id = "$($j.id)"; $o.Title = "$($j.title)"
    if (-not (Test-HUAddonId $o.Id)) { $o.Error = 'Kennung (id) ungueltig'; return $o }
    $o.Version = ConvertTo-HMVersion "$($j.version)"; $o.MinApp = ConvertTo-HMVersion "$($j.minApp)"
    if (-not $o.Version -or -not $o.MinApp) { $o.Error = 'version/minApp ungueltig'; return $o }
    if ($j.PSObject.Properties['order']) { try { $o.Order = [int]$j.order } catch { $o.Error = 'order ungueltig'; return $o } }
    $rel = { param($p) if ("$p" -and "$p" -notmatch '(^[\\/])|(\.\.)|:' -and "$p" -match '^[A-Za-z0-9._\\/ -]+$') { "$p" } else { '' } }
    if (-not $j.PSObject.Properties['tab'] -or -not $j.tab) { $o.Error = 'tab fehlt'; return $o }
    $o.Header = "$($j.tab.header)"; $o.Xaml = & $rel $j.tab.xaml; $o.Script = & $rel $j.tab.script; $o.Init = "$($j.tab.init)"
    if (-not $o.Xaml -or -not $o.Script -or $o.Init -cnotmatch '^Initialize-HUAddon[A-Za-z0-9]+$') { $o.Error = 'tab: xaml/script/init ungueltig'; return $o }
    if ($j.PSObject.Properties['drop']) {
        $ext = @($j.drop | ForEach-Object { "$_".Trim().ToLowerInvariant() })
        if (@($ext | Where-Object { $_ -notmatch '^\.[a-z0-9]{1,10}$' }).Count) { $o.Error = 'drop: Endungen ungueltig (z. B. ".pdf")'; return $o }
        $o.Drop = $ext
    }
    return $o
}

# Alle Dateien eines Modul-Ordners gegen die signierte Pruefsummen-Datei pruefen. '' = gueltig, sonst Grund.
# Ausfuehrbare Dateien (ps1, psm1, xaml), die nicht in der Pruefsummen-Datei stehen, machen das Modul ungueltig.
function Test-HUAddonFiles([string]$Dir, [string]$Thumbprint) {
    $mf = Join-Path $Dir $script:HUAddonManifest; $sf = Join-Path $Dir $script:HUAddonSignature
    if (-not (Test-Path -LiteralPath $mf) -or -not (Test-Path -LiteralPath $sf)) { return 'Pruefsummen- oder Signatur-Datei fehlt' }
    $man = [System.IO.File]::ReadAllBytes($mf)
    $why = Test-HMManifestSignature $man ([System.IO.File]::ReadAllBytes($sf)) $Thumbprint
    if ($why) { return $why }
    $m = ConvertFrom-HMManifest ([System.Text.Encoding]::UTF8.GetString($man))
    if (-not $m.Count) { return 'Pruefsummen-Datei leer' }
    foreach ($k in $m.Keys) {
        $p = Join-Path $Dir ($k -replace '/', '\')
        if (-not (Test-Path -LiteralPath $p)) { return "Datei fehlt: $k" }
        if ((Get-HMFileSha256 $p) -ne $m[$k]) { return "Datei veraendert: $k" }
    }
    $root = (Resolve-Path -LiteralPath $Dir).Path.TrimEnd('\', '/')
    # Endung selbst filtern: -Include wird von Windows PowerShell 5.1 zusammen mit -LiteralPath ignoriert
    foreach ($f in @(Get-ChildItem -LiteralPath $Dir -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in '.ps1', '.psm1', '.psd1', '.xaml' })) {
        $relp = $f.FullName.Substring($root.Length + 1) -replace '\\', '/'
        if (-not $m.ContainsKey($relp)) { return "fremde Datei: $relp" }
    }
    return ''
}

# Wartende neue Fassungen uebernehmen (vor dem Laden, beim Start)
function Use-HUAddonPending([string]$AddonsDir) {
    $pend = Join-Path $AddonsDir '.pending'
    if (-not (Test-Path -LiteralPath $pend)) { return }
    foreach ($d in @(Get-ChildItem -LiteralPath $pend -Directory -ErrorAction SilentlyContinue)) {
        if (-not (Test-HUAddonId $d.Name)) { continue }
        $dst = Join-Path $AddonsDir $d.Name
        try {
            if (Test-Path -LiteralPath $dst) { Remove-Item -LiteralPath $dst -Recurse -Force -ErrorAction Stop }
            Move-Item -LiteralPath $d.FullName -Destination $dst -Force -ErrorAction Stop
        } catch { }
    }
}

# Module eines Ordners lesen. -Signed: jede Datei gegen die signierte Pruefsumme pruefen. -SkipDisabled: Ordner mit _ oder . am Anfang auslassen
function Get-HUAddonsInDir([string]$Dir, [Version]$AppVersion, [string]$Source, [switch]$Signed, [string]$Thumbprint = '', [switch]$SkipDisabled) {
    if (-not (Test-Path -LiteralPath $Dir)) { return @() }
    $out = @()
    foreach ($d in @(Get-ChildItem -LiteralPath $Dir -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike '.*' -and -not ($SkipDisabled -and $_.Name -like '_*') })) {
        $i = Read-HUAddonInfo $d.FullName
        $i.Source = $Source
        if (-not $i.Error -and $i.Id -cne $d.Name) { $i.Error = "Ordnername '$($d.Name)' passt nicht zur Kennung '$($i.Id)'" }
        if (-not $i.Error -and $Signed) { $why = Test-HUAddonFiles $d.FullName $Thumbprint; if ($why) { $i.Error = "nicht vertrauenswuerdig: $why" } }
        if (-not $i.Error -and $AppVersion -lt $i.MinApp) { $i.Error = "braucht Programmversion $($i.MinApp) oder neuer" }
        $out += $i
    }
    return $out
}

# Reihenfolge der Reiter: Feld order (klein zuerst), dann Titel. Gleiche Kennung: spaetere Quelle ersetzt fruehere.
# Eingabe in der Reihenfolge der Quellen (mitgeliefert, privat, Entwicklung).
function Merge-HUAddonList([object[]]$Lists) {
    $by = [ordered]@{}
    foreach ($i in @($Lists)) { if ($i -and -not $i.Error) { $by[$i.Id] = $i } }
    return @($by.Values | Sort-Object @{ Expression = { $_.Order } }, @{ Expression = { $_.Title } })
}

# Passt eine Datei zu den Endungen eines Moduls?
function Test-HUAddonDropMatch($Info, [string]$Path) {
    if (-not $Info -or -not @($Info.Drop).Count) { return $false }
    return (@($Info.Drop) -contains ([IO.Path]::GetExtension($Path).ToLowerInvariant()))
}

# --- Postfach: Dateien von einem Modul an ein anderes weitergeben ------------------------------
# Data\Postfach\<Ziel>\<guid>\<Datei> + eintrag.json { von, pc, benutzer, zeit, name }.
# Die Datei wird KOPIERT (das Original kann lokal auf einem anderen PC liegen oder spaeter verschoben werden).
# eintrag.json wird zuletzt geschrieben: ein Eintrag ohne eintrag.json ist unfertig und wird nicht gelesen.
# Das Postfach liegt im gemeinsamen Datenordner - das Zielmodul nimmt nur, was Get-HUAddonInboxItems liefert
# (Endung erlaubt, Datei liegt wirklich im Eintragsordner).
$script:HUInboxEntry = 'eintrag.json'
function Get-HUAddonInboxRoot([string]$DataRoot, [string]$Id) {
    if (-not (Test-HUAddonId $Id)) { throw "Ungueltige Kennung fuer das Postfach: '$Id'" }
    return (Join-Path (Join-Path $DataRoot 'Postfach') $Id)
}

# Eine Datei ins Postfach legen. Rueckgabe: Pfad des Eintragsordners. Wirft bei Fehlern (nichts Halbes bleibt liegen).
function Add-HUAddonInboxItem([string]$DataRoot, [string]$To, [string]$Path, [string]$From, [string]$User = "$env:USERNAME", [string]$Pc = "$env:COMPUTERNAME") {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Datei nicht gefunden: $Path" }
    if ($From -and -not (Test-HUAddonId $From)) { throw "Ungueltige Kennung des Absenders: '$From'" }
    $root = Get-HUAddonInboxRoot $DataRoot $To
    $dir = Join-Path $root ([guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    try {
        $name = [IO.Path]::GetFileName($Path)
        Copy-Item -LiteralPath $Path -Destination (Join-Path $dir $name) -Force -ErrorAction Stop
        $e = [pscustomobject][ordered]@{ von = "$From"; pc = "$Pc"; benutzer = "$User"; zeit = (Get-Date).ToString('s'); name = $name }
        $json = $e | ConvertTo-Json -Depth 3
        [IO.File]::WriteAllText((Join-Path $dir $script:HUInboxEntry), $json, (New-Object Text.UTF8Encoding($true)))
        return $dir
    } catch {
        Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
        throw
    }
}

# Eintraege im Postfach eines Moduls, aelteste zuerst. -Extensions: nur diese Endungen (wie "drop" des Zielmoduls).
function Get-HUAddonInboxItems([string]$DataRoot, [string]$Id, [string[]]$Extensions = @()) {
    $root = Get-HUAddonInboxRoot $DataRoot $Id
    if (-not (Test-Path -LiteralPath $root -PathType Container)) { return @() }
    $ext = @($Extensions | ForEach-Object { "$_".ToLowerInvariant() })
    $out = @()
    foreach ($d in @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^[0-9a-f]{32}$' })) {
        if ($d.Attributes -band [IO.FileAttributes]::ReparsePoint) { continue }
        $ef = Join-Path $d.FullName $script:HUInboxEntry
        if (-not (Test-Path -LiteralPath $ef -PathType Leaf)) { continue }
        try { $j = [IO.File]::ReadAllText($ef, [Text.Encoding]::UTF8) | ConvertFrom-Json -ErrorAction Stop } catch { continue }
        $name = "$($j.name)"
        # nur ein reiner Dateiname, der wirklich im Eintragsordner liegt
        if (-not $name -or $name -ne [IO.Path]::GetFileName($name) -or $name -eq $script:HUInboxEntry) { continue }
        $f = Join-Path $d.FullName $name
        if (-not (Test-Path -LiteralPath $f -PathType Leaf)) { continue }
        # keine Verknuepfungen (Datei-Symlink wuerde dem Zielmodul eine fremde Datei uebergeben)
        try { if ((Get-Item -LiteralPath $f -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) { continue } } catch { continue }
        if ($ext.Count -and $ext -notcontains [IO.Path]::GetExtension($name).ToLowerInvariant()) { continue }
        $t = $null; try { $t = [datetime]::Parse("$($j.zeit)", [Globalization.CultureInfo]::InvariantCulture) } catch { $t = $d.CreationTime }
        $out += [pscustomobject]@{ Key = $d.Name; Path = $f; Name = $name; From = "$($j.von)"; Pc = "$($j.pc)"; User = "$($j.benutzer)"; Time = $t; Dir = $d.FullName; Inbox = $Id }
    }
    return @($out | Sort-Object Time)
}

# Eintrag samt Kopie entfernen. Nur Ordner direkt unter dem Postfach des Moduls.
function Remove-HUAddonInboxEntry([string]$DataRoot, $Item) {
    if (-not $Item -or "$($Item.Key)" -notmatch '^[0-9a-f]{32}$') { throw 'Ungueltiger Postfach-Eintrag' }
    $dir = Join-Path (Get-HUAddonInboxRoot $DataRoot "$($Item.Inbox)") "$($Item.Key)"
    if (Test-Path -LiteralPath $dir) { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction Stop }
}

# --- Installieren / Aktualisieren --------------------------------------------------------------
# Neueste signierte Fassung (Kanal wie das Programm) laden, pruefen und unter Addons\.pending\<Kennung> ablegen.
# Rueckgabe: Status (Neu, Aktuell, Fehler), Text, Id, Version.
function Install-HUAddon([string]$Owner, [string]$Repo, [string]$Token, [string]$AddonsDir, [Version]$AppVersion, [string]$Channel = 'Stable', [string]$Thumbprint) {
    $res = [pscustomobject]@{ Status = 'Fehler'; Text = ''; Id = ''; Version = $null }
    try {
        $rel = Select-HMRelease (Get-HMReleases $Owner $Repo $Token) $Channel -SignedOnly
        if (-not $rel) { $res.Text = 'kein signiertes Release gefunden'; return $res }
        $man = Get-HMReleaseAsset $rel.ManifestUrl $rel.ManifestApi $Token
        $sig = Get-HMReleaseAsset $rel.SignatureUrl $rel.SignatureApi $Token
        $why = Test-HMManifestSignature $man $sig $Thumbprint
        if ($why) { $res.Text = "Signatur: $why"; return $res }
        $m = ConvertFrom-HMManifest ([System.Text.Encoding]::UTF8.GetString($man))
        if (-not $m.ContainsKey('addon.json')) { $res.Text = 'kein Modul (addon.json fehlt im Release)'; return $res }
        # Schneller Weg (ab 1.0.5): erst nur addon.json holen (Pruefsumme laut signierter Liste). Ist diese Fassung schon
        # installiert oder wartet sie schon auf den Neustart, wird nichts weiter geladen. Vorher wurden bei jedem Start alle
        # Dateien jedes Moduls geladen und erst danach verglichen: die Meldung kam deshalb erst nach Minuten.
        $probe = Join-Path ([IO.Path]::GetTempPath()) ('hu-addon-' + [guid]::NewGuid().ToString('N') + '.json')
        try {
            $hp = @{ Accept = 'application/vnd.github.v3.raw'; 'User-Agent' = $script:HMUserAgent; Authorization = "token $Token" }
            Invoke-WebRequest "https://api.github.com/repos/$Owner/$Repo/contents/addon.json?ref=$($rel.Tag)" -Headers $hp -UseBasicParsing -TimeoutSec 60 -OutFile $probe -ErrorAction Stop
            if ((Get-HMFileSha256 $probe) -eq $m['addon.json']) {
                $pj = [IO.File]::ReadAllText($probe, [Text.Encoding]::UTF8) | ConvertFrom-Json
                $aid = "$($pj.id)"; $aver = ConvertTo-HMVersion "$($pj.version)"; $atitle = $(if ("$($pj.title)") { "$($pj.title)" } else { $aid })
                if ($aid -match '^[A-Za-z0-9_-]+$' -and $aver) {
                    $cur0 = Read-HUAddonInfo (Join-Path $AddonsDir $aid)
                    if (-not $cur0.Error -and $cur0.Version -ge $aver) { $res.Id = $aid; $res.Version = $aver; $res.Status = 'Aktuell'; $res.Text = "$atitle $aver ist aktuell"; return $res }
                    $pen0 = Read-HUAddonInfo (Join-Path (Join-Path $AddonsDir '.pending') $aid)
                    if (-not $pen0.Error -and $pen0.Version -ge $aver) { $res.Id = $aid; $res.Version = $aver; $res.Status = 'Aktuell'; $res.Text = "$atitle $aver ist geladen und wartet auf den Neustart"; return $res }
                }
            }
        } catch { }   # schneller Weg geht nicht: wie bisher alles laden und dann vergleichen
        finally { if (Test-Path -LiteralPath $probe) { Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue } }
        $stage = Join-Path (Join-Path $AddonsDir '.staging') ([guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $stage -Force | Out-Null
        try {
            $h = @{ Accept = 'application/vnd.github.v3.raw'; 'User-Agent' = $script:HMUserAgent; Authorization = "token $Token" }
            foreach ($k in $m.Keys) {
                if ($k -match '(^[\\/])|(\.\.)|:' -or $k -notmatch '^[A-Za-z0-9._/ -]+$') { throw "unzulaessiger Pfad im Release: $k" }
                $dst = Join-Path $stage ($k -replace '/', '\')
                $dd = Split-Path $dst -Parent; if (-not (Test-Path -LiteralPath $dd)) { New-Item -ItemType Directory -Path $dd -Force | Out-Null }
                $enc = ($k -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
                Invoke-WebRequest "https://api.github.com/repos/$Owner/$Repo/contents/${enc}?ref=$($rel.Tag)" -Headers $h -UseBasicParsing -TimeoutSec 60 -OutFile $dst -ErrorAction Stop
                if ((Get-HMFileSha256 $dst) -ne $m[$k]) { throw "Pruefsumme stimmt nicht: $k" }
            }
            [System.IO.File]::WriteAllBytes((Join-Path $stage $script:HUAddonManifest), $man)
            [System.IO.File]::WriteAllBytes((Join-Path $stage $script:HUAddonSignature), $sig)
            $i = Read-HUAddonInfo $stage
            if ($i.Error) { throw $i.Error }
            if ($AppVersion -lt $i.MinApp) { throw "braucht Programmversion $($i.MinApp) oder neuer" }
            $why = Test-HUAddonFiles $stage $Thumbprint
            if ($why) { throw $why }
            $res.Id = $i.Id; $res.Version = $i.Version
            $cur = Read-HUAddonInfo (Join-Path $AddonsDir $i.Id)
            if (-not $cur.Error -and $cur.Version -ge $i.Version) { $res.Status = 'Aktuell'; $res.Text = "$($i.Title) $($i.Version) ist aktuell"; return $res }
            $pend = Join-Path (Join-Path $AddonsDir '.pending') $i.Id
            if (Test-Path -LiteralPath $pend) { Remove-Item -LiteralPath $pend -Recurse -Force }
            New-Item -ItemType Directory -Path (Split-Path $pend -Parent) -Force | Out-Null
            Move-Item -LiteralPath $stage -Destination $pend -Force
            $res.Status = 'Neu'; $res.Text = "$($i.Title) $($i.Version) geladen - wirkt nach dem Neustart"
            return $res
        } finally { if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue } }
    } catch {
        $code = 0; try { $code = [int]$_.Exception.Response.StatusCode } catch { }
        $res.Text = $(if ($code -in 401, 403, 404) { "kein Zugriff auf $Owner/$Repo (HTTP $code) - Token pruefen" } else { "$($_.Exception.Message)" })
        return $res
    }
}
