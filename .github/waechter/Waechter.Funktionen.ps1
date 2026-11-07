#Requires -Version 5.1
<#
.SYNOPSIS
    Funktionen des Wächters für externe Werkzeuge und Bibliotheken (neuere Versionen, bekannte Sicherheitslücken, Ende der Pflege).
.DESCRIPTION
    Zielmaschine: GitHub-Runner (ubuntu-latest mit pwsh) oder Entwicklungs-PC. Läuft unter Windows PowerShell 5.1 und PowerShell 7.
    Liegt nur im Kern (.github\waechter). Ein Modul liefert nur .github\waechter.json und ruft den Workflow des Kerns auf
    (Vorlage: .github\waechter\Vorlage-Modul.yml). Nichts davon wird mit der App oder einem Modul ausgeliefert.
    Quellen:
      - neueste Version: GitHub-Releases (ohne Vorabversionen), GitHub-Commits (feste Stände), PyPI, NuGet, npm, endoflife.date
      - Sicherheitslücken: osv.dev (PyPI, NuGet, npm) und die NVD (CPE-Name mit Version, für Programme ohne Paketquelle)
    Der Wächter ändert nichts. Treffer der NVD können ungenau sein (CPE-Bereiche), deshalb immer von Hand prüfen.
#>

[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# Versionen als Zahlenfolge: "10.08.0" -> 10,8,0; "v5.4.0.20240606" -> 5,4,0,20240606
function ConvertTo-HUWVersion([string]$Text) {
    return , @([regex]::Matches("$Text", '\d+') | ForEach-Object { [long]$_.Value })
}

# -1 = A älter, 0 = gleich, 1 = A neuer. Fehlende Stellen zählen als 0.
function Compare-HUWVersion([string]$A, [string]$B) {
    $x = ConvertTo-HUWVersion $A; $y = ConvertTo-HUWVersion $B
    $n = [Math]::Max($x.Count, $y.Count)
    for ($i = 0; $i -lt $n; $i++) {
        $p = 0; $q = 0
        if ($i -lt $x.Count) { $p = $x[$i] }
        if ($i -lt $y.Count) { $q = $y[$i] }
        if ($p -lt $q) { return -1 }
        if ($p -gt $q) { return 1 }
    }
    return 0
}

# Tag eines Release in eine Version umsetzen. Ohne Muster: führendes "v" weg.
# Mit Muster: Gruppen in das Format einsetzen, {n} als Zahl ohne führende Nullen, {n:00} zweistellig.
# Beispiel Ghostscript: Tag gs10080, Muster ^gs(\d{2})(\d{2})(\d)$, Format {1}.{2:00}.{3} -> 10.08.0
function ConvertFrom-HUWTag([string]$Tag, [string]$Muster = '', [string]$Format = '') {
    if (-not $Muster) { return ($Tag -replace '^[vV]', '') }
    $m = [regex]::Match($Tag, $Muster)
    if (-not $m.Success) { return '' }
    if (-not $Format) { return $m.Groups[1].Value }
    $g = $m.Groups
    return [regex]::Replace($Format, '\{(\d+)(?::(0+))?\}', {
            param($x)
            $i = [int]$x.Groups[1].Value
            if ($i -ge $g.Count) { return '' }
            $w = $g[$i].Value
            if ($w -notmatch '^\d+$') { return $w }
            if ($x.Groups[2].Success) { return ([long]$w).ToString($x.Groups[2].Value) }
            return "$([long]$w)"
        })
}

# Version für die CPE-Abfrage kürzen, z. B. Tesseract 5.4.0.20240606 -> 5.4.0
function Get-HUWCpeVersion([string]$Version, [int]$Teile = 0) {
    if ($Teile -le 0) { return $Version }
    return ((@($Version -split '\.') | Select-Object -First $Teile) -join '.')
}

function Get-HUWHeaders([string]$Url) {
    $h = @{ 'User-Agent' = 'HU-Waechter' }
    $tok = $env:GITHUB_TOKEN; if (-not $tok) { $tok = $env:GH_TOKEN }
    if ($tok -and $Url -like 'https://api.github.com/*') { $h['Authorization'] = "Bearer $tok" }
    if ($env:NVD_API_KEY -and $Url -like 'https://services.nvd.nist.gov/*') { $h['apiKey'] = $env:NVD_API_KEY }
    return $h
}

function Invoke-HUWGet([string]$Url) {
    return Invoke-RestMethod -Uri $Url -Headers (Get-HUWHeaders $Url) -TimeoutSec 60
}

# neueste Version laut Quelle. Rückgabe: @{ Version; Zweig; Eol; NeuesterZweig; Hinweis }
function Get-HUWNeueste($Quelle, [string]$Fest = '') {
    $r = @{ Version = ''; Eol = ''; NeuesterZweig = ''; Hinweis = ''; Neuer = $null }
    switch ("$($Quelle.art)") {
        'github' {
            $j = Invoke-HUWGet "https://api.github.com/repos/$($Quelle.repo)/releases/latest"
            $r.Version = ConvertFrom-HUWTag "$($j.tag_name)" "$($Quelle.tag)" "$($Quelle.format)"
            if (-not $r.Version) { throw "Tag $($j.tag_name) passt nicht zum Muster" }
        }
        'githubCommit' {
            # letzte Änderung der Datei; "Version" ist der feste Commit im Paket. Neuer nur, wenn diese Änderung
            # nach dem festen Commit kam (Vergleich ahead/identical = Datei unverändert).
            $p = ''; if ($Quelle.pfad) { $p = "&path=$([uri]::EscapeDataString("$($Quelle.pfad)"))" }
            $j = Invoke-HUWGet "https://api.github.com/repos/$($Quelle.repo)/commits?per_page=1$p"
            $r.Version = "$(@($j)[0].sha)".Substring(0, 12)
            if ($Fest) {
                $c = Invoke-HUWGet "https://api.github.com/repos/$($Quelle.repo)/compare/$($r.Version)...$Fest"
                $r.Neuer = ("$($c.status)" -notin 'ahead', 'identical')
                if (-not $r.Neuer) { $r.Version = $Fest }
            }
        }
        'pypi' { $r.Version = "$((Invoke-HUWGet "https://pypi.org/pypi/$($Quelle.paket)/json").info.version)" }
        'nuget' {
            $j = Invoke-HUWGet "https://api.nuget.org/v3-flatcontainer/$("$($Quelle.paket)".ToLowerInvariant())/index.json"
            $r.Version = "$(@($j.versions | Where-Object { $_ -notmatch '-' })[-1])"
        }
        'npm' { $r.Version = "$((Invoke-HUWGet "https://registry.npmjs.org/$($Quelle.paket)/latest").version)" }
        'endoflife' {
            $j = Invoke-HUWGet "https://endoflife.date/api/$($Quelle.produkt)/$($Quelle.zweig).json"
            $r.Version = "$($j.latest)"; $r.Eol = "$($j.eol)"
            # erst in eine Variable: Windows PowerShell 5.1 gibt das Array als ein Objekt aus
            $alle = Invoke-HUWGet "https://endoflife.date/api/$($Quelle.produkt).json"
            $alle = @($alle)
            if ($alle.Count) { $r.NeuesterZweig = "$($alle[0].cycle)" }
        }
        default { throw "unbekannte Quelle: $($Quelle.art)" }
    }
    return $r
}

# Sicherheitslücken laut NVD für einen CPE-Namen mit Version. Rückgabe: Liste @{ Id; Wert; Stufe }
function Get-HUWNvd([string]$Cpe, [string]$Version) {
    $u = "https://services.nvd.nist.gov/rest/json/cves/2.0?virtualMatchString=$([uri]::EscapeDataString("${Cpe}:$Version"))&resultsPerPage=200"
    $j = Invoke-HUWGet $u
    $out = @()
    foreach ($v in @($j.vulnerabilities)) {
        if (-not $v) { continue }
        $m = $null
        foreach ($k in 'cvssMetricV40', 'cvssMetricV31', 'cvssMetricV30', 'cvssMetricV2') { if ($v.cve.metrics.$k) { $m = @($v.cve.metrics.$k)[0].cvssData; break } }
        $w = ''; $s = ''
        if ($m) { $w = "$($m.baseScore)"; $s = "$($m.baseSeverity)" }
        $out += @{ Id = "$($v.cve.id)"; Wert = $w; Stufe = $s }
    }
    return , $out
}

# Sicherheitslücken laut osv.dev für viele Pakete auf einmal. $Pakete: Liste @{ Oekosystem; Name; Version }
# Rückgabe: Hashtable "Oekosystem:name" -> Liste @{ Id; Aliase; Behoben }
function Get-HUWOsv($Pakete) {
    $res = @{}
    $liste = @($Pakete)
    if (-not $liste.Count) { return $res }
    $q = @{ queries = @($liste | ForEach-Object { @{ package = @{ name = $_.Name; ecosystem = $_.Oekosystem }; version = $_.Version } }) }
    $j = Invoke-RestMethod -Method Post -Uri 'https://api.osv.dev/v1/querybatch' -Body ($q | ConvertTo-Json -Depth 6 -Compress) -ContentType 'application/json' -TimeoutSec 60
    $erg = @($j.results)
    for ($i = 0; $i -lt $liste.Count; $i++) {
        $key = "$($liste[$i].Oekosystem):$("$($liste[$i].Name)".ToLowerInvariant())"
        $ids = @()
        # Achtung: @($null).Count ist 1, deshalb vorher auf vulns prüfen
        if ($i -lt $erg.Count -and $erg[$i].vulns) { $ids = @($erg[$i].vulns | ForEach-Object { "$($_.id)" }) }
        $res[$key] = @()
        foreach ($id in $ids) {
            $d = Invoke-HUWGet "https://api.osv.dev/v1/vulns/$id"
            $fix = @($d.affected | ForEach-Object { $_.ranges } | ForEach-Object { $_.events } | Where-Object { $_.fixed } | ForEach-Object { "$($_.fixed)" } | Select-Object -Unique)
            $res[$key] += @{ Id = $id; Aliase = (@($d.aliases) -join ', '); Behoben = ($fix -join ', ') }
        }
    }
    return $res
}

# Gibt es die Datei für eine Version? (z. B. Windows-Paket von Python). HEAD-Anfrage, $true/$false.
function Test-HUWDownload([string]$Muster, [string]$Version) {
    $u = $Muster.Replace('{version}', $Version)
    try { $r = Invoke-WebRequest -Uri $u -Method Head -UseBasicParsing -TimeoutSec 60; return ([int]$r.StatusCode -lt 400) } catch { return $false }
}

# Python-Pakete aus requirements.txt: Liste @{ Name; Version }
function Read-HUWRequirements([string]$Pfad) {
    $out = @()
    foreach ($l in @(Get-Content -LiteralPath $Pfad -Encoding UTF8)) {
        $m = [regex]::Match($l, '^\s*([A-Za-z0-9_.\-]+)==([^\s;]+)')
        if ($m.Success) { $out += @{ Name = $m.Groups[1].Value; Version = $m.Groups[2].Value } }
    }
    return , $out
}

# Status eines Eintrags aus den Einzelergebnissen
function Get-HUWStatus($E) {
    if ($E.Fehler) { return 'fehler' }
    if (@($E.Luecken).Count) { return 'luecke' }
    if ($E.EolVorbei) { return 'eol' }
    if ($E.Neuer) { return 'neuer' }
    return 'aktuell'
}

# Signatur der Befunde: ändert sich nur, wenn ein Befund neu ist oder wegfällt (für die Teams-Nachricht)
function Get-HUWSignatur($Eintraege) {
    $teile = @()
    foreach ($e in @($Eintraege)) {
        if ($e.Status -eq 'aktuell') { continue }
        $teile += "$($e.Id)|$($e.Status)|$($e.Neueste)|$((@($e.Luecken) | ForEach-Object { $_.Id } | Sort-Object) -join ',')"
    }
    $txt = (@($teile | Sort-Object) -join "`n")
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $b = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($txt)) } finally { $sha.Dispose() }
    return (($b | ForEach-Object { $_.ToString('x2') }) -join '').Substring(0, 16)
}

# Teams-Nachricht (Adaptive Card wie in HUMig). Wirft nie, Rückgabe '' = gesendet, sonst Grund.
function Send-HUWTeams([string]$Url, [string]$Titel, [string[]]$Zeilen, [string]$Stufe = 'Warning') {
    if (-not $Url) { return 'kein Webhook' }
    try {
        $farbe = @{ OK = 'Good'; Warning = 'Warning'; Error = 'Attention'; Info = 'Accent' }[$Stufe]
        $body = New-Object System.Collections.Generic.List[object]
        $body.Add([ordered]@{ type = 'TextBlock'; text = $Titel; weight = 'Bolder'; size = 'Medium'; wrap = $true; color = $farbe })
        foreach ($l in @($Zeilen | Where-Object { "$_".Trim() })) { $body.Add([ordered]@{ type = 'TextBlock'; text = "$l"; wrap = $true; spacing = 'Small' }) }
        $body.Add([ordered]@{ type = 'TextBlock'; text = "Wächter · $((Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm')) UTC"; isSubtle = $true; size = 'Small'; wrap = $true; spacing = 'Medium' })
        $card = [ordered]@{ '$schema' = 'http://adaptivecards.io/schemas/adaptive-card.json'; type = 'AdaptiveCard'; version = '1.4'; body = $body.ToArray() }
        $msg = [ordered]@{ type = 'message'; attachments = @([ordered]@{ contentType = 'application/vnd.microsoft.card.adaptive'; contentUrl = $null; content = $card }) }
        $json = $msg | ConvertTo-Json -Depth 10
        Invoke-RestMethod -Method Post -Uri $Url -Body ([Text.Encoding]::UTF8.GetBytes($json)) -ContentType 'application/json; charset=utf-8' -TimeoutSec 30 | Out-Null
        return ''
    } catch { return "$($_.Exception.Message)" }
}

# Ganze Prüfung. $Wurzel = Modulordner, $Konfig = waechter.json als Objekt.
# $NvdPause: Sekunden zwischen NVD-Abfragen (ohne Schlüssel erlaubt die NVD 5 Abfragen in 30 Sekunden).
function Invoke-HUWaechter([string]$Wurzel, $Konfig, [int]$NvdPause = 7) {
    $paket = $null; $versionen = @{}
    if ($Konfig.werkzeugpaket) {
        $paket = Get-Content -LiteralPath (Join-Path $Wurzel "$($Konfig.werkzeugpaket)") -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($d in @($paket.dateien)) { $versionen["$($d.id)"] = "$($d.version)" }
    }
    $req = @()
    if ($Konfig.pythonPakete) { $req = Read-HUWRequirements (Join-Path $Wurzel "$($Konfig.pythonPakete)") }
    $reqV = @{}; foreach ($p in $req) { $reqV["$($p.Name)".ToLowerInvariant()] = $p.Version }

    $eintraege = @(); $osvFragen = @()
    $ersteNvd = $true
    foreach ($k in @($Konfig.eintraege)) {
        $e = @{ Id = "$($k.id)"; Version = "$($k.version)"; Neueste = ''; Neuer = $false; Luecken = @(); EolVorbei = $false; Hinweise = @(); Fehler = ''; Status = '' }
        if (-not $e.Version -and $versionen.ContainsKey($e.Id)) { $e.Version = $versionen[$e.Id] }
        if (-not $e.Version -and $k.pythonPaket -and $reqV.ContainsKey("$($k.pythonPaket)".ToLowerInvariant())) { $e.Version = $reqV["$($k.pythonPaket)".ToLowerInvariant()] }
        if (-not $e.Version) { $e.Fehler = 'keine Version gefunden'; $e.Status = Get-HUWStatus $e; $eintraege += $e; continue }
        try {
            if ($k.neueste) {
                $n = Get-HUWNeueste $k.neueste $e.Version
                $e.Neueste = $n.Version
                if ($null -ne $n.Neuer) { $e.Neuer = [bool]$n.Neuer }
                elseif ($n.Version) { $e.Neuer = ((Compare-HUWVersion $n.Version $e.Version) -gt 0) }
                if ($n.Eol) {
                    $eol = [datetime]::ParseExact($n.Eol, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
                    if ($eol -lt (Get-Date)) { $e.EolVorbei = $true; $e.Hinweise += "Pflege endete am $($n.Eol)" }
                    elseif ($eol -lt (Get-Date).AddDays(365)) { $e.Hinweise += "Pflege endet am $($n.Eol)" }
                }
                if ($n.NeuesterZweig -and $k.neueste.zweig -and (Compare-HUWVersion $n.NeuesterZweig "$($k.neueste.zweig)") -gt 0) { $e.Hinweise += "neuester Zweig: $($n.NeuesterZweig)" }
                if ($e.Neuer -and $k.download -and -not (Test-HUWDownload "$($k.download)" $e.Neueste)) { $e.Hinweise += "für $($e.Neueste) gibt es kein fertiges Paket, nur einen neueren Zweig" }
            }
            if ($k.cpe) {
                if (-not $ersteNvd) { Start-Sleep -Seconds $NvdPause }
                $ersteNvd = $false
                $teile = 0; if ($k.cpeTeile) { $teile = [int]$k.cpeTeile }
                $e.Luecken = Get-HUWNvd "$($k.cpe)" (Get-HUWCpeVersion $e.Version $teile)
                if (@($e.Luecken).Count) { $e.Hinweise += 'Treffer der NVD, von Hand prüfen (CPE-Bereiche können ungenau sein)' }
            }
            if ($k.osv) { $osvFragen += @{ Oekosystem = "$($k.osv.oekosystem)"; Name = "$($k.osv.name)"; Version = $e.Version; Ziel = $e } }
        } catch { $e.Fehler = "$($_.Exception.Message)" }
        $eintraege += $e
    }

    # alle Python-Pakete aus requirements.txt gegen osv.dev, dazu Einträge mit osv-Angabe
    foreach ($p in $req) { $osvFragen += @{ Oekosystem = 'PyPI'; Name = $p.Name; Version = $p.Version; Ziel = $null } }
    $pyLuecken = @()
    try {
        $osv = Get-HUWOsv $osvFragen
        foreach ($f in $osvFragen) {
            $l = @($osv["$($f.Oekosystem):$("$($f.Name)".ToLowerInvariant())"])
            $l = @($l | Where-Object { $_ })
            if (-not $l.Count) { continue }
            if ($f.Ziel) { $f.Ziel.Luecken = @($f.Ziel.Luecken) + $l }
            else { $pyLuecken += @{ Id = "$($f.Name)"; Version = $f.Version; Neueste = ''; Neuer = $false; Luecken = $l; EolVorbei = $false; Hinweise = @('Python-Paket aus ' + "$($Konfig.pythonPakete)"); Fehler = ''; Status = '' } }
        }
    } catch { $eintraege += @{ Id = 'osv.dev'; Version = ''; Neueste = ''; Neuer = $false; Luecken = @(); EolVorbei = $false; Hinweise = @(); Fehler = "$($_.Exception.Message)"; Status = '' } }
    $eintraege += $pyLuecken

    # geprüfte Fehlalarme aus waechter.json ("ignorieren": [{ "id": "CVE-…", "grund": "…" }]) herausnehmen, im Bericht vermerken
    $ign = @{}
    foreach ($k in @($Konfig.eintraege)) { foreach ($x in @($k.ignorieren | Where-Object { $_ })) { $ign["$($k.id)|$($x.id)"] = "$($x.grund)" } }
    foreach ($e in $eintraege) {
        $weg = @(@($e.Luecken) | Where-Object { $_ -and $ign.ContainsKey("$($e.Id)|$($_.Id)") })
        if ($weg.Count) {
            $e.Luecken = @(@($e.Luecken) | Where-Object { $_ -and -not $ign.ContainsKey("$($e.Id)|$($_.Id)") })
            $e.Hinweise += "$($weg.Count) geprüfte Fehlalarme ausgenommen ($((@($weg) | ForEach-Object { $_.Id }) -join ', '))"
            if (-not @($e.Luecken).Count) { $e.Hinweise = @($e.Hinweise | Where-Object { $_ -notlike 'Treffer der NVD*' }) }
        }
    }
    foreach ($e in $eintraege) { $e.Status = Get-HUWStatus $e }

    $befunde = @($eintraege | Where-Object { $_.Status -ne 'aktuell' })
    return [ordered]@{
        modul        = "$($Konfig.modul)"
        paket        = $(if ($paket) { "$($paket.paket)" } else { '' })
        stand        = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        pythonPakete = @($req).Count
        befunde      = $befunde.Count
        signatur     = (Get-HUWSignatur $eintraege)
        eintraege    = @($eintraege | ForEach-Object {
                [ordered]@{ id = $_.Id; version = $_.Version; neueste = $_.Neueste; status = $_.Status
                    luecken = @(@($_.Luecken) | ForEach-Object { [ordered]@{ id = $_.Id; wert = "$($_.Wert)"; stufe = "$($_.Stufe)"; aliase = "$($_.Aliase)"; behoben = "$($_.Behoben)" } })
                    hinweise = @($_.Hinweise); fehler = $_.Fehler }
            })
    }
}

# waechter.json eines Moduls prüfen: gültig, jede Quelle bekannt, jedes Werkzeug im Werkzeugpaket abgedeckt.
# $Wurzel = Modulordner. Rückgabe: die Probleme einzeln (nichts = in Ordnung), Aufruf mit @(...). Ohne Netz.
function Test-HUWKonfig([string]$Wurzel) {
    $f = Join-Path (Join-Path $Wurzel '.github') 'waechter.json'
    if (-not (Test-Path -LiteralPath $f)) { return '.github/waechter.json fehlt' }
    try { $k = Get-Content -LiteralPath $f -Raw -Encoding UTF8 | ConvertFrom-Json } catch { return "waechter.json ungültig: $($_.Exception.Message)" }
    $fehlt = @()
    if (-not "$($k.modul)".Trim()) { $fehlt += 'modul fehlt' }
    $ids = @($k.eintraege | ForEach-Object { "$($_.id)" })
    if (-not $ids.Count) { $fehlt += 'keine Einträge' }
    foreach ($q in @($k.eintraege | Where-Object { $_ })) {
        if (-not "$($q.id)".Trim()) { $fehlt += 'Eintrag ohne id'; continue }
        if (-not $q.neueste -and -not $q.cpe -and -not $q.osv) { $fehlt += "$($q.id): keine Quelle" }
        if ($q.neueste -and "$($q.neueste.art)" -notin 'github', 'githubCommit', 'pypi', 'nuget', 'npm', 'endoflife') { $fehlt += "$($q.id): Quelle $($q.neueste.art)" }
        if ($q.cpe -and "$($q.cpe)" -notmatch '^cpe:2\.3:[aho]:[^:]+:[^:]+$') { $fehlt += "$($q.id): CPE ohne Hersteller/Produkt" }
        if ($q.osv -and (-not "$($q.osv.oekosystem)" -or -not "$($q.osv.name)")) { $fehlt += "$($q.id): osv ohne oekosystem oder name" }
        if (-not $q.version -and -not $k.werkzeugpaket -and -not $q.pythonPaket) { $fehlt += "$($q.id): keine Version" }
        foreach ($x in @($q.ignorieren | Where-Object { $_ })) { if ("$($x.id)" -notmatch '^(CVE|GHSA|PYSEC|OSV)-' -or -not "$($x.grund)".Trim()) { $fehlt += "$($q.id): Ausnahme ohne Kennung oder Grund" } }
    }
    foreach ($d in @($ids | Group-Object | Where-Object { $_.Count -gt 1 })) { $fehlt += "$($d.Name): doppelt" }
    if ($k.werkzeugpaket) {
        $wp = Join-Path $Wurzel "$($k.werkzeugpaket)"
        if (-not (Test-Path -LiteralPath $wp)) { $fehlt += "Werkzeugpaket $($k.werkzeugpaket) fehlt" }
        else {
            $w = Get-Content -LiteralPath $wp -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($d in @($w.dateien)) { if ($ids -notcontains "$($d.id)") { $fehlt += "$($d.id): kein Wächter-Eintrag" } }
        }
    }
    if ($k.pythonPakete -and -not (Test-Path -LiteralPath (Join-Path $Wurzel "$($k.pythonPakete)"))) { $fehlt += "Python-Pakete $($k.pythonPakete) fehlen" }
    return $fehlt
}

# Bericht als Markdown (Issue und Zusammenfassung der CI)
function ConvertTo-HUWBericht($Ergebnis) {
    $sym = @{ aktuell = '✅'; neuer = '🆕'; luecke = '🔴'; eol = '⛔'; fehler = '⚠️' }
    $txt = @{ aktuell = 'aktuell'; neuer = 'neuere Version'; luecke = 'Sicherheitslücke'; eol = 'Pflege beendet'; fehler = 'Prüfung gescheitert' }
    $z = @("## Wächter: $($Ergebnis.modul)", '')
    $kopf = "Stand $($Ergebnis.stand)"
    if ($Ergebnis.paket) { $kopf += ", Werkzeugpaket $($Ergebnis.paket)" }
    $kopf += ", $($Ergebnis.pythonPakete) Python-Pakete gegen osv.dev geprüft."
    $z += $kopf; $z += ''
    if (-not $Ergebnis.befunde) { $z += 'Keine Befunde.' } else { $z += "**$($Ergebnis.befunde) Befund(e).** Der Wächter ändert nichts, ein neues Werkzeugpaket wird gebaut und getestet wie bisher." }
    $z += ''
    $z += '| Werkzeug | Version | neueste | Stand | Hinweise |'
    $z += '|---|---|---|---|---|'
    foreach ($e in @($Ergebnis.eintraege)) {
        $h = @($e.hinweise)
        if ($e.fehler) { $h += "Fehler: $($e.fehler)" }
        # schwerste Lücken zuerst
        $l = @(@($e.luecken) | Sort-Object -Property @{ Expression = { $w = 0.0; [void][double]::TryParse("$($_.wert)", [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$w); $w }; Descending = $true })
        if ($l.Count) {
            $ids = @($l | Select-Object -First 8 | ForEach-Object { $s = $_.id; if ($_.wert) { $s += " ($($_.wert))" }; if ($_.behoben) { $s += " behoben in $($_.behoben)" }; $s })
            $mehr = ''; if ($l.Count -gt 8) { $mehr = " und $($l.Count - 8) weitere" }
            $h = @("$($l.Count) Lücke(n): $($ids -join ', ')$mehr") + $h
        }
        $z += "| $($e.id) | $($e.version) | $($e.neueste) | $($sym[$e.status]) $($txt[$e.status]) | $(($h -join '; ').Replace('|', '/')) |"
    }
    $z += ''
    $z += "<!-- huwaechter:$($Ergebnis.signatur) -->"
    return ($z -join "`n")
}
