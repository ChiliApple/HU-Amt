#Requires -Version 5.1
<#
.SYNOPSIS
    HU-Amt Pull - laedt eine HU-Amt-Version von GitHub (Release), prueft sie und startet das Tool neu.
.DESCRIPTION
    Zielordner (in dieser Reihenfolge):
      1. -Target angegeben                                 -> gewinnt immer
      2. Pull.ps1 liegt IN einer Installation              -> Update an Ort und Stelle
         (Main.ps1 oder Core\ daneben)
      3. Pull.ps1 liegt allein in einem eigenen Ordner     -> Erstinstallation in diesen Ordner
         (nicht in Desktop, Downloads, Dokumente oder Laufwerkswurzel)
      4. sonst (Erstinstallation)                          -> $env:USERPROFILE\Desktop\HU-Amt

    Welche Version:
      -Version 2.0.1      -> genau diese Version (auch aelter = "Vorversion")
      sonst Kanal         -> -Channel oder Config\update.json "Channel": Stable (freigegebene Releases, Standard) / Test (auch Vorab-Releases)
      -Branch / UseBranch -> Entwicklungsstand eines Branches statt Release (ohne Pruefsumme, nur ohne Signaturpflicht)

    Sicherheit (gleicher Ablauf in allen HU-Tools):
      - jedes Release enthaelt HU-Amt-files.sha256 (SHA256 aller Dateien, erstellt von den automatischen Tests auf GitHub)
      - zuerst werden ALLE Dateien geladen und geprueft, erst dann ersetzt - bei einer Abweichung bleibt alles unveraendert
      - Signatur: nur Releases mit gueltiger Signatur HU-Amt-files.sha256.p7s des Herausgeber-Zertifikats werden angenommen
        (offizielle Quelle: Fingerabdruck eingebaut; eigene Quelle: Config\update.json "SignerThumbprint").
        Abschalten nur bewusst mit Config\update.json "AllowUnsigned": true (Einstellungen > Update).

    -WaitPid <PID>: wird vom Update-Knopf im Tool uebergeben. Pull wartet, bis das Tool beendet ist (max. 60 s).

    Gemeinsamer Netzordner: Das Update ersetzt nichts, solange jemand anderer die App offen hat (Data\Laufend\*.lock
    laesst sich nicht loeschen). Dann kommt "Bitte zuerst schliessen: ...".

    Lokale Daten bleiben unangetastet: Data\ (alle Daten, Benutzer, Token), Addons\, AddonsDev\, bin\, Config\update.json,
    Config\installed.json, Config\powershell-host.txt, HU-Amt.exe.
    Token (nur fuer private Repos): Umgebungsvariable HU_GITHUB_TOKEN oder Data\Benutzer\<Name>\Token\GitHub.xml (DPAPI).

    Erstinstallation: Wer installiert, wird Verwalter (Data\Verwalter.txt mit dem Windows-Benutzernamen), aber nur,
    wenn es noch keine Installation und keine Verwalter-Liste gibt. Ein Update aendert die Liste nie.
.NOTES
    Erstinstallation: Pull.ps1 allein in einen neuen Ordner legen (z. B. auf der Netzfreigabe), dann
        Rechtsklick auf Pull.ps1 > "Mit PowerShell ausfuehren"
    oder mit vollem Pfad:
        powershell -ExecutionPolicy Bypass -File "\\Server\Freigabe\HU-Amt\Pull.ps1"
    Zielmaschine: der PC, auf dem HU-Amt laufen soll, oder der Netzordner (Windows PowerShell 5.1).
#>
param(
    [string]$Target,
    [int]$WaitPid = 0,
    [switch]$NoStart,
    [string]$Owner,
    [string]$Repo,
    [string]$Branch,
    [ValidateSet('', 'Stable', 'Test')][string]$Channel = '',
    [string]$Version = '',
    [switch]$NonInteractive
)
$ErrorActionPreference = 'Stop'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }

#region HMUpdateLib
# Gemeinsame Update-Funktionen - identisch in Pull.ps1 und Functions\Core-Update.ps1 (die automatischen Tests pruefen das)
$script:HMManifestName  = 'HU-Amt-files.sha256'
$script:HMSignatureName = 'HU-Amt-files.sha256.p7s'
$script:HMUserAgent     = 'HU-Amt'
# Offizielle Update-Quelle und Fingerabdruck des Zertifikats, mit dem ihre Releases signiert werden (oeffentlich, kein Geheimnis)
$script:HMDefaultOwner  = 'ChiliApple'
$script:HMDefaultRepo   = 'HU-Amt'
$script:HMDefaultSigner = '1B669AE240DA1A91043C4576763D9F8E0BF762FA'

function Get-HMDefaultSigner([string]$Owner, [string]$Repo) {
    if ($Owner -eq $script:HMDefaultOwner -and $Repo -eq $script:HMDefaultRepo) { return $script:HMDefaultSigner }
    return ''
}
# Update-Einstellungen aus Config\update.json (fehlt die Datei: offizielle Quelle, Kanal Stabil, nur signierte Releases)
#   Signaturpflicht gilt, sobald ein Fingerabdruck bekannt ist (offizielle Quelle: eingebaut) - ausser "AllowUnsigned": true.
#   Mit Signaturpflicht gibt es keinen Branch-Modus (ein Branch-Stand ist nicht signiert).
function Get-HMUpdateConfig([string]$ConfigDir) {
    $c = [ordered]@{ Owner = $script:HMDefaultOwner; Repo = $script:HMDefaultRepo; Branch = 'main'; UseBranch = $false; Channel = 'Stable'; AllowUnsigned = $false; SignerThumbprint = ''; DefaultSigner = $false; RequireSignature = $false }
    $f = Join-Path $ConfigDir 'update.json'
    if (Test-Path -LiteralPath $f) {
        try {
            $u = Get-Content -LiteralPath $f -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($k in @('Owner', 'Repo', 'Branch', 'SignerThumbprint')) { if ($u.PSObject.Properties[$k] -and "$($u.$k)".Trim()) { $c[$k] = "$($u.$k)".Trim() } }
            if ($u.PSObject.Properties['Channel'] -and "$($u.Channel)" -match '^(Stable|Test)$') { $c.Channel = "$($u.Channel)" }
            if ($u.PSObject.Properties['UseBranch']) { $c.UseBranch = ($u.UseBranch -eq $true) }
            if ($u.PSObject.Properties['AllowUnsigned']) { $c.AllowUnsigned = ($u.AllowUnsigned -eq $true) }
        } catch { }
    }
    $c.SignerThumbprint = ("$($c.SignerThumbprint)" -replace '[^0-9A-Fa-f]', '').ToUpper()
    if (-not $c.SignerThumbprint) {
        $d = Get-HMDefaultSigner $c.Owner $c.Repo
        if ($d) { $c.SignerThumbprint = $d; $c.DefaultSigner = $true }
    }
    $c.RequireSignature = ([bool]$c.SignerThumbprint -and -not $c.AllowUnsigned)
    if ($c.RequireSignature) { $c.UseBranch = $false }
    return [pscustomobject]$c
}
function ConvertTo-HMVersion([string]$Tag) {
    $v = "$Tag".Trim() -replace '^[vV]', ''
    $o = $null
    if ($v -match '^\d+(\.\d+){1,3}$' -and [Version]::TryParse($v, [ref]$o)) { return $o }
    return $null
}
# GitHub-Releases -> Liste (neueste zuerst): Version, Tag, Prerelease (= Kanal Test), Datum, Notizen, Pruefsummen-/Signatur-Datei
function ConvertTo-HMReleaseList($Raw) {
    $out = @()
    foreach ($r in @($Raw)) {
        if (-not $r -or $r.draft -eq $true) { continue }
        $ver = ConvertTo-HMVersion "$($r.tag_name)"
        if (-not $ver) { continue }
        $assets = @($r.assets | Where-Object { $_ })
        $man = @($assets | Where-Object { "$($_.name)" -eq $script:HMManifestName })[0]
        $sig = @($assets | Where-Object { "$($_.name)" -eq $script:HMSignatureName })[0]
        $d = ''
        try { $d = ([datetime]$r.published_at).ToLocalTime().ToString('dd.MM.yyyy HH:mm') } catch { $d = "$($r.published_at)" }
        $out += [pscustomobject]@{
            Version = $ver; Tag = "$($r.tag_name)"; Prerelease = ($r.prerelease -eq $true); Date = $d; Notes = "$($r.body)"
            ManifestUrl = $(if ($man) { "$($man.browser_download_url)" } else { '' }); ManifestApi = $(if ($man) { "$($man.url)" } else { '' })
            SignatureUrl = $(if ($sig) { "$($sig.browser_download_url)" } else { '' }); SignatureApi = $(if ($sig) { "$($sig.url)" } else { '' })
        }
    }
    return @($out | Sort-Object Version -Descending)
}
function Get-HMReleases([string]$Owner, [string]$Repo, [string]$Token) {
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
    $h = @{ Accept = 'application/vnd.github+json'; 'User-Agent' = $script:HMUserAgent }
    if ($Token) { $h.Authorization = "token $Token" }
    # GitHub liefert die Liste mit "Cache-Control: max-age=60" - Proxys/Zwischenspeicher wuerden z.B. eine gerade
    # angehaengte Signatur bis zu 1 Minute nicht zeigen. Darum: no-cache + eindeutige Adresse je Abfrage
    $h['Cache-Control'] = 'no-cache'
    $raw = Invoke-RestMethod "https://api.github.com/repos/$Owner/$Repo/releases?per_page=100&nocache=$([DateTime]::UtcNow.Ticks)" -Headers $h -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
    return @(ConvertTo-HMReleaseList $raw)
}
# Kanal Stabil = nur freigegebene Releases, Test = auch Vorab-Releases; jeweils die hoechste Version
# -SignedOnly: nur Releases mit Pruefsummen- UND Signatur-Datei (bei Signaturpflicht)
function Select-HMRelease($Releases, [string]$Channel, [switch]$SignedOnly) {
    $l = @($Releases | Where-Object { $_ -and $_.Version })
    if ($Channel -ne 'Test') { $l = @($l | Where-Object { -not $_.Prerelease }) }
    if ($SignedOnly) { $l = @($l | Where-Object { $_.ManifestUrl -and $_.SignatureUrl }) }
    return (@($l | Sort-Object Version -Descending) | Select-Object -First 1)
}
# Release-Datei (Pruefsummen/Signatur) als Bytes laden - mit Token ueber die API (private Repos).
# Ausweichweg: liefert der Download-Link von github.com einen Fehler (z.B. 503), wird dieselbe Datei ueber die API geladen
# (gleicher Inhalt; die Echtheit sichert ohnehin die Signatur/Pruefsumme)
function Get-HMReleaseAsset([string]$Url, [string]$ApiUrl, [string]$Token) {
    $tries = @()
    if ($Token -and $ApiUrl) { $tries += , @($ApiUrl, @{ Accept = 'application/octet-stream'; 'User-Agent' = $script:HMUserAgent; Authorization = "token $Token" }) }
    if ($Url) { $tries += , @($Url, @{ 'User-Agent' = $script:HMUserAgent }) }
    if ($ApiUrl) { $tries += , @($ApiUrl, @{ Accept = 'application/octet-stream'; 'User-Agent' = $script:HMUserAgent }) }
    $last = $null
    foreach ($t in $tries) {
        $tmp = [System.IO.Path]::GetTempFileName()
        try {
            Invoke-WebRequest $t[0] -Headers $t[1] -UseBasicParsing -TimeoutSec 60 -OutFile $tmp -ErrorAction Stop
            return ,([System.IO.File]::ReadAllBytes($tmp))
        } catch { $last = $_ }
        finally { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
    }
    if ($last) { throw $last }
    throw 'keine Download-Adresse'
}
# Pruefsummen-Datei (sha256sum-Format: "<hash>  <pfad>") -> Hashtable Pfad -> Hash (klein)
function ConvertFrom-HMManifest([string]$Text) {
    $m = @{}
    foreach ($ln in ("$Text" -split "`r?`n")) {
        if ($ln -match '^([0-9a-fA-F]{64}) [ \*](.+?)\s*$') { $m[$Matches[2]] = $Matches[1].ToLower() }
    }
    return $m
}
function Get-HMFileSha256([string]$Path) {
    $s = [System.Security.Cryptography.SHA256]::Create()
    $fs = [System.IO.File]::OpenRead($Path)
    try { return (-join ($s.ComputeHash($fs) | ForEach-Object { $_.ToString('x2') })) } finally { $fs.Dispose(); $s.Dispose() }
}
# Signatur (PKCS#7/CMS, abgetrennt) der Pruefsummen-Datei pruefen. Rueckgabe: '' = gueltig, sonst Grund.
# Geprueft werden die Signatur selbst und der Fingerabdruck des Zertifikats (fest eingestellt) - keine Online-Sperrlistenpruefung.
function Test-HMManifestSignature([byte[]]$Manifest, [byte[]]$Signature, [string]$Thumbprint) {
    $want = ("$Thumbprint" -replace '[^0-9A-Fa-f]', '').ToUpper()
    if (-not $want) { return 'kein Fingerabdruck (Thumbprint) des Signatur-Zertifikats eingestellt' }
    if (-not $Signature -or -not $Signature.Length) { return 'keine Signatur-Datei im Release' }
    try {
        try { Add-Type -AssemblyName System.Security -ErrorAction Stop } catch { }
        $ci = New-Object System.Security.Cryptography.Pkcs.ContentInfo -ArgumentList (, [byte[]]$Manifest)
        $cms = New-Object System.Security.Cryptography.Pkcs.SignedCms -ArgumentList $ci, $true
        $cms.Decode($Signature)
        $cms.CheckSignature($true)
        $tps = @(foreach ($si in $cms.SignerInfos) { if ($si.Certificate) { "$($si.Certificate.Thumbprint)".ToUpper() } })
        if (-not $tps.Count) { return 'Signatur ohne Zertifikat' }
        if ($tps -notcontains $want) { return "signiert mit einem anderen Zertifikat ($($tps -join ', '))" }
        return ''
    } catch { return "Signatur ungueltig: $($_.Exception.Message)" }
}
#endregion HMUpdateLib

function Stop-HMPull([string]$Msg) {
    Write-Host "[FEHLER] $Msg" -ForegroundColor Red
    if (-not $NonInteractive) { Read-Host 'Enter zum Beenden' | Out-Null }
    exit 1
}

# Liegt Pull.ps1 allein in einem Ordner, wird dorthin installiert (z. B. ein neuer Ordner auf der Netzfreigabe).
# Nicht in Sammelordnern wie Desktop, Downloads oder Dokumente des Benutzers, dort bleibt es bei Desktop\HU-Amt.
function Test-HMPullEmptyHome([string]$Dir, [string]$Self) {
    $d = "$Dir".TrimEnd('\')
    $sammel = @($env:USERPROFILE, (Join-Path $env:USERPROFILE 'Desktop'), (Join-Path $env:USERPROFILE 'Downloads'), (Join-Path $env:USERPROFILE 'Documents'))
    try { $sammel += [Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('MyDocuments') } catch { }
    foreach ($s in $sammel) { if ($s -and $d -ieq "$s".TrimEnd('\')) { return $false } }
    if ($d -match '^[A-Za-z]:$') { return $false }
    $rest = @(Get-ChildItem -LiteralPath $Dir -Force -ErrorAction SilentlyContinue | Where-Object { $_.FullName -ne $Self -and $_.Name -notin 'desktop.ini', 'Thumbs.db' })
    return ($rest.Count -eq 0)
}

# --- Zielordner bestimmen ---
if (-not $Target) {
    $here = $PSScriptRoot
    if ($here -and ((Test-Path (Join-Path $here 'Main.ps1')) -or (Test-Path (Join-Path $here 'Core')))) { $Target = $here; $mode = 'Update (an Ort und Stelle)' }
    elseif ($here -and (Test-HMPullEmptyHome $here $PSCommandPath)) { $Target = $here; $mode = 'Erstinstallation (dieser Ordner)' }
    else { $Target = Join-Path $env:USERPROFILE 'Desktop\HU-Amt'; $mode = 'Erstinstallation (Desktop)' }
} else { $mode = 'Ziel per -Target' }
$Target = $Target.TrimEnd('\')
# Erstinstallation = im Ziel liegt noch kein Programm (dann wird, wer installiert, Verwalter - siehe Schritt 7)
$isFirstInstall = -not (Test-Path -LiteralPath (Join-Path $Target 'Main.ps1'))

# --- Quelle, Kanal, Signaturpflicht ---
$cfgDir = Join-Path $Target 'Config'
# abgebrochenes Update (Journal vorhanden): bisherige Dateien aus *.pullold zuruecksetzen
$jr = Join-Path $cfgDir 'pull-journal.json'
if (Test-Path -LiteralPath $jr) {
    Write-Host '[WARN] Letztes Update wurde abgebrochen - stelle bisherige Dateien wieder her...' -ForegroundColor Yellow
    # Erst zuweisen, dann durchlaufen: PS 5.1 gibt ein JSON-Array aus ConvertFrom-Json als EIN Objekt aus -
    # "foreach ($x in @(... | ConvertFrom-Json))" liefe nur einmal mit allen Pfaden zusammen und stellte nichts zurueck
    $jl = $null
    try { $jl = Get-Content -LiteralPath $jr -Raw -Encoding UTF8 | ConvertFrom-Json }
    catch { Stop-HMPull "Journal $jr nicht lesbar ($($_.Exception.Message)) - bitte *.pullold im Programmordner pruefen." }
    $jbad = 0
    foreach ($jp in @($jl)) {
        if (-not "$jp") { continue }
        $jo = "$jp.pullold"
        if (Test-Path -LiteralPath $jo) {
            try { Move-Item -LiteralPath $jo -Destination "$jp" -Force -ErrorAction Stop; Write-Host "  zurueckgestellt: $jp" -ForegroundColor DarkGray }
            catch { $jbad++; Write-Host "  [WARN] $($jp): $($_.Exception.Message)" -ForegroundColor Yellow }
        }
        Remove-Item -LiteralPath "$jp.pulltmp" -Force -ErrorAction SilentlyContinue
    }
    # Journal nur loeschen, wenn alles zurueckgestellt ist (sonst bleibt die Start-Sperre bestehen)
    if ($jbad) { Stop-HMPull "$jbad Datei(en) konnten nicht zurueckgestellt werden (gesperrt?) - Programm schliessen und Pull erneut ausfuehren." }
    Remove-Item -LiteralPath $jr -Force -ErrorAction SilentlyContinue
}
$cfg = Get-HMUpdateConfig $cfgDir
if (-not $Owner) { $Owner = $cfg.Owner }
if (-not $Repo)  { $Repo = $cfg.Repo }
# andere Quelle per Parameter: eingebauter Fingerabdruck gilt nur fuer die offizielle Quelle
if ($cfg.DefaultSigner -and ($Owner -ne $cfg.Owner -or $Repo -ne $cfg.Repo)) {
    $cfg.SignerThumbprint = Get-HMDefaultSigner $Owner $Repo
    $cfg.RequireSignature = ([bool]$cfg.SignerThumbprint -and -not $cfg.AllowUnsigned)
}
if ($Branch -and $cfg.RequireSignature) { Stop-HMPull "Branch-Stand ist nicht signiert - nur moeglich, wenn 'Nur signierte Updates' ausgeschaltet ist (Einstellungen > Update)." }
if (-not $Channel) { $Channel = $cfg.Channel }
$useBranch = [bool]$Branch -or ($cfg.UseBranch -and -not $Version)
if (-not $Branch) { $Branch = $cfg.Branch }

# --- Token (nur private Repos) ---
# Benutzerordner wie Get-HUSafeName in Core\HU.Paths.psm1 (Pull laeuft ohne die Module der App)
$tokUser = ("$env:USERNAME" -replace '[^\p{L}\p{Nd}\.\-_]', '_').Trim('.', ' '); if (-not $tokUser) { $tokUser = 'unbekannt' }
$tokFile = Join-Path $Target ('Data\Benutzer\{0}\Token\GitHub.xml' -f $tokUser)
$Token = "$env:HU_GITHUB_TOKEN".Trim()
if (-not $Token -and (Test-Path -LiteralPath $tokFile)) {
    try { $c = Import-Clixml -Path $tokFile; if ($c -is [System.Management.Automation.PSCredential]) { $Token = $c.GetNetworkCredential().Password.Trim() } }
    catch { Write-Host '[WARN] Gespeicherter GitHub-Token nicht lesbar (anderer Benutzer?)' -ForegroundColor Yellow }
}

Write-Host '=== HU-Amt Pull ===' -ForegroundColor Cyan
Write-Host "Benutzer:   $env:USERNAME" -ForegroundColor Gray
Write-Host "Zielordner: $Target" -ForegroundColor Gray
Write-Host "Modus:      $mode" -ForegroundColor Gray
Write-Host "Quelle:     $Owner/$Repo$(if ($Token) { ' (mit Token)' })" -ForegroundColor Gray
Write-Host ''

# --- Auf Beenden des Tools warten (Update-Button) ---
if ($WaitPid -gt 0) {
    $proc = Get-Process -Id $WaitPid -ErrorAction SilentlyContinue
    if ($proc) {
        Write-Host "Warte bis HU-Amt (PID $WaitPid) beendet ist..." -ForegroundColor Yellow
        try { $null = $proc.WaitForExit(60000) } catch { }
        if (-not $proc.HasExited) {
            Write-Host '[WARN] HU-Amt nach 60 s noch aktiv - wird beendet' -ForegroundColor Yellow
            try { Stop-Process -Id $WaitPid -Force; $null = $proc.WaitForExit(10000) } catch { Write-Host "[WARN] Beenden fehlgeschlagen: $($_.Exception.Message)" -ForegroundColor Yellow }
        } else { Write-Host 'HU-Amt beendet.' -ForegroundColor Green }
    }
    Start-Sleep -Milliseconds 800
}

# --- Gemeinsamer Netzordner: niemand sonst darf die App offen haben (Data\Laufend\*.lock wird offen gehalten) ---
# Eine Lock-Datei, die sich loeschen laesst, ist verwaist (Absturz) und wird entfernt.
function Get-HMOpenSessions([string]$Root) {
    $run = Join-Path $Root 'Data\Laufend'
    $out = @()
    if (-not (Test-Path -LiteralPath $run)) { return $out }
    foreach ($f in @(Get-ChildItem -LiteralPath $run -Filter '*.lock' -File -ErrorAction SilentlyContinue)) {
        try { [IO.File]::Delete($f.FullName); continue } catch { }
        $who = $f.BaseName -replace '_', ' am PC '
        try {
            $fs = [IO.File]::Open($f.FullName, 'Open', 'Read', 'ReadWrite')
            try { $p = @((New-Object IO.StreamReader($fs)).ReadToEnd() -split '\|'); if ($p.Count -ge 2) { $who = "$($p[0]) am PC $($p[1])" } } finally { $fs.Dispose() }
        } catch { }
        $out += $who
    }
    return $out
}
$open = @(Get-HMOpenSessions $Target)
if ($open.Count) { Stop-HMPull "HU-Amt ist noch geoeffnet - bitte zuerst schliessen: $($open -join ', '). Es wurde NICHTS veraendert." }

if (-not (Test-Path $Target)) { New-Item -ItemType Directory -Path $Target -Force | Out-Null }
try {
    $probe = Join-Path $Target ('.pullprobe_' + [Guid]::NewGuid().ToString('N'))
    [System.IO.File]::WriteAllText($probe, 'x'); Remove-Item $probe -Force
} catch { Stop-HMPull "Kein Schreibzugriff auf '$Target'." }

function New-GHHeaders([string]$Accept) {
    $h = @{ Accept = $Accept; 'User-Agent' = 'HU-Amt-Pull' }
    if ($script:Token) { $h['Authorization'] = "token $($script:Token)" }
    return $h
}
$ProgressPreference = 'SilentlyContinue'

# --- 1. Version waehlen (Release) bzw. Branch
$ref = $null; $rel = $null; $manifest = $null; $verified = 'ohne Pruefsumme'; $newCred = $null; $triedNoToken = $false
for ($attempt = 1; $attempt -le 3 -and -not $ref; $attempt++) {
    try {
        if ($useBranch) {
            $r = Invoke-RestMethod "https://api.github.com/repos/$Owner/$Repo/git/refs/heads/$Branch" -Headers (New-GHHeaders 'application/vnd.github.v3+json') -UseBasicParsing
            $ref = "$($r.object.sha)"
        } else {
            $list = @(Get-HMReleases $Owner $Repo $Token)
            if ($Version) {
                $want = ConvertTo-HMVersion $Version
                $rel = @($list | Where-Object { $want -and $_.Version -eq $want })[0]
                if (-not $rel) { Stop-HMPull "Version $Version gibt es nicht als Release ($Owner/$Repo)." }
            } else {
                $rel = Select-HMRelease $list $Channel -SignedOnly:$cfg.RequireSignature
                if (-not $rel) { Stop-HMPull "Kein $(if ($cfg.RequireSignature) { 'signiertes ' })Release im Kanal $(if ($Channel -eq 'Test') { 'Test' } else { 'Stabil' }) gefunden ($Owner/$Repo)." }
            }
            # Tag -> Commit (funktioniert auch bei annotierten Tags)
            $cm = Invoke-RestMethod "https://api.github.com/repos/$Owner/$Repo/commits/$($rel.Tag)" -Headers (New-GHHeaders 'application/vnd.github.v3+json') -UseBasicParsing
            $ref = "$($cm.sha)"
        }
    } catch {
        $code = 0; try { $code = [int]$_.Exception.Response.StatusCode } catch { }
        # Gespeicherter Token abgelaufen/widerrufen, Repo evtl. inzwischen oeffentlich -> einmal ohne Token versuchen
        if ($Token -and -not $newCred -and -not $triedNoToken -and $code -eq 401) {
            Write-Host '[INFO] Token wird abgelehnt (HTTP 401) - versuche ohne Token...' -ForegroundColor Yellow
            $Token = ''; $triedNoToken = $true
            continue
        }
        if (-not $newCred -and -not $NonInteractive -and $code -in 401, 403, 404) {
            Write-Host "[INFO] $Owner/$Repo ist $(if ($Token) { 'mit dem Token' } else { 'ohne Token' }) nicht erreichbar (HTTP $code) - privates Repo?" -ForegroundColor Yellow
            Write-Host 'Read-only GitHub-Token eingeben (Fine-grained PAT, Contents: Read) - leer = Abbruch:' -ForegroundColor Yellow
            $sec = Read-Host -AsSecureString
            if ($sec -and $sec.Length -gt 0) {
                $newCred = New-Object System.Management.Automation.PSCredential ('github', $sec)
                $Token = $newCred.GetNetworkCredential().Password.Trim()
                continue
            }
        }
        Stop-HMPull "GitHub nicht erreichbar: $($_.Exception.Message)"
    }
}
if (-not $ref) { Stop-HMPull 'Version nicht ermittelbar' }
if ($newCred) {
    try { if (-not (Test-Path -LiteralPath $cfgDir)) { New-Item -ItemType Directory -Path $cfgDir -Force | Out-Null }; $newCred | Export-Clixml -Path $tokFile -Force; Write-Host "GitHub-Token gespeichert (DPAPI): $tokFile" -ForegroundColor Green }
    catch { Write-Host "[WARN] Token nicht gespeichert: $($_.Exception.Message)" -ForegroundColor Yellow }
}
if ($useBranch) { Write-Host "Version:    Entwicklungsstand Branch '$Branch' ($($ref.Substring(0, [Math]::Min(12, $ref.Length))))" -ForegroundColor Yellow }
else { Write-Host "Version:    $($rel.Tag)  ($(if ($rel.Prerelease) { 'Test' } else { 'Stabil' }), $($rel.Date))" -ForegroundColor Gray }

# --- 2. Pruefsumme und Signatur
if (-not $useBranch) {
    $manBytes = $null
    if ($rel.ManifestUrl) {
        try { $manBytes = Get-HMReleaseAsset $rel.ManifestUrl $rel.ManifestApi $Token } catch { Stop-HMPull "Pruefsummen-Datei nicht ladbar: $($_.Exception.Message)" }
        $manifest = ConvertFrom-HMManifest ([System.Text.Encoding]::UTF8.GetString($manBytes))
        if (-not $manifest.Count) { Stop-HMPull 'Pruefsummen-Datei ist leer oder beschaedigt.' }
        $verified = 'Pruefsumme (SHA256)'
    }
    if ($cfg.RequireSignature) {
        if (-not $manBytes) { Stop-HMPull "Nur signierte Updates erlaubt - $($rel.Tag) hat keine Pruefsummen-Datei." }
        $sigBytes = $null
        if ($rel.SignatureUrl) { try { $sigBytes = Get-HMReleaseAsset $rel.SignatureUrl $rel.SignatureApi $Token } catch { Stop-HMPull "Signatur nicht ladbar: $($_.Exception.Message)" } }
        $why = Test-HMManifestSignature $manBytes $sigBytes $cfg.SignerThumbprint
        if ($why) { Stop-HMPull "Nur signierte Updates erlaubt - $($rel.Tag): $why" }
        $verified = "Pruefsumme + Signatur ($($cfg.SignerThumbprint.Substring(0, [Math]::Min(8, $cfg.SignerThumbprint.Length)))...)"
        Write-Host 'Signatur:   gueltig' -ForegroundColor Green
    }
    if (-not $manifest) {
        Write-Host "[WARN] $($rel.Tag) hat keine Pruefsummen-Datei (die automatischen Tests sind noch nicht fertig oder fehlgeschlagen)." -ForegroundColor Yellow
        if ($NonInteractive) { Stop-HMPull 'ohne Pruefsumme nicht erlaubt (-NonInteractive)' }
        $a = Read-Host 'Trotzdem ohne Pruefung laden? (j/N)'
        if ("$a".Trim() -notmatch '^[jJyY]') { Stop-HMPull 'abgebrochen' }
    }
}
Write-Host "Pruefung:   $verified" -ForegroundColor Gray

# --- 3. Dateiliste
try { $tree = Invoke-RestMethod "https://api.github.com/repos/$Owner/$Repo/git/trees/${ref}?recursive=1" -Headers (New-GHHeaders 'application/vnd.github.v3+json') -UseBasicParsing }
catch { Stop-HMPull "Dateiliste nicht lesbar: $($_.Exception.Message)" }
$files = @($tree.tree | Where-Object { $_.type -eq 'blob' -and "$($_.path)" -notlike '.github/*' })
if (-not $files.Count) { Stop-HMPull 'Repo-Inhalt leer' }
Write-Host "$($files.Count) Dateien`n" -ForegroundColor Gray

# --- 4. alles laden und pruefen (noch nichts ersetzen)
$staged = New-Object System.Collections.Generic.List[object]
$fail = 0
foreach ($f in $files) {
    $local = Join-Path $Target ($f.path -replace '/', '\')
    $dir = Split-Path $local -Parent
    try { if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null } } catch { Write-Host "  $($f.path): Ordner nicht anlegbar" -ForegroundColor Red; $fail++; continue }
    Write-Host ('  {0,-48} ... ' -f $f.path) -NoNewline
    $tmp = "$local.pulltmp"
    $done = $false; $lastErr = ''
    for ($try = 1; $try -le 5 -and -not $done; $try++) {
        try {
            $enc = (($f.path -split '/') | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
            # privat: Contents-API mit Token; oeffentlich: raw-CDN (kein API-Limit), jeweils per Commit-SHA -> nie veraltet
            if ($Token) { Invoke-WebRequest "https://api.github.com/repos/$Owner/$Repo/contents/${enc}?ref=$ref" -Headers (New-GHHeaders 'application/vnd.github.v3.raw') -UseBasicParsing -OutFile $tmp }
            else { Invoke-WebRequest "https://raw.githubusercontent.com/$Owner/$Repo/$ref/$enc" -Headers @{ 'User-Agent' = 'HU-Amt-Pull' } -UseBasicParsing -OutFile $tmp }
            $done = $true
        } catch { $lastErr = $_.Exception.Message; Start-Sleep -Seconds 1 }
    }
    if (-not $done) { Write-Host "FEHLER: $lastErr" -ForegroundColor Red; $fail++; continue }
    if ($manifest) {
        $want = $manifest[$f.path]
        $have = Get-HMFileSha256 $tmp
        if (-not $want) { Write-Host 'FEHLER: nicht in der Pruefsummen-Datei' -ForegroundColor Red; $fail++; $staged.Add(@($tmp, $local)); continue }
        if ($have -ne $want) { Write-Host 'FEHLER: Pruefsumme stimmt nicht' -ForegroundColor Red; $fail++; $staged.Add(@($tmp, $local)); continue }
        Write-Host 'OK (geprueft)' -ForegroundColor Green
    } else { Write-Host "OK ($((Get-Item -LiteralPath $tmp -Force).Length) bytes)" -ForegroundColor Green }
    $staged.Add(@($tmp, $local))
}
if ($fail) {
    foreach ($s in $staged) { Remove-Item -LiteralPath $s[0] -Force -ErrorAction SilentlyContinue }
    Stop-HMPull "$fail Datei(en) nicht geladen oder Pruefsumme falsch - es wurde NICHTS veraendert, HU-Amt bleibt auf der bisherigen Version."
}

# --- 5. ersetzen - mit Ruecksicherung: jede bisherige Datei wird erst zu *.pullold umbenannt;
#        scheitert ein Schritt, wird alles zurueckgestellt (nie halb aktualisiert).
#        Bricht Pull mittendrin ab (Absturz, Strom), stellt der naechste Start anhand des Journals zurueck.
$journal = Join-Path $cfgDir 'pull-journal.json'
$moved = New-Object System.Collections.Generic.List[object]   # @(Ziel, Sicherung oder '')
$repl = 0; $lastErr = ''
try { if (-not (Test-Path -LiteralPath $cfgDir)) { New-Item -ItemType Directory -Path $cfgDir -Force | Out-Null } } catch { }
try { ConvertTo-Json -Depth 10 -InputObject @($staged | ForEach-Object { "$($_[1])" }) | Set-Content -LiteralPath $journal -Encoding UTF8 -ErrorAction Stop }
catch {
    foreach ($s in $staged) { Remove-Item -LiteralPath $s[0] -Force -ErrorAction SilentlyContinue }
    Stop-HMPull "Journal nicht schreibbar ($($_.Exception.Message)) - es wurde NICHTS veraendert."
}
foreach ($s in $staged) {
    $old = "$($s[1]).pullold"
    $done = $false
    $had = Test-Path -LiteralPath $s[1]      # einmal vor den Wiederholungen bestimmen
    $backedUp = $false                       # bisherige Datei liegt gerade als *.pullold
    for ($try = 1; $try -le 5 -and -not $done; $try++) {
        try {
            if ($had -and -not $backedUp) { Move-Item -LiteralPath $s[1] -Destination $old -Force; $backedUp = $true }
            try { Move-Item -LiteralPath $s[0] -Destination $s[1] -Force }
            catch { if ($backedUp) { try { Move-Item -LiteralPath $old -Destination $s[1] -Force; $backedUp = $false } catch { } }; throw }
            $moved.Add(@($s[1], $(if ($had) { $old } else { '' })))
            $done = $true
        } catch { $lastErr = $_.Exception.Message; Start-Sleep -Seconds 1 }
    }
    if (-not $done) {
        if ($backedUp) { try { Move-Item -LiteralPath $old -Destination $s[1] -Force } catch { Write-Host "  [WARN] $($s[1]): Sicherung nicht zurueckgestellt ($($_.Exception.Message))" -ForegroundColor Yellow } }
        Write-Host "  $($s[1]): nicht ersetzbar ($lastErr)" -ForegroundColor Red; $repl++; break
    }
}
if ($repl) {
    # zurueckstellen (umgekehrte Reihenfolge): neue Dateien entfernen, Sicherungen zurueck
    for ($k = $moved.Count - 1; $k -ge 0; $k--) {
        $m = $moved[$k]
        try { if ($m[1]) { Move-Item -LiteralPath $m[1] -Destination $m[0] -Force } else { Remove-Item -LiteralPath $m[0] -Force } }
        catch { Write-Host "  [WARN] $($m[0]): nicht zurueckgestellt ($($_.Exception.Message))" -ForegroundColor Yellow }
    }
    foreach ($s in $staged) { Remove-Item -LiteralPath $s[0] -Force -ErrorAction SilentlyContinue }
    Remove-Item -LiteralPath $journal -Force -ErrorAction SilentlyContinue
    Stop-HMPull "Eine Datei war gesperrt ($lastErr) - alles wurde zurueckgestellt, HU-Amt bleibt auf der bisherigen Version. HU-Amt schliessen und erneut versuchen."
}
# Abschluss: zuerst das Journal loeschen (ab hier gilt der neue Stand), dann die Sicherungen.
# Umgekehrt koennte ein Abbruch beim Aufraeumen einen Mischstand aus alt und neu herstellen.
Remove-Item -LiteralPath $journal -Force -ErrorAction SilentlyContinue
foreach ($m in $moved) {
    try { Unblock-File -LiteralPath $m[0] -ErrorAction SilentlyContinue } catch { }
    if ($m[1]) { Remove-Item -LiteralPath $m[1] -Force -ErrorAction SilentlyContinue }
}
$ok = $moved.Count

# --- 6. Dateien entfernen, die es im neuen Stand nicht mehr gibt (nur solche, die frueher per Pull installiert wurden)
$newFiles = @($files | ForEach-Object { "$($_.path)" })
$instFile = Join-Path $cfgDir 'installed.json'
$prevFiles = @()
try { if (Test-Path -LiteralPath $instFile) { $pi = Get-Content -LiteralPath $instFile -Raw -Encoding UTF8 | ConvertFrom-Json; if ($pi.PSObject.Properties['FileList']) { $prevFiles = @($pi.FileList) } } } catch { }
$removedOld = 0
$keepDirs = '^(Config|Data|Addons|AddonsDev|bin|Logs)/'
$rootFull = [IO.Path]::GetFullPath($Target).TrimEnd('\') + '\'
foreach ($oldPath in @($prevFiles | Where-Object { $_ -and $newFiles -notcontains $_ -and $_ -notmatch $keepDirs -and $_ -notmatch '(^|/)\.\.(/|$)' -and $_ -notmatch '^[\\/]|:' })) {
    $p = [IO.Path]::GetFullPath((Join-Path $Target ($oldPath -replace '/', '\')))
    if (-not $p.StartsWith($rootFull, [StringComparison]::OrdinalIgnoreCase)) { continue }
    if (Test-Path -LiteralPath $p -PathType Leaf) {
        try { Remove-Item -LiteralPath $p -Force; $removedOld++; Write-Host "  entfernt (nicht mehr im Programm): $oldPath" -ForegroundColor DarkGray } catch { }
    }
}
try {
    [pscustomobject][ordered]@{
        Version = $(if ($useBranch) { "Branch $Branch" } else { "$($rel.Version)" }); Ref = "$ref"; Channel = $(if ($useBranch) { 'Branch' } elseif ($rel.Prerelease) { 'Test' } else { 'Stable' })
        Check = $verified; Date = (Get-Date).ToString('yyyy-MM-dd HH:mm'); Files = $ok; FileList = $newFiles
    } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $instFile -Encoding UTF8
} catch { }
# --- 7. Erstinstallation: wer installiert, wird Verwalter (nur ohne bestehende Liste; ein Update aendert sie nie)
if ($isFirstInstall -and $repl -eq 0) {
    $admFile = Join-Path $Target 'Data\Verwalter.txt'
    if (-not (Test-Path -LiteralPath $admFile) -and "$env:USERNAME".Trim()) {
        try {
            $admDir = Split-Path -Parent $admFile
            if (-not (Test-Path -LiteralPath $admDir)) { [void](New-Item -ItemType Directory -Path $admDir -Force) }
            [IO.File]::WriteAllText($admFile, "# Verwalter: ein Windows-Benutzername je Zeile`r`n$("$env:USERNAME".Trim())`r`n", (New-Object System.Text.UTF8Encoding $true))
            Write-Host "Verwalter:  $env:USERNAME (Erstinstallation)" -ForegroundColor Gray
        } catch { Write-Host "Verwalter-Liste nicht angelegt: $($_.Exception.Message)" -ForegroundColor Yellow }
    }
}
Write-Host "`n=== Pull fertig === $ok Dateien ($verified)$(if ($repl) { " | $repl NICHT ersetzt" })" -ForegroundColor Cyan
if ($NoStart) { if ($repl) { exit 1 } else { exit 0 } }
if ($repl -eq 0) {
    Write-Host 'Starte HU-Amt...' -ForegroundColor Yellow
    Start-Sleep -Seconds 1
    $exe = Join-Path $Target 'HU-Amt.exe'
    if (Test-Path $exe) { Start-Process -FilePath $exe -WorkingDirectory $Target }
    else {
        # ohne Starter: gleiche Regel wie HU-Amt.exe - PowerShell 7, wenn installiert (Config\powershell-host.txt: auto | 5 | 7)
        $mode = 'auto'; $hf = Join-Path $Target 'Config\powershell-host.txt'
        try { if (Test-Path -LiteralPath $hf) { $mode = ([IO.File]::ReadAllText($hf)).Trim().ToLowerInvariant() } } catch { }
        $pw = ''
        if ($mode -ne '5') { foreach ($p in @($(if ($env:ProgramW6432) { Join-Path $env:ProgramW6432 'PowerShell\7\pwsh.exe' }), (Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'))) { if ($p -and (Test-Path -LiteralPath $p)) { $pw = $p; break } } }
        $main = "`"$(Join-Path $Target 'Main.ps1')`""
        if ($pw) { Start-Process -FilePath $pw -ArgumentList '-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-File', $main -WorkingDirectory $Target }
        else { Start-Process powershell.exe -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-File', $main -WorkingDirectory $Target }
    }
    exit 0
}
Stop-HMPull "Einzelne Dateien konnten nicht ersetzt werden (gesperrt?) - HU-Amt wird NICHT automatisch gestartet. Manuell: cd '$Target'; .\Pull.ps1"
