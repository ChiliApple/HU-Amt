#Requires -Version 5.1
<#
.SYNOPSIS
    Zentrale und Aussenstelle: Art der Installation, freigeschaltete Module und das Paket (ZIP) fuer eine Aussenstelle.
.DESCRIPTION
    Config\installation.json (gilt fuer die ganze Installation, bleibt bei Updates erhalten):
        { "modus": "Zentrale" | "Aussenstelle", "name": "Kindergarten", "moduleAus": [ "OCR", "PDF" ] }
    Fehlt die Datei, gilt: Zentrale, alle Module aktiv.

    Paket fuer eine Aussenstelle (New-HUOutpostPackage), Aufbau des ZIP:
        Einrichten.cmd      startet Programm\Einrichten.ps1 (kopiert nach %LOCALAPPDATA%\HU-Amt-<Name>, ohne Adminrechte)
        LIESMICH.txt        kurze Anleitung
        Programm\           die Installation fuer die Aussenstelle:
          - Programmdateien der eigenen Installation, geprueft gegen die signierte Pruefsummen-Datei des Releases
            (liegt als Config\HU-Amt-files.sha256 und .p7s bei, damit die Aussenstelle beim ersten Start pruefen kann)
          - gewaehlte Module aus Addons\ (nur Dateien laut ihrer signierten Pruefsummen-Datei)
          - Data\Verwalter.txt, Config\installation.json, installed.json, update.json (nur Kanal), paket.json
          - auf Wunsch je Modul Data\<Id>\einstellungen.json (z. B. Reise mit dem Schluessel fuer die Routenberechnung)
          - auf Wunsch Zugang.humodule mit den Nur-Lese-Tokens fuer die mitgelieferten Module (Updates in der Aussenstelle)
      Nie im Paket: Daten, Benutzer, Token-Dateien, Protokolle, Sicherungen der Zentrale.

    Grenze: Die Pruefung beim ersten Start erkennt beschaedigte oder unvollstaendige Pakete. Gegen absichtliche Aenderungen
    am ZIP schuetzt sie nicht, weil Pruefcode und Pruefsummen-Datei im selben ZIP liegen. Das Paket deshalb nur auf
    vertrauenswuerdigem Weg weitergeben.
.NOTES
    Dot-Source aus Main.ps1 (nach Core-Update.ps1 und Core-Addons.ps1). Laeuft unter Windows PowerShell 5.1 und PowerShell 7.
    Zielmaschine: der PC, auf dem die App laeuft.
#>

$script:HUInstConfigName  = 'installation.json'
$script:HUPackageInfoName = 'paket.json'
$script:HUPackageKeyName  = 'Zugang.humodule'
$script:HUPackageAppDir   = 'Programm'

# ------------------------------------------------------------------ Art der Installation -----------------------------------------------
function Get-HUInstallationConfig([string]$ConfigDir) {
    $o = [pscustomobject]@{ Mode = 'Zentrale'; Name = ''; ModulesOff = @() }
    $f = Join-Path $ConfigDir $script:HUInstConfigName
    if (-not (Test-Path -LiteralPath $f)) { return $o }
    try { $j = [IO.File]::ReadAllText($f, [Text.Encoding]::UTF8) | ConvertFrom-Json -ErrorAction Stop } catch { return $o }
    if ($j.PSObject.Properties['modus'] -and "$($j.modus)" -eq 'Aussenstelle') { $o.Mode = 'Aussenstelle' }
    if ($j.PSObject.Properties['name']) { $o.Name = "$($j.name)".Trim() }
    if ($j.PSObject.Properties['moduleAus']) { $o.ModulesOff = @($j.moduleAus | ForEach-Object { "$_".Trim() } | Where-Object { Test-HUAddonId $_ } | Sort-Object -Unique) }
    return $o
}

function Save-HUInstallationConfig([string]$ConfigDir, $Config) {
    if ("$($Config.Mode)" -notin 'Zentrale', 'Aussenstelle') { throw "Unbekannte Art der Installation: '$($Config.Mode)'" }
    $name = "$($Config.Name)".Trim()
    if ("$($Config.Mode)" -eq 'Aussenstelle' -and -not $name) { throw 'Eine Außenstelle braucht einen Namen.' }
    if ($name.Length -gt 60) { throw 'Der Name ist zu lang (höchstens 60 Zeichen).' }
    $o = [ordered]@{ modus = "$($Config.Mode)"; name = $name; moduleAus = @(@($Config.ModulesOff) | ForEach-Object { "$_" } | Where-Object { Test-HUAddonId $_ } | Sort-Object -Unique) }
    if (-not (Test-Path -LiteralPath $ConfigDir)) { [void](New-Item -ItemType Directory -Path $ConfigDir -Force) }
    [IO.File]::WriteAllText((Join-Path $ConfigDir $script:HUInstConfigName), (ConvertTo-Json -InputObject $o -Depth 5), (New-Object Text.UTF8Encoding $true))
}

# Anzeige im Fenstertitel, z. B. "Außenstelle Kindergarten" ('' in der Zentrale)
function Get-HUOutpostLabel($Config) {
    if (-not $Config -or "$($Config.Mode)" -ne 'Aussenstelle') { return '' }
    return ("Au$([char]0xDF)enstelle $($Config.Name)").Trim()
}

# Name fuer Ordner und ZIP: nur Buchstaben, Ziffern, Bindestrich (Umlaute bleiben)
function ConvertTo-HUPackageName([string]$Name) {
    $n = ("$Name".Trim() -replace '[^\p{L}\p{Nd}]+', '-').Trim('-')
    if (-not $n) { throw 'Der Name der Außenstelle ergibt keinen gültigen Ordnernamen.' }
    if ($n.Length -gt 40) { $n = $n.Substring(0, 40).Trim('-') }
    return $n
}

# Reste abgebrochener Pakete (Klartext-Zugang, Einstellungen) aus TEMP entfernen, wenn aelter als eine Stunde
function Remove-HUOldPackageTemp {
    $lim = (Get-Date).AddHours(-1)
    foreach ($d in @(Get-ChildItem -LiteralPath ([IO.Path]::GetTempPath()) -Directory -Filter 'HUAmtPaket-*' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt $lim })) {
        try { Remove-Item -LiteralPath $d.FullName -Recurse -Force -ErrorAction Stop } catch { }
    }
}

# ------------------------------------------------------------------ Pruefsummen-Datei des Programms ------------------------------------
<#
.SYNOPSIS
    Signierte Pruefsummen-Datei der installierten Programmversion holen und pruefen.
    Zuerst Config\HU-Amt-files.sha256 (+ .p7s, legt Pull.ps1 ab), sofern sie zu den Programmdateien passt,
    sonst vom Release der installierten Version (oeffentlich, ohne Token).
    Rueckgabe: @{ Manifest = byte[]; Signature = byte[]; Source = 'lokal'|'GitHub' } oder wirft mit Grund und naechstem Schritt.
#>
function Get-HUVerifiedInstallManifest {
    param([string]$ConfigDir, [string]$Version, [string]$Owner, [string]$Repo, [string]$Thumbprint, [string]$AppRoot = '')
    $mf = Join-Path $ConfigDir $script:HMManifestName; $sf = Join-Path $ConfigDir $script:HMSignatureName
    if ((Test-Path -LiteralPath $mf) -and (Test-Path -LiteralPath $sf)) {
        $man = [IO.File]::ReadAllBytes($mf); $sig = [IO.File]::ReadAllBytes($sf)
        # die lokale Datei gilt nur, wenn sie signiert ist UND zu den Programmdateien passt
        # (sonst stammt sie von einem anderen Stand, z. B. nach dem Laden einer Vorversion mit einem aelteren Pull.ps1)
        $root = $(if ($AppRoot) { $AppRoot } else { Split-Path $ConfigDir -Parent })
        if (-not (Test-HMManifestSignature $man $sig $Thumbprint) -and -not @(Test-HUInstallFiles -AppRoot $root -Manifest $man).Count) { return @{ Manifest = $man; Signature = $sig; Source = 'lokal' } }
    }
    $next = 'Bitte mit Update die aktuelle Version laden und das Paket danach noch einmal erstellen.'
    $v = ConvertTo-HMVersion $Version
    if (-not $v) { throw "Diese Installation ist keine freigegebene Version ($Version).`n`n$next" }
    try { $rel = @(Get-HMReleases $Owner $Repo '' | Where-Object { $_.Version -eq $v })[0] }
    catch { throw "Die freigegebene Version konnte nicht geprüft werden (keine Verbindung zu GitHub?).`n`nBitte die Internetverbindung prüfen und es noch einmal versuchen.`n`n($($_.Exception.Message))" }
    if (-not $rel) { throw "Zur installierten Version $v gibt es keine freigegebene Version mehr.`n`n$next" }
    if (-not $rel.ManifestUrl -or -not $rel.SignatureUrl) { throw "Die Version $v ist noch nicht signiert.`n`n$next" }
    try {
        $man = Get-HMReleaseAsset $rel.ManifestUrl $rel.ManifestApi ''
        $sig = Get-HMReleaseAsset $rel.SignatureUrl $rel.SignatureApi ''
    } catch { throw "Die Prüfsummen der Version $v konnten nicht geladen werden.`n`nBitte die Internetverbindung prüfen und es noch einmal versuchen." }
    $why = Test-HMManifestSignature $man $sig $Thumbprint
    if ($why) { throw "Die Signatur der Version $v stimmt nicht ($why).`n`n$next" }
    return @{ Manifest = $man; Signature = $sig; Source = 'GitHub' }
}

# Programmdateien gegen die Pruefsummen-Datei pruefen. Rueckgabe: Liste der Abweichungen (leer = alles gleich)
function Test-HUInstallFiles([string]$AppRoot, [byte[]]$Manifest) {
    $m = ConvertFrom-HMManifest ([Text.Encoding]::UTF8.GetString($Manifest))
    if (-not $m.Count) { return @('Prüfsummen-Datei leer') }
    $bad = @()
    foreach ($k in $m.Keys) {
        if ($k -match '(^[\\/])|(\.\.)|:') { $bad += "unzulässiger Pfad: $k"; continue }
        $p = Join-Path $AppRoot ($k -replace '/', '\')
        if (-not (Test-Path -LiteralPath $p -PathType Leaf)) { $bad += "fehlt: $k"; continue }
        if ((Get-HMFileSha256 $p) -ne $m[$k]) { $bad += "verändert: $k" }
    }
    return $bad
}

# Token zu einem Modul: Repository genau "HU-Amt-<Id>" oder "<Id>" (kein Teiltreffer)
function Find-HUAddonSourceFor([string]$Id, [object[]]$Sources) {
    foreach ($s in @($Sources)) { if ("$($s.Repo)" -ieq "HU-Amt-$Id") { return $s } }
    foreach ($s in @($Sources)) { if ("$($s.Repo)" -ieq $Id) { return $s } }
    return $null
}

# Dateien eines Modulordners laut seiner signierten Pruefsummen-Datei (relativ, mit \). Ohne Datei: alle Dateien.
function Get-HUAddonPackageFiles([string]$Dir) {
    $mf = Join-Path $Dir $script:HUAddonManifest
    if (Test-Path -LiteralPath $mf) {
        $m = ConvertFrom-HMManifest ([IO.File]::ReadAllText($mf, [Text.Encoding]::UTF8))
        $list = @($m.Keys | Where-Object { $_ -notmatch '(^[\\/])|(\.\.)|:' } | ForEach-Object { $_ -replace '/', '\' })
        return @($list + @($script:HUAddonManifest, $script:HUAddonSignature) | Where-Object { Test-Path -LiteralPath (Join-Path $Dir $_) -PathType Leaf })
    }
    $full = [IO.Path]::GetFullPath($Dir).TrimEnd('\')
    return @(Get-ChildItem -LiteralPath $Dir -Recurse -File -Force | ForEach-Object { $_.FullName.Substring($full.Length + 1) })
}

# ------------------------------------------------------------------ Paket erstellen ----------------------------------------------------
<#
.SYNOPSIS
    ZIP fuer eine Aussenstelle erstellen (Aufbau siehe Kopf der Datei). Netz wird nicht gebraucht.
.PARAMETER Manifest / Signature
    Signierte Pruefsummen-Datei der installierten Programmversion (Get-HUVerifiedInstallManifest). Jede Programmdatei
    muss dazu passen, sonst entsteht kein Paket.
.PARAMETER Modules
    Je Modul @{ Id; Active; Settings } - Settings = Data\<Id>\einstellungen.json mitgeben (falls vorhanden).
.PARAMETER Sources
    Hinterlegte Nur-Lese-Tokens (Owner, Repo, Token) - nur die der mitgelieferten Module kommen in Zugang.humodule.
.PARAMETER Overwrite
    Vorhandene ZIP-Datei ersetzen. Ersetzt wird erst, wenn das neue Paket fertig ist.
.OUTPUTS
    @{ Path; Folder; Files; Modules; Settings; Tokens; Warnings }
#>
function New-HUOutpostPackage {
    param(
        [Parameter(Mandatory)][string]$AppRoot,
        [Parameter(Mandatory)][string]$DataRoot,
        [Parameter(Mandatory)][string]$ZipPath,
        [Parameter(Mandatory)][string]$Name,
        [string[]]$Admins = @(),
        [object[]]$Modules = @(),
        [object[]]$Sources = @(),
        [switch]$IncludeKeys,
        [Parameter(Mandatory)][byte[]]$Manifest,
        [byte[]]$Signature = @(),
        [string]$Version = '',
        [string]$Channel = 'Stable',
        [string]$Thumbprint = '',
        [string]$CreatedBy = "$env:USERNAME",
        [switch]$Overwrite
    )
    $res = [pscustomobject]@{ Path = ''; Folder = ''; Files = 0; Modules = @(); Settings = @(); Tokens = @(); Warnings = @() }
    $name = "$Name".Trim()
    if (-not $name) { throw 'Bitte einen Namen für die Außenstelle angeben.' }
    $folderName = "HU-Amt-$(ConvertTo-HUPackageName $name)"
    $adm = @(@($Admins) | ForEach-Object { "$_".Trim() } | Where-Object { $_ } | Sort-Object -Unique)
    if (-not $adm.Count) { throw 'Bitte mindestens einen Verwalter für die Außenstelle angeben.' }
    if (@($adm | Where-Object { $_ -match '[\r\n#]' }).Count) { throw 'Ein Verwaltername enthält unzulässige Zeichen.' }
    $mods = @(@($Modules) | Where-Object { $_ -and "$($_.Id)" })
    foreach ($mo in $mods) { if (-not (Test-HUAddonId "$($mo.Id)")) { throw "Ungültige Modul-Kennung: '$($mo.Id)'" } }
    if (-not @($mods | Where-Object { $_.Active }).Count) { $res.Warnings += 'Kein Modul ist aktiv: In der Außenstelle erscheint nur die Startseite.' }
    $zip = [IO.Path]::GetFullPath($ZipPath)
    if ((Test-Path -LiteralPath $zip) -and -not $Overwrite) { throw "Die Datei gibt es schon: $zip" }
    $zipDir = Split-Path $zip -Parent
    if (-not (Test-Path -LiteralPath $zipDir)) { throw "Den Zielordner gibt es nicht: $zipDir" }

    # 1. Programmdateien pruefen (alle muessen zur signierten Pruefsummen-Datei passen)
    $bad = @(Test-HUInstallFiles -AppRoot $AppRoot -Manifest $Manifest)
    if ($bad.Count) { throw ("Die Programmdateien passen nicht zur freigegebenen Version, es wurde kein Paket erstellt.`n`nBitte mit Update die aktuelle Version laden und es noch einmal versuchen.`n`n" + (($bad | Select-Object -First 8) -join "`n") + $(if ($bad.Count -gt 8) { "`n... und $($bad.Count - 8) weitere" })) }
    $man = ConvertFrom-HMManifest ([Text.Encoding]::UTF8.GetString($Manifest))

    $stageRoot = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtPaket-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
    $stage = Join-Path $stageRoot $script:HUPackageAppDir
    $utf8Bom = New-Object Text.UTF8Encoding $true
    $writeJson = { param($Path, $Obj) $d = Split-Path $Path -Parent; if (-not (Test-Path -LiteralPath $d)) { [void](New-Item -ItemType Directory -Path $d -Force) }; [IO.File]::WriteAllText($Path, (ConvertTo-Json -InputObject $Obj -Depth 6), $utf8Bom) }
    $tmpZip = "$zip.tmp"
    try {
        [void](New-Item -ItemType Directory -Path $stage -Force)
        foreach ($k in $man.Keys) {
            $dst = Join-Path $stage ($k -replace '/', '\')
            $dd = Split-Path $dst -Parent; if (-not (Test-Path -LiteralPath $dd)) { [void](New-Item -ItemType Directory -Path $dd -Force) }
            Copy-Item -LiteralPath (Join-Path $AppRoot ($k -replace '/', '\')) -Destination $dst -Force
            if ((Get-HMFileSha256 $dst) -ne $man[$k]) { throw "Kopie stimmt nicht: $k" }
            $res.Files++
        }
        if (-not (Test-Path -LiteralPath (Join-Path $stage 'Einrichten.ps1'))) { throw 'Diese Programmversion kann noch keine Pakete einrichten (Einrichten.ps1 fehlt). Bitte zuerst mit Update die aktuelle Version laden.' }
        $cfg = Join-Path $stage 'Config'
        if (-not (Test-Path -LiteralPath $cfg)) { [void](New-Item -ItemType Directory -Path $cfg -Force) }
        [IO.File]::WriteAllBytes((Join-Path $cfg $script:HMManifestName), $Manifest)
        if ($Signature -and $Signature.Length) { [IO.File]::WriteAllBytes((Join-Path $cfg $script:HMSignatureName), $Signature) }

        # 2. Module (nur Dateien laut ihrer signierten Pruefsumme)
        $addons = Join-Path $AppRoot 'Addons'
        foreach ($mo in $mods) {
            $id = "$($mo.Id)"
            $src = Join-Path $addons $id
            if (-not (Test-Path -LiteralPath $src -PathType Container)) { throw "Das Modul $id ist in dieser Installation nicht vorhanden." }
            $info = Read-HUAddonInfo $src
            if ($info.Error) { throw "Modul $($id): $($info.Error)" }
            if ($Thumbprint) { $why = Test-HUAddonFiles $src $Thumbprint; if ($why) { throw "Modul $id ist nicht vertrauenswürdig: $why" } }
            $dst = Join-Path (Join-Path $stage 'Addons') $id
            [void](New-Item -ItemType Directory -Path $dst -Force)
            foreach ($rel in @(Get-HUAddonPackageFiles $src)) {
                $t = Join-Path $dst $rel
                $td = Split-Path $t -Parent; if (-not (Test-Path -LiteralPath $td)) { [void](New-Item -ItemType Directory -Path $td -Force) }
                Copy-Item -LiteralPath (Join-Path $src $rel) -Destination $t -Force
            }
            if ($Thumbprint) { $why = Test-HUAddonFiles $dst $Thumbprint; if ($why) { throw "Kopie des Moduls $id stimmt nicht: $why" } }
            $res.Modules += "$($info.Title) $($info.Version)$(if (-not $mo.Active) { ' (aus)' })"
            # Einstellungen des Moduls (nur die Datei einstellungen.json, keine Daten); ohne Datei still nichts
            if ($mo.Settings) {
                $sf = Join-Path (Join-Path $DataRoot $id) 'einstellungen.json'
                if (Test-Path -LiteralPath $sf -PathType Leaf) {
                    $sd = Join-Path (Join-Path $stage 'Data') $id
                    [void](New-Item -ItemType Directory -Path $sd -Force)
                    Copy-Item -LiteralPath $sf -Destination (Join-Path $sd 'einstellungen.json') -Force
                    $res.Settings += $info.Title
                }
            }
        }

        # 3. Verwalter, Art der Installation, Update-Kanal, Herkunft
        $dataDir = Join-Path $stage 'Data'
        if (-not (Test-Path -LiteralPath $dataDir)) { [void](New-Item -ItemType Directory -Path $dataDir -Force) }
        [IO.File]::WriteAllText((Join-Path $dataDir 'Verwalter.txt'), "# Verwalter: ein Windows-Benutzername je Zeile`r`n$($adm -join "`r`n")`r`n", $utf8Bom)
        Save-HUInstallationConfig -ConfigDir $cfg -Config ([pscustomobject]@{ Mode = 'Aussenstelle'; Name = $name; ModulesOff = @($mods | Where-Object { -not $_.Active } | ForEach-Object { "$($_.Id)" }) })
        & $writeJson (Join-Path $cfg 'update.json') ([ordered]@{ Channel = $(if ("$Channel" -eq 'Test') { 'Test' } else { 'Stable' }) })
        & $writeJson (Join-Path $cfg 'installed.json') ([ordered]@{ Version = "$Version"; Ref = ''; Channel = $(if ("$Channel" -eq 'Test') { 'Test' } else { 'Stable' }); Check = 'Paket (Prüfsumme + Signatur)'; Date = (Get-Date).ToString('yyyy-MM-dd HH:mm'); Files = $res.Files; FileList = @($man.Keys | Sort-Object) })

        # 4. Zugang fuer Updates der mitgelieferten Module (Nur-Lese-Tokens der Zentrale)
        $keySha = ''
        if ($IncludeKeys) {
            $entries = @()
            foreach ($mo in $mods) {
                $s = Find-HUAddonSourceFor "$($mo.Id)" $Sources
                if ($s) { $entries += [ordered]@{ repo = "$($s.Owner)/$($s.Repo)"; token = "$($s.Token)" }; $res.Tokens += "$($s.Owner)/$($s.Repo)" }
                else { $res.Warnings += "Für $($mo.Id) ist kein Zugang hinterlegt: Dieses Modul bekommt in der Außenstelle keine Updates." }
            }
            if ($entries.Count) {
                $kf = Join-Path $stage $script:HUPackageKeyName
                [IO.File]::WriteAllText($kf, (ConvertTo-Json -InputObject ([ordered]@{ module = @($entries) }) -Depth 5), (New-Object Text.UTF8Encoding $false))
                $keySha = Get-HMFileSha256 $kf
            }
        }
        & $writeJson (Join-Path $cfg $script:HUPackageInfoName) ([ordered]@{ name = $name; erstellt = (Get-Date).ToString('yyyy-MM-dd HH:mm'); von = "$CreatedBy"; version = "$Version"; verwalter = @($adm); zugang = $(if ($keySha) { $script:HUPackageKeyName } else { '' }); zugangSha256 = $keySha; geprueft = $false })

        # 5. Einrichten.cmd und Anleitung (nur Tastaturzeichen, fuer Laien)
        $cmd = "@echo off`r`nrem HU-Amt fuer die Aussenstelle einrichten (ohne Administratorrechte)`r`n" +
               "if not exist `"%~dp0$($script:HUPackageAppDir)\Einrichten.ps1`" (`r`n  echo Bitte die ZIP-Datei zuerst mit Rechtsklick und Alle extrahieren auspacken.`r`n  pause`r`n  exit /b 1`r`n)`r`n" +
               "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"%~dp0$($script:HUPackageAppDir)\Einrichten.ps1`"`r`n"
        [IO.File]::WriteAllText((Join-Path $stageRoot 'Einrichten.cmd'), $cmd, (New-Object Text.ASCIIEncoding))
        $liesmich = @(
            "HU-Amt für die Außenstelle $name"
            ''
            'So richtest du HU-Amt ein (einmal je Windows-Benutzer):'
            ''
            '1. Die ZIP-Datei mit der rechten Maustaste anklicken und "Alle extrahieren ..." wählen.'
            '   Nicht direkt aus der ZIP-Datei starten.'
            '2. Im ausgepackten Ordner auf "Einrichten" doppelklicken.'
            '3. HU-Amt richtet sich ein, legt das Symbol "HU-Amt ' + $name + '" auf den Desktop und startet.'
            '   Beim ersten Start braucht der PC Internet.'
            '4. Fertig. Den ausgepackten Ordner und die ZIP-Datei danach löschen.'
            ''
            "Verwalter in dieser Außenstelle: $($adm -join ', ')"
            ''
            'Bei Fragen bitte in der Gemeinde melden.'
            ''
            "Erstellt am $((Get-Date).ToString('dd.MM.yyyy')) mit HU-Amt $Version."
        ) -join "`r`n"
        [IO.File]::WriteAllText((Join-Path $stageRoot 'LIESMICH.txt'), $liesmich + "`r`n", $utf8Bom)

        # 6. ZIP (vorhandene Datei erst ersetzen, wenn das neue Paket fertig ist)
        try { Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop } catch { }
        if (Test-Path -LiteralPath $tmpZip) { Remove-Item -LiteralPath $tmpZip -Force }
        [IO.Compression.ZipFile]::CreateFromDirectory($stageRoot, $tmpZip, [IO.Compression.CompressionLevel]::Optimal, $false)
        Move-Item -LiteralPath $tmpZip -Destination $zip -Force
        $res.Path = $zip; $res.Folder = $folderName
        return $res
    } finally {
        if (Test-Path -LiteralPath $stageRoot) { Remove-Item -LiteralPath $stageRoot -Recurse -Force -ErrorAction SilentlyContinue }
        if (Test-Path -LiteralPath $tmpZip) { Remove-Item -LiteralPath $tmpZip -Force -ErrorAction SilentlyContinue }
    }
}

# ------------------------------------------------------------------ Erster Start eines Pakets ------------------------------------------
<#
.SYNOPSIS
    Programm einer aus einem Paket eingerichteten Installation pruefen, bis es einmal geklappt hat.
    Rueckgabe: @{ Status = 'OK'|'Fehler'|'Keins'; Text } - 'Keins' = kein Paket oder schon erfolgreich geprueft.
    Bei 'Fehler' bleibt die Pruefung offen und kommt beim naechsten Start wieder.
#>
function Test-HUPackageFirstStart([string]$AppRoot, [string]$Thumbprint) {
    $cfg = Join-Path $AppRoot 'Config'
    $pf = Join-Path $cfg $script:HUPackageInfoName
    if (-not (Test-Path -LiteralPath $pf)) { return @{ Status = 'Keins'; Text = '' } }
    try { $p = [IO.File]::ReadAllText($pf, [Text.Encoding]::UTF8) | ConvertFrom-Json -ErrorAction Stop } catch { return @{ Status = 'Fehler'; Text = 'paket.json nicht lesbar' } }
    if ($p.PSObject.Properties['geprueft'] -and $p.geprueft -eq $true) { return @{ Status = 'Keins'; Text = '' } }
    $mf = Join-Path $cfg $script:HMManifestName; $sf = Join-Path $cfg $script:HMSignatureName
    $text = ''
    if (-not (Test-Path -LiteralPath $mf) -or -not (Test-Path -LiteralPath $sf)) { $text = 'Prüfsummen- oder Signatur-Datei fehlt im Paket.' }
    else {
        $man = [IO.File]::ReadAllBytes($mf)
        $why = Test-HMManifestSignature $man ([IO.File]::ReadAllBytes($sf)) $Thumbprint
        if ($why) { $text = "Signatur: $why" }
        else {
            $bad = @(Test-HUInstallFiles -AppRoot $AppRoot -Manifest $man)
            # ungepruefte Module duerfen in einem Paket nicht liegen (Modules\ und AddonsDev\ laden ohne Signatur)
            foreach ($d in 'Modules', 'AddonsDev') {
                $dp = Join-Path $AppRoot $d
                if ((Test-Path -LiteralPath $dp) -and @(Get-ChildItem -LiteralPath $dp -Recurse -File -Force -ErrorAction SilentlyContinue).Count) { $bad += "ungeprüfter Ordner: $d" }
            }
            if ($bad.Count) { $text = "$($bad.Count) Abweichungen: $((@($bad) | Select-Object -First 5) -join '; ')" }
        }
    }
    if ($text) { return @{ Status = 'Fehler'; Text = $text } }
    try { $p | Add-Member -NotePropertyName geprueft -NotePropertyValue $true -Force; [IO.File]::WriteAllText($pf, (ConvertTo-Json -InputObject $p -Depth 6), (New-Object Text.UTF8Encoding $true)) } catch { }
    return @{ Status = 'OK'; Text = "Paket $($p.name) vom $($p.erstellt) geprüft: Programm unverändert." }
}

# Beigelegte Zugangsdatei des Pakets (Zugang.humodule im Programmordner, Pruefsumme laut paket.json). '' = keine
function Get-HUPackageKeyFile([string]$AppRoot) {
    $kf = Join-Path $AppRoot $script:HUPackageKeyName
    if (-not (Test-Path -LiteralPath $kf -PathType Leaf)) { return '' }
    $pf = Join-Path (Join-Path $AppRoot 'Config') $script:HUPackageInfoName
    try { $p = [IO.File]::ReadAllText($pf, [Text.Encoding]::UTF8) | ConvertFrom-Json -ErrorAction Stop } catch { return '' }
    if (-not $p.PSObject.Properties['zugangSha256'] -or -not "$($p.zugangSha256)") { return '' }
    if ((Get-HMFileSha256 $kf) -ne "$($p.zugangSha256)") { return '' }
    return $kf
}
