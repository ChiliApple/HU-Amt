#Requires -Version 5.1
<#
.SYNOPSIS
    Zentrale und Aussenstelle: Art der Installation, freigeschaltete Module und das Paket (ZIP) fuer eine Aussenstelle.
.DESCRIPTION
    Config\installation.json (gilt fuer die ganze Installation, bleibt bei Updates erhalten):
        { "modus": "Zentrale" | "Aussenstelle", "name": "Kindergarten", "moduleAus": [ "OCR", "PDF" ] }
    Fehlt die Datei, gilt: Zentrale, alle Module aktiv.

    Paket fuer eine Aussenstelle (New-HUOutpostPackage): ZIP mit einem Ordner HU-Amt-<Name>\ darin
        - Programmdateien der eigenen Installation, geprueft gegen die signierte Pruefsummen-Datei des Releases
          (liegt danach als Config\HU-Amt-files.sha256 und .p7s im Paket, damit die Aussenstelle beim ersten Start pruefen kann)
        - gewaehlte Module aus Addons\ (je Modul mit eigener Signatur, vorher geprueft)
        - Data\Verwalter.txt, Config\installation.json, Config\installed.json, Config\update.json (nur Kanal), Config\paket.json
        - auf Wunsch je Modul Data\<Id>\einstellungen.json (z. B. Reise mit dem Schluessel fuer die Routenberechnung)
        - auf Wunsch Zugang.humodule mit den Nur-Lese-Tokens fuer die mitgelieferten Module (Updates in der Aussenstelle)
        - LIESMICH.txt
      Nie im Paket: Daten, Benutzer, Token-Dateien, Protokolle, Sicherungen der Zentrale.
.NOTES
    Dot-Source aus Main.ps1 (nach Core-Update.ps1 und Core-Addons.ps1). Laeuft unter Windows PowerShell 5.1 und PowerShell 7.
    Zielmaschine: der PC, auf dem die App laeuft.
#>

$script:HUInstConfigName  = 'installation.json'
$script:HUPackageInfoName = 'paket.json'
$script:HUPackageKeyName  = 'Zugang.humodule'

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

# ------------------------------------------------------------------ Pruefsummen-Datei des Programms ------------------------------------
<#
.SYNOPSIS
    Signierte Pruefsummen-Datei der installierten Programmversion holen und pruefen.
    Zuerst Config\HU-Amt-files.sha256 (+ .p7s, legt Pull.ps1 ab), sonst vom Release der installierten Version (oeffentlich, ohne Token).
    Rueckgabe: @{ Manifest = byte[]; Signature = byte[]; Source = 'lokal'|'GitHub' } oder wirft mit Grund.
#>
function Get-HUVerifiedInstallManifest {
    param([string]$ConfigDir, [string]$Version, [string]$Owner, [string]$Repo, [string]$Thumbprint)
    $mf = Join-Path $ConfigDir $script:HMManifestName; $sf = Join-Path $ConfigDir $script:HMSignatureName
    if ((Test-Path -LiteralPath $mf) -and (Test-Path -LiteralPath $sf)) {
        $man = [IO.File]::ReadAllBytes($mf); $sig = [IO.File]::ReadAllBytes($sf)
        if (-not (Test-HMManifestSignature $man $sig $Thumbprint)) { return @{ Manifest = $man; Signature = $sig; Source = 'lokal' } }
    }
    $v = ConvertTo-HMVersion $Version
    if (-not $v) { throw "Die installierte Version ist keine freigegebene Version ($Version). Ein Paket geht nur aus einer über Pull installierten Version." }
    $rel = @(Get-HMReleases $Owner $Repo '' | Where-Object { $_.Version -eq $v })[0]
    if (-not $rel) { throw "Zur installierten Version $v gibt es auf GitHub kein Release." }
    if (-not $rel.ManifestUrl -or -not $rel.SignatureUrl) { throw "Das Release $($rel.Tag) hat keine signierte Prüfsummen-Datei." }
    $man = Get-HMReleaseAsset $rel.ManifestUrl $rel.ManifestApi ''
    $sig = Get-HMReleaseAsset $rel.SignatureUrl $rel.SignatureApi ''
    $why = Test-HMManifestSignature $man $sig $Thumbprint
    if ($why) { throw "Prüfsummen-Datei von $($rel.Tag): $why" }
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

# Token zu einem Modul: Repository "<Besitzer>/HU-Amt-<Id>" bzw. "<Besitzer>/<Id>"
function Find-HUAddonSourceFor([string]$Id, [object[]]$Sources) {
    foreach ($s in @($Sources)) { if ("$($s.Repo)" -ieq $Id -or "$($s.Repo)" -like "*-$Id") { return $s } }
    return $null
}

# ------------------------------------------------------------------ Paket erstellen ----------------------------------------------------
<#
.SYNOPSIS
    ZIP fuer eine Aussenstelle erstellen (Aufbau siehe Kopf der Datei). Netz wird nicht gebraucht.
.PARAMETER Manifest / Signature
    Signierte Pruefsummen-Datei der installierten Programmversion (Get-HUVerifiedInstallManifest). Jede Programmdatei
    muss dazu passen, sonst entsteht kein Paket.
.PARAMETER Modules
    Je Modul @{ Id; Active; Settings } - Settings = Data\<Id>\einstellungen.json mitgeben.
.PARAMETER Sources
    Hinterlegte Nur-Lese-Tokens (Owner, Repo, Token) - nur die der mitgelieferten Module kommen in Zugang.humodule.
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
        [string]$CreatedBy = "$env:USERNAME"
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
    if (Test-Path -LiteralPath $zip) { throw "Die Datei gibt es schon: $zip" }
    $zipDir = Split-Path $zip -Parent
    if (-not (Test-Path -LiteralPath $zipDir)) { throw "Den Zielordner gibt es nicht: $zipDir" }

    # 1. Programmdateien pruefen (alle muessen zur signierten Pruefsummen-Datei passen)
    $bad = @(Test-HUInstallFiles -AppRoot $AppRoot -Manifest $Manifest)
    if ($bad.Count) { throw ("Die Programmdateien passen nicht zur freigegebenen Version, es wurde kein Paket erstellt:`n" + (($bad | Select-Object -First 8) -join "`n") + $(if ($bad.Count -gt 8) { "`n... und $($bad.Count - 8) weitere" })) }
    $man = ConvertFrom-HMManifest ([Text.Encoding]::UTF8.GetString($Manifest))

    $stageRoot = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtPaket-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
    $stage = Join-Path $stageRoot $folderName
    $utf8Bom = New-Object Text.UTF8Encoding $true
    $writeJson = { param($Path, $Obj) $d = Split-Path $Path -Parent; if (-not (Test-Path -LiteralPath $d)) { [void](New-Item -ItemType Directory -Path $d -Force) }; [IO.File]::WriteAllText($Path, (ConvertTo-Json -InputObject $Obj -Depth 6), $utf8Bom) }
    try {
        [void](New-Item -ItemType Directory -Path $stage -Force)
        foreach ($k in $man.Keys) {
            $dst = Join-Path $stage ($k -replace '/', '\')
            $dd = Split-Path $dst -Parent; if (-not (Test-Path -LiteralPath $dd)) { [void](New-Item -ItemType Directory -Path $dd -Force) }
            Copy-Item -LiteralPath (Join-Path $AppRoot ($k -replace '/', '\')) -Destination $dst -Force
            if ((Get-HMFileSha256 $dst) -ne $man[$k]) { throw "Kopie stimmt nicht: $k" }
            $res.Files++
        }
        $cfg = Join-Path $stage 'Config'
        if (-not (Test-Path -LiteralPath $cfg)) { [void](New-Item -ItemType Directory -Path $cfg -Force) }
        [IO.File]::WriteAllBytes((Join-Path $cfg $script:HMManifestName), $Manifest)
        if ($Signature -and $Signature.Length) { [IO.File]::WriteAllBytes((Join-Path $cfg $script:HMSignatureName), $Signature) }

        # 2. Module (jedes mit eigener signierter Pruefsumme)
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
            foreach ($f in @(Get-ChildItem -LiteralPath $src -Recurse -File -Force -ErrorAction Stop)) {
                $rel = $f.FullName.Substring($src.TrimEnd('\').Length + 1)
                $t = Join-Path $dst $rel
                $td = Split-Path $t -Parent; if (-not (Test-Path -LiteralPath $td)) { [void](New-Item -ItemType Directory -Path $td -Force) }
                Copy-Item -LiteralPath $f.FullName -Destination $t -Force
            }
            if ($Thumbprint) { $why = Test-HUAddonFiles $dst $Thumbprint; if ($why) { throw "Kopie des Moduls $id stimmt nicht: $why" } }
            $res.Modules += "$($info.Title) $($info.Version)$(if (-not $mo.Active) { ' (aus)' })"
            # Einstellungen des Moduls (nur die Datei einstellungen.json, keine Daten)
            if ($mo.Settings) {
                $sf = Join-Path (Join-Path $DataRoot $id) 'einstellungen.json'
                if (Test-Path -LiteralPath $sf -PathType Leaf) {
                    $sd = Join-Path (Join-Path $stage 'Data') $id
                    [void](New-Item -ItemType Directory -Path $sd -Force)
                    Copy-Item -LiteralPath $sf -Destination (Join-Path $sd 'einstellungen.json') -Force
                    $res.Settings += $info.Title
                } else { $res.Warnings += "$($info.Title): keine Einstellungen vorhanden, nichts mitgegeben." }
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
                else { $res.Warnings += "Für $($mo.Id) ist kein Token hinterlegt: Dieses Modul bekommt in der Außenstelle keine Updates." }
            }
            if ($entries.Count) {
                $kf = Join-Path $stage $script:HUPackageKeyName
                [IO.File]::WriteAllText($kf, (ConvertTo-Json -InputObject ([ordered]@{ module = @($entries) }) -Depth 5), (New-Object Text.UTF8Encoding $false))
                $keySha = Get-HMFileSha256 $kf
            }
        }
        & $writeJson (Join-Path $cfg $script:HUPackageInfoName) ([ordered]@{ name = $name; erstellt = (Get-Date).ToString('yyyy-MM-dd HH:mm'); von = "$CreatedBy"; version = "$Version"; zugang = $(if ($keySha) { $script:HUPackageKeyName } else { '' }); zugangSha256 = $keySha; geprueft = $false })

        # 5. Anleitung
        $liesmich = @(
            "HU-Amt für die Außenstelle $name"
            ''
            'Einrichten (einmal):'
            '1. Diesen Ordner an einen festen Platz kopieren, z. B. C:\HU-Amt.'
            '   Nicht in Downloads, nicht in OneDrive.'
            '2. Im Ordner "Start-HU-Amt.cmd" doppelklicken.'
            '3. Die Fragen nach dem Desktop-Symbol und der Taskleiste mit Ja beantworten.'
            ''
            'Danach startet HU-Amt über das Symbol am Desktop.'
            ''
            "Erstellt am $((Get-Date).ToString('dd.MM.yyyy')) mit HU-Amt $Version."
        ) -join "`r`n"
        [IO.File]::WriteAllText((Join-Path $stage 'LIESMICH.txt'), $liesmich + "`r`n", $utf8Bom)

        # 6. ZIP
        try { Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop } catch { }
        $tmpZip = "$zip.tmp"
        if (Test-Path -LiteralPath $tmpZip) { Remove-Item -LiteralPath $tmpZip -Force }
        [IO.Compression.ZipFile]::CreateFromDirectory($stageRoot, $tmpZip, [IO.Compression.CompressionLevel]::Optimal, $false)
        Move-Item -LiteralPath $tmpZip -Destination $zip
        $res.Path = $zip; $res.Folder = $folderName
        return $res
    } finally {
        if (Test-Path -LiteralPath $stageRoot) { Remove-Item -LiteralPath $stageRoot -Recurse -Force -ErrorAction SilentlyContinue }
        if (Test-Path -LiteralPath "$zip.tmp") { Remove-Item -LiteralPath "$zip.tmp" -Force -ErrorAction SilentlyContinue }
    }
}

# ------------------------------------------------------------------ Erster Start eines Pakets ------------------------------------------
# Programmdateien einmal gegen die mitgelieferte, signierte Pruefsummen-Datei pruefen. Rueckgabe: @{ Status = 'OK'|'Fehler'|'Keins'; Text }
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
            if ($bad.Count) { $text = "$($bad.Count) Programmdateien weichen ab: $((@($bad) | Select-Object -First 5) -join '; ')" }
        }
    }
    try { $p | Add-Member -NotePropertyName geprueft -NotePropertyValue $true -Force; [IO.File]::WriteAllText($pf, (ConvertTo-Json -InputObject $p -Depth 6), (New-Object Text.UTF8Encoding $true)) } catch { }
    if ($text) { return @{ Status = 'Fehler'; Text = $text } }
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
