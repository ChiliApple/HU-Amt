#Requires -Version 5.1
<#
.SYNOPSIS
    Update (GitHub-Releases, Kanal Stabil/Test, Pruefsumme + Signatur), Info-Fenster,
    Release signieren / freigeben (nur Verwalter mit Herausgeber-Zertifikat).
.DESCRIPTION
    Gleicher Ablauf wie in allen HU-Tools:
      - Start: Pruefung im Hintergrund, bei neuer Version wird der Knopf 'Update' gold - fuer alle Benutzer
      - Klick: zuerst pruefen, ob jemand anderer die App offen hat (Data\Laufend); dann Pull.ps1 (wartet auf das Ende,
        laedt + prueft ALLE Dateien, ersetzt erst dann, startet neu)
      - Rechtsklick: jetzt pruefen, Info; nur Verwalter: andere Version, Kanal, GitHub-Token, Module, signieren, freigeben
    Verwalter: wer in Data\Verwalter.txt steht (ein Windows-Benutzername je Zeile). Fehlt die Datei, nur wer das
    Herausgeber-Zertifikat hat.
    Einstellungen: Config\update.json (Kanal, Signaturpflicht, Quelle). Installierte Version: Config\installed.json.
.NOTES
    Dot-Source aus Main.ps1. Zielmaschine: der PC, auf dem die App laeuft.
#>

$script:ConfigDir = Join-Path $script:AppRoot 'Config'
$script:UpdateLib = Join-Path $script:AppRoot 'Functions\Core-Update.ps1'
$script:UpdateCfg = $null
$script:UpdateAvailable = $false
$script:UpdateRemoteVer = ''
$script:UpdateQuiet = $false
$script:SignCtx = $null; $script:PubCtx = $null; $script:UpdMenuAdmin = @(); $script:UpdMenuSign = @()
$script:SignTokFile = ''; $script:PubTokFile = ''

function Update-UpdateConfig {
    if (-not (Get-Command Get-HMUpdateConfig -ErrorAction SilentlyContinue)) { return $null }
    $script:UpdateCfg = Get-HMUpdateConfig $script:ConfigDir
    return $script:UpdateCfg
}
function Format-HUChannel([string]$Channel) { if ($Channel -eq 'Test') { 'Test' } elseif ($Channel -eq 'Branch') { 'Entwicklung (Branch)' } else { 'Stabil' } }

# --- Verwalter -----------------------------------------------------------------------------------
function Get-HUVerwalterFile { return (Join-Path $script:DataRoot 'Verwalter.txt') }
function Get-HUVerwalterList {
    $f = Get-HUVerwalterFile
    if (-not (Test-Path -LiteralPath $f)) { return $null }
    return @([IO.File]::ReadAllLines($f, [Text.Encoding]::UTF8) | ForEach-Object { "$_".Trim() } | Where-Object { $_ -and $_ -notlike '#*' })
}
function Test-HUVerwalter {
    $l = Get-HUVerwalterList
    if ($null -ne $l) { return [bool](@($l | Where-Object { $_ -ieq $script:UserName -or $_ -ieq "$env:USERDOMAIN\$($script:UserName)" }).Count) }
    return (Test-HUCanSign)
}

# --- GitHub-Token (nur private Repos; je Benutzer, DPAPI) ------------------------------------------
function Get-GitHubTokenFile { return (Join-Path (Join-Path $script:UserRoot 'Token') 'GitHub.xml') }
function Read-GitHubToken {
    if ("$env:HU_GITHUB_TOKEN".Trim()) { return "$env:HU_GITHUB_TOKEN".Trim() }
    $f = Get-GitHubTokenFile
    if (Test-Path -LiteralPath $f) {
        try { $c = Import-Clixml -Path $f -ErrorAction Stop; if ($c -is [System.Management.Automation.PSCredential]) { return $c.GetNetworkCredential().Password.Trim() } } catch { }
    }
    return ''
}
function Save-GitHubTokenFromPrompt([string]$File, [string]$Message) {
    $c = $null
    try { $c = Get-Credential -UserName 'github' -Message $Message } catch { $c = $null }
    if (-not $c -or -not $c.GetNetworkCredential().Password.Trim()) { Write-HULogWarn 'Kein Token eingegeben'; return $false }
    try {
        $d = Split-Path $File -Parent
        if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
        $c | Export-Clixml -Path $File -Force -ErrorAction Stop
        return $true
    } catch { Write-HULogError "Token nicht gespeichert: $($_.Exception.Message)"; return $false }
}
function Get-HUInstalledInfo { return (Read-HUJsonFile (Join-Path $script:ConfigDir 'installed.json')) }

# Kanal in Config\update.json setzen (gilt fuer diese Installation, im Netzordner fuer alle)
function Set-HUUpdateChannel([ValidateSet('Stable', 'Test')][string]$Channel) {
    $f = Join-Path $script:ConfigDir 'update.json'
    $u = Read-HUJsonFile $f
    if (-not $u) { $u = [pscustomobject]@{} }
    if ($u.PSObject.Properties['Channel']) { $u.Channel = $Channel } else { $u | Add-Member -NotePropertyName Channel -NotePropertyValue $Channel }
    Write-HUJsonFile -Path $f -Object $u -Depth 5
    [void](Update-UpdateConfig)
    Write-HULogOK "Update-Kanal: $(Format-HUChannel $Channel)"
    Update-VersionDisplay
}

function Set-HUUpdateButton([string]$RemoteVer = '', [bool]$Pre = $false) {
    $b = $script:Controls['btnUpdate']
    if ($RemoteVer) {
        $b.Content = "Update v$RemoteVer$(if ($Pre) { ' (Test)' })"
        $b.Background = [System.Windows.Media.Brushes]::Gold
        $b.Foreground = Get-HUBrush '#1E1E1E'
        $b.FontWeight = 'Bold'
    } else {
        $b.Content = 'Update'
        $b.ClearValue([System.Windows.Controls.Control]::BackgroundProperty)
        $b.ClearValue([System.Windows.Controls.Control]::ForegroundProperty)
        $b.ClearValue([System.Windows.Controls.Control]::FontWeightProperty)
    }
}

# Update-Pruefung im Hintergrund. -Quiet: nur bei neuer Version etwas protokollieren
function Invoke-UpdateCheck([switch]$Quiet) {
    $cfg = Update-UpdateConfig
    if (-not $cfg) { Write-HULogWarn 'Update-Pruefung: Functions\Core-Update.ps1 fehlt - Pull.ps1 ausführen'; return }
    $script:UpdateQuiet = [bool]$Quiet
    Invoke-AsyncCommand -ScriptBlock {
        param($lib, $token, $cfg, $ua)
        try {
            . $lib
            try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
            if ($cfg.UseBranch) {
                $h = @{ Accept = 'application/vnd.github.v3.raw'; 'User-Agent' = $ua }
                if ($token) { $h['Authorization'] = "token $token" }
                $r = Invoke-WebRequest "https://api.github.com/repos/$($cfg.Owner)/$($cfg.Repo)/contents/Config/version.json?ref=$($cfg.Branch)" -Headers $h -UseBasicParsing -TimeoutSec 20 -ErrorAction Stop
                $text = if ($r.Content -is [byte[]]) { [System.Text.Encoding]::UTF8.GetString($r.Content) } else { [string]$r.Content }
                $v = ($text | ConvertFrom-Json).version
                if ($v) { return "REL:$v|Branch $($cfg.Branch)|False" }
                return 'ERR:Version im Branch nicht gefunden'
            }
            $rel = Select-HMRelease @(Get-HMReleases $cfg.Owner $cfg.Repo $token) $cfg.Channel -SignedOnly:$cfg.RequireSignature
            if (-not $rel) { return 'NONE' }
            return "REL:$($rel.Version)|$($rel.Tag)|$($rel.Prerelease)"
        } catch {
            $code = 0; try { $code = [int]$_.Exception.Response.StatusCode } catch { }
            # 403 mit verbrauchtem Kontingent = Abfragelimit (ohne Token 60 je Stunde und Internetanschluss, auch bei oeffentlichen Repos)
            $left = ''; $resp = $_.Exception.Response
            try { $left = "$($resp.Headers['X-RateLimit-Remaining'])" } catch { }                                   # PS 5.1 (WebHeaderCollection)
            if (-not $left) { try { $left = "$(@($resp.Headers.GetValues('X-RateLimit-Remaining'))[0])" } catch { } }  # PS 7 (HttpResponseHeaders)
            if ($code -in 403, 429 -and $left -eq '0') { return "LIMIT:$code|$([bool]$token)" }
            if ($code -in 401, 403, 404) { return "AUTH:$code|$([bool]$token)" }
            return "ERR:$($_.Exception.Message)"
        }
    } -ArgumentList @($script:UpdateLib, (Read-GitHubToken), $cfg, "$($script:App.id)") -TimeoutSec 45 -OnComplete {
        param($result)
        $r = "$result".Trim()
        $cfg = $script:UpdateCfg
        $chan = if ($cfg.UseBranch) { "Branch $($cfg.Branch)" } else { "Kanal $(Format-HUChannel $cfg.Channel)" }
        if ($r -match '^REL:([^|]+)\|([^|]*)\|(True|False)$') {
            $remote = $Matches[1].Trim(); $pre = ($Matches[3] -eq 'True')
            $cmp = 0
            try { $cmp = ([Version]$remote).CompareTo([Version]$script:Version) } catch { $cmp = 0 }
            if ($cmp -gt 0) {
                $script:UpdateAvailable = $true; $script:UpdateRemoteVer = $remote
                Write-HULogWarn "Neue Version verfügbar: v$remote$(if ($pre) { ' (Test)' }) - installiert v$($script:Version). Goldener Knopf 'Update' oben rechts."
                Set-HUUpdateButton $remote $pre
            } else {
                $script:UpdateAvailable = $false
                if (-not $script:UpdateQuiet) { Write-HULogOK "Version aktuell: v$($script:Version) ($chan$(if ($cfg.RequireSignature) { ', nur signierte Updates' }))" }
                Set-HUUpdateButton
            }
        } elseif ($r -eq 'NONE') {
            if (-not $script:UpdateQuiet) { Write-HULogInfo "Update-Prüfung: noch keine $(if ($cfg.RequireSignature) { 'signierte ' })Version im $chan" }
        } elseif ($r -match '^LIMIT:(\d+)\|(True|False)$') {
            Write-HULogWarn "Update-Prüfung: GitHub-Abfragelimit erreicht (HTTP $($Matches[1])) - später erneut versuchen."
        } elseif ($r -match '^AUTH:(\d+)\|(True|False)$') {
            Write-HULogWarn "Update-Prüfung: $($cfg.Owner)/$($cfg.Repo) nicht erreichbar (HTTP $($Matches[1]))."
        } else {
            # auch beim stillen Start melden - sonst sieht man nie, warum kein Update angeboten wird
            Write-HULogWarn "Update-Prüfung fehlgeschlagen: $($r -replace '^(ERR:|FEHLER:)\s*', '')"
        }
    }
}

# Wer hat die App sonst noch offen? (gemeinsamer Netzordner) - Text fuer die Meldung oder ''
function Get-HUOtherSessionsText {
    $o = @(); try { $o = @(Get-HURunningInstances -RunDir $script:Paths.RunDir -Self $script:RunLock) } catch { }
    if (-not $o.Count) { return '' }
    return (($o | ForEach-Object { "  $($_.UserName) am PC $($_.ComputerName)$(if ($_.Since) { " (seit $($_.Since.ToString('HH:mm')))" })" }) -join "`n")
}

function Start-HUPull([string]$Version = '') {
    $pullScript = Join-Path $script:AppRoot 'Pull.ps1'
    if (-not (Test-Path -LiteralPath $pullScript)) { Write-HULogError "Pull.ps1 nicht gefunden: $pullScript"; return }
    $others = Get-HUOtherSessionsText
    if ($others) {
        Show-HUMessage "Das Update geht erst, wenn alle anderen $($script:AppTitle) geschlossen haben.`n`nBitte zuerst schließen:`n$others" -Icon Warning
        return
    }
    $cfg = Update-UpdateConfig
    $what = if ($Version) { "Version $Version" }
            elseif ($cfg -and $cfg.UseBranch) { "den Entwicklungsstand (Branch $($cfg.Branch), ohne Prüfsumme)" }
            elseif ($script:UpdateAvailable) { "v$($script:UpdateRemoteVer)" }
            else { "die neueste Version im Kanal $(Format-HUChannel $(if ($cfg) { $cfg.Channel } else { 'Stable' }))" }
    $check = if ($cfg -and $cfg.RequireSignature -and -not $cfg.UseBranch) { 'Prüfsumme und Signatur werden geprüft.' } elseif ($cfg -and $cfg.UseBranch) { 'OHNE Prüfung.' } else { 'Prüfsumme wird geprüft.' }
    if (-not (Confirm-HU "$($script:AppTitle) schließen, $what laden und neu starten?`n`n  Installiert:  v$($script:Version)`n  $check`n`nDaten und Einstellungen bleiben erhalten." 'Update')) { return }
    try {
        $exe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $argLine = "-NoProfile -ExecutionPolicy Bypass -File `"$pullScript`" -Target `"$($script:AppRoot)`" -WaitPid $PID"
        if ($Version) { $argLine += " -Version $Version" }
        Start-Process -FilePath $exe -ArgumentList $argLine -WorkingDirectory $script:AppRoot | Out-Null
        Write-HULogOK "Update gestartet - $($script:AppTitle) wird beendet und neu gestartet."
        $script:Window.Close()
    } catch { Write-HULogError "Update-Start fehlgeschlagen: $($_.Exception.Message)" }
}

# Einfaches Auswahlfenster mit Tabelle (Versionen). Rueckgabe: gewaehlte Zeile oder $null
function Show-HUTableDialog {
    param([string]$Title, [string]$Hint, [object[]]$Rows, [string]$OkText = 'OK', [int]$Width = 900, [int]$Height = 480)
    $x = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        WindowStartupLocation="CenterOwner" Background="#1E1E1E" ShowInTaskbar="False">
    <Window.Resources>
        <!--HU:THEME-->
    </Window.Resources>
    <DockPanel Margin="12">
        <TextBlock x:Name="lblHint" DockPanel.Dock="Top" Style="{StaticResource HintText}" FontSize="11" Margin="0,0,0,8" TextWrapping="Wrap"/>
        <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button x:Name="btnOk" Width="200" Background="#4CAF50" Style="{StaticResource DarkButton}" Margin="0,0,8,0"/>
            <Button x:Name="btnCancel" Content="Abbrechen" Width="100" Background="#555555" Style="{StaticResource DarkButton}" IsCancel="True"/>
        </StackPanel>
        <DataGrid x:Name="grid" Style="{StaticResource DarkDataGrid}" AutoGenerateColumns="True"/>
    </DockPanel>
</Window>
'@
    $d = New-HUWindow -XamlText (Get-HUThemedXaml $x)
    $w = $d.Window; $c = $d.C
    $w.Title = $Title; $w.Width = $Width; $w.Height = $Height
    $c.lblHint.Text = $Hint
    $c.btnOk.Content = $OkText
    $c.grid.ItemsSource = @($Rows)
    $st = @{ Sel = $null }
    $c.btnOk.Add_Click({ if ($c.grid.SelectedItem) { $st.Sel = $c.grid.SelectedItem; $w.Close() } })
    $c.grid.Add_MouseDoubleClick({ if ($c.grid.SelectedItem) { $st.Sel = $c.grid.SelectedItem; $w.Close() } })
    [void]$w.ShowDialog()
    return $st.Sel
}
# Theme an der Markierung <!--HU:THEME--> einsetzen
function Get-HUThemedXaml([string]$Xaml) {
    $theme = Get-HUXaml 'Theme'
    $m = [regex]::Match($theme, '(?s)<ResourceDictionary[^>]*>(.*)</ResourceDictionary>')
    return $Xaml.Replace('<!--HU:THEME-->', $m.Groups[1].Value)
}

function Show-HUVersionPicker {
    $cfg = Update-UpdateConfig
    if (-not $cfg) { Write-HULogError 'Functions\Core-Update.ps1 fehlt - Pull.ps1 ausführen'; return }
    Write-HULogInfo 'Versionen werden von GitHub gelesen ...'
    Invoke-AsyncCommand -ScriptBlock {
        param($lib, $token, $owner, $repo)
        try {
            . $lib
            $l = @(Get-HMReleases $owner $repo $token | ForEach-Object {
                [pscustomobject]@{ Version = "$($_.Version)"; Tag = $_.Tag; Prerelease = [bool]$_.Prerelease; Date = $_.Date; Notes = $_.Notes; Manifest = [bool]$_.ManifestUrl; Signature = [bool]$_.SignatureUrl }
            })
            return ('LIST:' + (ConvertTo-Json -InputObject @($l) -Compress -Depth 3))
        } catch {
            $code = 0; try { $code = [int]$_.Exception.Response.StatusCode } catch { }
            if ($code -in 401, 403, 404) { return "ERR:Repo nicht erreichbar (HTTP $code)" }
            return "ERR:$($_.Exception.Message)"
        }
    } -ArgumentList @($script:UpdateLib, (Read-GitHubToken), $cfg.Owner, $cfg.Repo) -TimeoutSec 40 -OnComplete {
        param($result)
        $r = "$result".Trim()
        if ($r -notmatch '^LIST:') { Write-HULogError "Versionen nicht lesbar: $($r -replace '^(ERR:|FEHLER:)\s*', '')"; return }
        $list = @()
        try { $list = @(($r.Substring(5) | ConvertFrom-Json) | Where-Object { $_ -and $_.Version }) } catch { Write-HULogError "Versionsliste nicht lesbar: $($_.Exception.Message)"; return }
        if (-not $list.Count) { Write-HULogWarn 'Keine Versionen gefunden.'; return }
        $rows = foreach ($x in $list) {
            $cmp = 0; try { $cmp = ([Version]"$($x.Version)").CompareTo([Version]$script:Version) } catch { }
            $first = @("$($x.Notes)" -split "`r?`n" | Where-Object { "$_".Trim() -and "$_" -notmatch '^\s*#' } | ForEach-Object { ("$_".Trim().TrimStart('-', ' ', '*') -replace '\*\*|`', '') })[0]
            [pscustomobject][ordered]@{
                Version = "$($x.Version)"; Kanal = $(if ($x.Prerelease) { 'Test' } else { 'Stabil' })
                Stand = $(if ($cmp -eq 0) { 'installiert' } elseif ($cmp -lt 0) { 'älter' } else { 'neuer' })
                Datum = "$($x.Date)"; Prüfsumme = $(if ($x.Manifest) { 'ja' } else { 'nein' }); Signatur = $(if ($x.Signature) { 'ja' } else { 'nein' }); Änderungen = "$first"
            }
        }
        $sel = Show-HUTableDialog -Title "$($script:AppTitle) - Version wählen" -OkText 'Diese Version installieren' -Rows @($rows) `
            -Hint "Installiert: v$($script:Version). Version markieren, dann installieren. Daten und Einstellungen bleiben erhalten.$(if ($script:UpdateCfg.RequireSignature) { ' Nur signierte Versionen sind installierbar.' })"
        if (-not $sel) { return }
        if ($script:UpdateCfg.RequireSignature -and $sel.Signatur -ne 'ja') { Show-HUMessage "Version $($sel.Version) ist nicht signiert und kann nicht installiert werden." -Icon Warning; return }
        Start-HUPull -Version $sel.Version
    }
}

# --- Herausgeber: Release signieren / freigeben (Zertifikat mit privatem Schluessel nur auf dessen PC) ---
# Schreib-Token: nie im Programmordner (auch nicht im Netzordner), sondern %APPDATA%\<App>\ (DPAPI)
function Get-HUSignTokenFile([string]$Owner = '', [string]$Repo = '') {
    $d = Join-Path $env:APPDATA "$($script:App.id)"
    if ($Repo) { return (Join-Path $d "GitHubSignToken-$Owner`_$Repo.xml") }
    return (Join-Path $d 'GitHubSignToken.xml')
}
function Read-HUSignToken([string]$Owner = '', [string]$Repo = '') {
    $f = Get-HUSignTokenFile $Owner $Repo
    if (Test-Path -LiteralPath $f) { try { $c = Import-Clixml -Path $f; if ($c -is [System.Management.Automation.PSCredential]) { return $c.GetNetworkCredential().Password.Trim() } } catch { } }
    return ''
}
function Test-HUCanSign {
    $c = Update-UpdateConfig
    if (-not $c -or -not (Get-Command Get-HMSigningCert -ErrorAction SilentlyContinue)) { return $false }
    return [bool](Get-HMSigningCert $c.SignerThumbprint)
}
function Get-HUSignTokenOrPrompt([object]$Cfg, [switch]$Addon) {
    $o = if ($Addon) { "$($Cfg.Owner)" } else { '' }; $r = if ($Addon) { "$($Cfg.Repo)" } else { '' }
    $tok = Read-HUSignToken $o $r
    if ($tok) { return $tok }
    if (-not (Save-GitHubTokenFromPrompt (Get-HUSignTokenFile $o $r) "GitHub-Token mit Schreibrecht für $($Cfg.Owner)/$($Cfg.Repo) (Fine-grained PAT, 'Contents: Read and write'). Wird verschlüsselt nur für deinen Windows-Benutzer gespeichert.")) { return '' }
    return (Read-HUSignToken $o $r)
}

# Modul-Repository zum Signieren/Freigeben waehlen (nur hinterlegte Module); $null = abgebrochen
function Select-HUAddonSignTarget([string]$What) {
    $src = @(Get-HUAddonSources)
    if (-not $src.Count) { Show-HUMessage 'Kein Modul hinterlegt (Rechtsklick auf Update > Module ...).' 'Module' -Icon Info; return $null }
    foreach ($s in $src) {
        if ($src.Count -eq 1 -or (Confirm-HU "$What für $($s.Owner)/$($s.Repo)?" 'Module')) { return $s }
    }
    return $null
}

function Write-HUPub([string]$Text, [ValidateSet('OK', 'WARN', 'ERROR', 'INFO')][string]$Level = 'INFO') {
    switch ($Level) { 'OK' { Write-HULogOK $Text } 'WARN' { Write-HULogWarn $Text } 'ERROR' { Write-HULogError $Text } default { Write-HULogInfo $Text } }
}

function Start-HUReleaseSigning([switch]$Addon) {
    $cfg = Update-UpdateConfig
    if (-not $cfg) { return }
    if ($Addon) {
        $t = Select-HUAddonSignTarget 'Release signieren'; if (-not $t) { return }
        $cfg = [pscustomobject]@{ Owner = $t.Owner; Repo = $t.Repo; SignerThumbprint = $cfg.SignerThumbprint }
    }
    Write-HUPub "=== Release signieren ($($cfg.Owner)/$($cfg.Repo)) ===" INFO
    $cert = Get-HMSigningCert $cfg.SignerThumbprint
    if (-not $cert) {
        $msg = "Signatur-Zertifikat $($cfg.SignerThumbprint) mit privatem Schlüssel ist auf diesem PC nicht vorhanden (oder abgelaufen)."
        Write-HUPub $msg ERROR; Show-HUMessage $msg 'Release signieren' -Icon Error
        return
    }
    Write-HUPub "Zertifikat: $($cert.Subject) - gültig bis $($cert.NotAfter.ToString('dd.MM.yyyy'))" INFO
    $tok = Get-HUSignTokenOrPrompt $cfg -Addon:$Addon
    if (-not $tok) { Write-HUPub 'Abgebrochen: kein GitHub-Token mit Schreibrecht.' WARN; return }
    $script:SignCtx = [pscustomobject]@{ Owner = $cfg.Owner; Repo = $cfg.Repo; Thumb = $cfg.SignerThumbprint; Token = $tok; Addon = [bool]$Addon }
    Write-HUPub 'Releases werden gelesen ...' INFO
    Invoke-AsyncCommand -ScriptBlock {
        param($lib, $tok, $owner, $repo)
        try {
            . $lib
            $all = @(Get-HMReleases $owner $repo $tok)
            $l = @($all | Where-Object { $_.ManifestUrl -and -not $_.SignatureUrl } | ForEach-Object { "$($_.Tag)" })
            $noMan = @($all | Where-Object { -not $_.ManifestUrl } | Select-Object -First 5 | ForEach-Object { "$($_.Tag)" })
            return ('TAGS:' + ($l -join ',') + '|' + ($noMan -join ','))
        } catch { return "ERR:$($_.Exception.Message)" }
    } -ArgumentList @($script:UpdateLib, $tok, $cfg.Owner, $cfg.Repo) -TimeoutSec 40 -OnComplete {
        param($result)
        $r = "$result".Trim()
        if ($r -notmatch '^TAGS:([^|]*)\|([^|]*)$') {
            $msg = "Releases nicht lesbar: $($r -replace '^(ERR:|FEHLER:)\s*', '')"
            Write-HUPub $msg ERROR; Show-HUMessage $msg 'Release signieren' -Icon Error
            $script:SignCtx = $null; return
        }
        $tags = @($Matches[1] -split ',' | Where-Object { $_ }); $noMan = "$($Matches[2])"
        if ($noMan) { Write-HUPub "Ohne Prüfsummen-Datei (automatische Tests noch nicht fertig oder fehlgeschlagen): $noMan" WARN }
        if (-not $tags.Count) {
            Write-HUPub 'Nichts zu signieren - alle Releases mit Prüfsumme sind bereits signiert.' OK
            Show-HUMessage "Nichts zu signieren - alle Releases mit Prüfsumme sind bereits signiert.$(if ($noMan) { "`n`nOhne Prüfsumme: $noMan (ein paar Minuten warten, bis die automatischen Tests fertig sind)" })" 'Release signieren'
            $script:SignCtx = $null; return
        }
        if (-not (Confirm-HU -Warning "Diese Releases jetzt signieren?`n`n$($tags -join ', ')`n`nDanach werden sie allen Installationen angeboten (je nach Kanal Stabil/Test)." 'Release signieren')) { Write-HUPub 'Signieren abgebrochen.' WARN; $script:SignCtx = $null; return }
        Write-HUPub "Signiere $($tags -join ', ') ..." INFO
        $sc = $script:SignCtx
        $script:SignTokFile = if ($sc.Addon) { Get-HUSignTokenFile $sc.Owner $sc.Repo } else { Get-HUSignTokenFile }
        Invoke-AsyncCommand -ScriptBlock {
            param($lib, $owner, $repo, $tp, $tok, $tags)
            try {
                . $lib
                $items = @(Invoke-HMReleaseSigning $owner $repo $tp $tok $tags | ForEach-Object { [pscustomobject]@{ Tag = "$($_.Tag)"; Ok = [bool]$_.Ok; Text = "$($_.Text)" } })
                return ('RES:' + (ConvertTo-Json -Depth 10 -InputObject @($items) -Compress))
            } catch { return "ERR:$($_.Exception.Message)" }
        } -ArgumentList @($script:UpdateLib, $sc.Owner, $sc.Repo, $sc.Thumb, $sc.Token, [string[]]$tags) -TimeoutSec 180 -OnComplete {
            param($res)
            $script:SignCtx = $null
            $s = "$res".Trim()
            if ($s -notmatch '^RES:') {
                $msg = "Signieren fehlgeschlagen: $($s -replace '^(ERR:|FEHLER:)\s*', '')"
                Write-HUPub $msg ERROR; Show-HUMessage $msg 'Release signieren' -Icon Error; return
            }
            $items = @(($s.Substring(4) | ConvertFrom-Json) | Where-Object { $_ })
            foreach ($x in $items) { if ($x.Ok) { Write-HUPub "$($x.Tag): $($x.Text)" OK } else { Write-HUPub "$($x.Tag): $($x.Text)" ERROR } }
            $bad = @($items | Where-Object { -not $_.Ok })
            $sum = ($items | ForEach-Object { "$($_.Tag): $($_.Text)" }) -join "`n"
            if ($bad.Count) { Show-HUMessage "Signieren mit Fehlern:`n`n$sum" 'Release signieren' -Icon Error }
            else { Show-HUMessage "Signiert:`n`n$sum`n`nNächster Schritt: Rechtsklick auf Update > Release freigeben." 'Release signieren' }
            if (@($items | Where-Object { -not $_.Ok -and "$($_.Text)" -match 'Schreibrecht' }).Count) {
                Remove-Item -LiteralPath $script:SignTokFile -Force -ErrorAction SilentlyContinue
                Write-HUPub 'Gespeicherter Schreib-Token gelöscht - beim nächsten Signieren neu eingeben.' WARN
            }
            Invoke-UpdateCheck
        }
    }
}

function Start-HUReleasePublish([switch]$Addon) {
    $cfg = Update-UpdateConfig
    if (-not $cfg) { return }
    if ($Addon) {
        $t = Select-HUAddonSignTarget 'Release freigeben'; if (-not $t) { return }
        $cfg = [pscustomobject]@{ Owner = $t.Owner; Repo = $t.Repo; SignerThumbprint = $cfg.SignerThumbprint }
    }
    Write-HUPub "=== Release freigeben ($($cfg.Owner)/$($cfg.Repo)) ===" INFO
    $tok = Get-HUSignTokenOrPrompt $cfg -Addon:$Addon
    if (-not $tok) { Write-HUPub 'Abgebrochen: kein GitHub-Token mit Schreibrecht.' WARN; return }
    $script:PubCtx = [pscustomobject]@{ Owner = $cfg.Owner; Repo = $cfg.Repo; Thumb = $cfg.SignerThumbprint; Token = $tok; Addon = [bool]$Addon }
    $script:PubTokFile = if ($Addon) { Get-HUSignTokenFile $cfg.Owner $cfg.Repo } else { Get-HUSignTokenFile }
    Write-HUPub 'Signierte Vorab-Releases werden gesucht ...' INFO
    Invoke-AsyncCommand -ScriptBlock {
        param($lib, $tok, $owner, $repo)
        try {
            . $lib
            $l = @(Get-HMReleases $owner $repo $tok)
            $stable = @($l | Where-Object { -not $_.Prerelease } | Select-Object -First 1)[0]
            $cand = @($l | Where-Object { $_.Prerelease -and $_.ManifestUrl -and $_.SignatureUrl -and (-not $stable -or $_.Version -gt $stable.Version) } | Select-Object -First 1)[0]
            $unsigned = @($l | Where-Object { $_.Prerelease -and -not $_.SignatureUrl -and (-not $stable -or $_.Version -gt $stable.Version) } | ForEach-Object { $_.Tag })
            return "PUB:$(if ($cand) { $cand.Tag })|$(if ($stable) { $stable.Tag })|$($unsigned -join ',')"
        } catch { return "ERR:$($_.Exception.Message)" }
    } -ArgumentList @($script:UpdateLib, $tok, $cfg.Owner, $cfg.Repo) -TimeoutSec 40 -OnComplete {
        param($result)
        $r = "$result".Trim()
        if ($r -notmatch '^PUB:([^|]*)\|([^|]*)\|([^|]*)$') {
            $msg = "Releases nicht lesbar: $($r -replace '^(ERR:|FEHLER:)\s*', '')"
            Write-HUPub $msg ERROR; Show-HUMessage $msg 'Release freigeben' -Icon Error
            $script:PubCtx = $null; return
        }
        $tag = $Matches[1]; $stable = $Matches[2]; $uns = $Matches[3]
        Write-HUPub "Bisher freigegeben: $(if ($stable) { $stable } else { '-' })" INFO
        if (-not $tag) {
            $msg = "Kein signiertes Vorab-Release neuer als $(if ($stable) { $stable } else { '(keines)' }).$(if ($uns) { " Noch nicht signiert: $uns - zuerst 'Release signieren'." })"
            Write-HUPub $msg WARN; Show-HUMessage $msg 'Release freigeben' -Icon Warning
            $script:PubCtx = $null; return
        }
        if (-not (Confirm-HU -Warning "$tag jetzt freigeben?`n`nDanach ist es im Kanal Stabil die neueste Version und wird ALLEN Installationen als Update angeboten.`nBisher freigegeben: $(if ($stable) { $stable } else { '-' })" 'Release freigeben')) { Write-HUPub 'Freigabe abgebrochen.' WARN; $script:PubCtx = $null; return }
        Write-HUPub "Gebe $tag frei ... (Signatur wird vorher erneut geprüft)" INFO
        $pc = $script:PubCtx
        Invoke-AsyncCommand -ScriptBlock {
            param($lib, $owner, $repo, $tag, $tp, $tok)
            try { . $lib; Publish-HMRelease $owner $repo $tag $tp $tok; return "OK:$tag" }
            catch {
                $code = 0; try { $code = [int]$_.Exception.Response.StatusCode } catch { }
                if ($code -in 401, 403, 404) { return "AUTH:kein Schreibrecht (HTTP $code) - Token prüfen" }
                return "ERR:$($_.Exception.Message)"
            }
        } -ArgumentList @($script:UpdateLib, $pc.Owner, $pc.Repo, $tag, $pc.Thumb, $pc.Token) -TimeoutSec 60 -OnComplete {
            param($res)
            $script:PubCtx = $null
            $s = "$res".Trim()
            if ($s -match '^OK:(.+)$') { $t = $Matches[1]; Write-HUPub "${t}: freigegeben (Kanal Stabil)" OK; Show-HUMessage "$t ist freigegeben (Kanal Stabil) und wird allen Installationen als Update angeboten." 'Release freigeben'; Invoke-UpdateCheck }
            elseif ($s -match '^AUTH:(.+)$') { $msg = "Freigabe fehlgeschlagen: $($Matches[1])"; Write-HUPub $msg ERROR; Remove-Item -LiteralPath $script:PubTokFile -Force -ErrorAction SilentlyContinue; Show-HUMessage "$msg`n`nGespeicherter Token wurde gelöscht - beim nächsten Versuch neu eingeben." 'Release freigeben' -Icon Error }
            else { $msg = "Freigabe fehlgeschlagen: $($s -replace '^(ERR:|FEHLER:)\s*', '')"; Write-HUPub $msg ERROR; Show-HUMessage $msg 'Release freigeben' -Icon Error }
        }
    }
}

# --- Info-Fenster ---
function Show-HUAbout {
    $x = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Info" Width="500" MinWidth="500" SizeToContent="Height" WindowStartupLocation="CenterOwner" ResizeMode="CanResizeWithGrip" Background="#1E1E1E" ShowInTaskbar="False">
  <StackPanel Margin="20">
    <Image x:Name="img" Width="96" Height="96" RenderOptions.BitmapScalingMode="HighQuality"/>
    <TextBlock x:Name="ttl" FontSize="26" FontWeight="Bold" Foreground="White" HorizontalAlignment="Center" Margin="0,8,0,0"/>
    <TextBlock x:Name="sub" FontSize="13" Foreground="#AAAAAA" HorizontalAlignment="Center"/>
    <TextBlock x:Name="ver" FontSize="12" Foreground="#4FC3F7" HorizontalAlignment="Center" TextAlignment="Center" TextWrapping="Wrap" Margin="0,10,0,0"/>
    <TextBlock TextWrapping="Wrap" x:Name="src" FontSize="11" Foreground="#4FC3F7" HorizontalAlignment="Center" Margin="0,6,0,0" Cursor="Hand" TextDecorations="Underline" ToolTip="Projektseite öffnen"/>
    <TextBlock x:Name="lic" Text="Nutzungslizenz - siehe LICENSE" FontSize="11" Foreground="#4FC3F7" HorizontalAlignment="Center" Margin="0,2,0,0" Cursor="Hand" TextDecorations="Underline" ToolTip="Lizenz anzeigen"/>
  </StackPanel>
</Window>
'@
    $d = New-HUWindow -XamlText $x
    $w = $d.Window; $c = $d.C
    try { $big = Get-HUSplashIcon $script:AppRoot; $c.img.Source = $(if ($big) { $big } else { $script:AppIcon }) } catch { }
    $c.ttl.Text = $script:AppTitle; $c.sub.Text = "$($script:App.subtitle)"
    $inst = Get-HUInstalledInfo
    $uc = Update-UpdateConfig
    $chk = if ($inst -and $inst.PSObject.Properties['Check'] -and "$($inst.Check)") { "$($inst.Check)" } else { '' }
    $chan = if (-not $uc) { '?' } elseif ($uc.UseBranch) { "Branch $($uc.Branch)" } else { Format-HUChannel $uc.Channel }
    $c.ver.Text = "Version $($script:Version)  |  PowerShell $($PSVersionTable.PSVersion)`nKanal $chan$(if ($uc -and $uc.RequireSignature) { ', nur signierte Updates' })" + $(if ($chk) { "`ninstalliert $($inst.Date), geprüft: $chk" } else { '' })
    $repoUrl = "https://github.com/$($script:App.repo.owner)/$($script:App.repo.name)"
    if ($uc) { $repoUrl = "https://github.com/$($uc.Owner)/$($uc.Repo)" }
    $c.src.Text = $repoUrl -replace '^https://', ''
    $c.src.Add_MouseLeftButtonUp({ Open-HUUrl $repoUrl })
    $c.lic.Add_MouseLeftButtonUp({ $lf = Join-Path $script:AppRoot 'LICENSE'; if (Test-Path -LiteralPath $lf) { Start-Process notepad.exe -ArgumentList "`"$lf`"" } else { Open-HUUrl "$repoUrl/blob/main/LICENSE" } })
    [void]$w.ShowDialog()
}

function Register-HUUpdateHandlers {
    $b = $script:Controls['btnUpdate']
    $b.Add_Click({ Start-HUPull })
    $b.ToolTip = "Neue Version laden und installieren (Prüfsumme und Signatur). Gold = neue Version da.`nRechtsklick: jetzt prüfen, Info"
    $cm = New-Object System.Windows.Controls.ContextMenu
    $add = { param($h, $sb) $mi = New-Object System.Windows.Controls.MenuItem; $mi.Header = $h; $mi.Add_Click($sb); [void]$cm.Items.Add($mi); $mi }
    [void](& $add 'Jetzt nach Updates suchen' { Invoke-UpdateCheck; if (Test-HUVerwalter) { try { Start-HUAddonUpdate -Loud } catch { } } })
    [void](& $add 'Info ...' { Show-HUAbout })
    # ab hier nur Verwalter
    $sepA = New-Object System.Windows.Controls.Separator; [void]$cm.Items.Add($sepA)
    $miVer = & $add 'Andere Version / Vorversion installieren ...' { Show-HUVersionPicker }
    $miChan = & $add 'Kanal: Testversionen annehmen' { $cfg = Update-UpdateConfig; Set-HUUpdateChannel $(if ($cfg -and $cfg.Channel -eq 'Test') { 'Stable' } else { 'Test' }); Invoke-UpdateCheck }
    $miChan.IsCheckable = $true
    $miMod = & $add 'Module (private Repositories) ...' { Show-HUAddonDialog }
    $miTok = & $add 'GitHub-Token (nur privates Repo) ...' {
        $f = Get-GitHubTokenFile
        $has = Test-Path -LiteralPath $f
        $a = Confirm-HUYesNoCancel "Gespeicherter GitHub-Token: $(if ($has) { 'vorhanden' } else { 'keiner' })`n`nNur für PRIVATE Repos nötig (Fine-grained PAT, nur 'Contents: Read').`n`nJa = Token eingeben/ersetzen`nNein = gespeicherten Token löschen`nAbbrechen = nichts tun" 'GitHub-Token'
        if ($a -eq 'Yes') { if (Save-GitHubTokenFromPrompt $f 'GitHub-Token (Nur-Lese) als Kennwort eingeben (verschlüsselt, nur für deinen Benutzer lesbar)') { Write-HULogOK 'GitHub-Token gespeichert'; Invoke-UpdateCheck } }
        elseif ($a -eq 'No' -and $has) { if (Confirm-HU -Warning 'Gespeicherten GitHub-Token wirklich löschen?' 'GitHub-Token') { Remove-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue; Write-HULogOK 'GitHub-Token gelöscht' } }
    }
    $script:UpdMenuAdmin = @($sepA, $miVer, $miChan, $miMod, $miTok)
    # nur Herausgeber (Zertifikat mit privatem Schluessel auf diesem PC)
    $sep = New-Object System.Windows.Controls.Separator; [void]$cm.Items.Add($sep)
    $miSign = & $add 'Release signieren (Herausgeber) ...' { Start-HUReleaseSigning }
    $miPub = & $add 'Release freigeben (Herausgeber) ...' { Start-HUReleasePublish }
    $miASign = & $add 'Modul: Release signieren (Herausgeber) ...' { Start-HUReleaseSigning -Addon }
    $miAPub = & $add 'Modul: Release freigeben (Herausgeber) ...' { Start-HUReleasePublish -Addon }
    $script:UpdMenuSign = @($sep, $miSign, $miPub, $miASign, $miAPub)
    $script:UpdMenuChan = $miChan
    $cm.Add_Opened({
            $adm = Test-HUVerwalter
            $va = $(if ($adm) { 'Visible' } else { 'Collapsed' })
            foreach ($m in $script:UpdMenuAdmin) { $m.Visibility = $va }
            $vs = $(if ($adm -and (Test-HUCanSign)) { 'Visible' } else { 'Collapsed' })
            foreach ($m in $script:UpdMenuSign) { $m.Visibility = $vs }
            $cfg = Update-UpdateConfig
            $script:UpdMenuChan.IsChecked = [bool]($cfg -and $cfg.Channel -eq 'Test')
        })
    $b.ContextMenu = $cm
}
