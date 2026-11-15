#Requires -Version 5.1
# Pester 5: Update (Pull.ps1, Functions\Core-Update.ps1). Kein Netzzugriff; Sperren mit kuenstlichen Dateien unter TEMP.

BeforeAll {
    Set-StrictMode -Version 1.0   # wie Main.ps1: nicht gesetzte Variablen sind Fehler
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    $script:PullText = [IO.File]::ReadAllText((Join-Path $script:Root 'Pull.ps1'), [Text.Encoding]::UTF8)
    $script:LibText = [IO.File]::ReadAllText((Join-Path $script:Root 'Functions\Core-Update.ps1'), [Text.Encoding]::UTF8)
    function Get-Region([string]$Text) {
        $m = [regex]::Match($Text, '(?s)#region HMUpdateLib.*?#endregion HMUpdateLib')
        return ($m.Value -replace "`r`n", "`n")
    }
    # eine Funktion aus Pull.ps1 holen, ohne Pull auszufuehren
    $tok = $null; $err = $null
    $ast = [Management.Automation.Language.Parser]::ParseInput($script:PullText, [ref]$tok, [ref]$err)
    foreach ($fname in 'Get-HMOpenSessions', 'Get-HMReplaceBlock') {
        $fn = $ast.FindAll({ $args[0] -is [Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq $fname }, $true) | Select-Object -First 1
        if ($fn) { . ([scriptblock]::Create($fn.Extent.Text)) }   # fehlt sie, scheitert der zugehoerige Test (nicht alle)
    }
}

Describe 'Update-Bibliothek' {
    It 'HMUpdateLib ist in Pull.ps1 und Core-Update.ps1 wortgleich' {
        $a = Get-Region $script:PullText; $b = Get-Region $script:LibText
        $a.Length | Should -BeGreaterThan 1000
        $a | Should -BeExactly $b
    }
    It 'Konstanten passen zu Config\app.json' {
        $app = [IO.File]::ReadAllText((Join-Path $script:Root 'Config\app.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
        $r = Get-Region $script:LibText
        $r | Should -Match ([regex]::Escape("`$script:HMManifestName  = '$($app.id)-files.sha256'"))
        $r | Should -Match ([regex]::Escape("`$script:HMDefaultOwner  = '$($app.repo.owner)'"))
        $r | Should -Match ([regex]::Escape("`$script:HMDefaultRepo   = '$($app.repo.name)'"))
        $r | Should -Match ([regex]::Escape("`$script:HMDefaultSigner = '$($app.signer)'"))
    }
    It 'Pull laesst Data, Addons, AddonsDev, bin und Config in Ruhe' {
        $m = [regex]::Match($script:PullText, "\`$keepDirs = '([^']+)'")
        $m.Success | Should -BeTrue
        foreach ($d in 'Config/x.json', 'Data/Benutzer/a/ui-state.json', 'Addons/X/addon.json', 'AddonsDev/X/addon.json', 'bin/ocr.exe') { $d | Should -Match $m.Groups[1].Value }
        'Functions/UI-Common.ps1' | Should -Not -Match $m.Groups[1].Value
    }
}

Describe 'Update im gemeinsamen Netzordner' {
    BeforeEach {
        $script:App = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtPull-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        $script:Run = Join-Path $script:App 'Data\Laufend'
        [void](New-Item -ItemType Directory -Path $script:Run -Force)
    }
    AfterEach { Remove-Item -LiteralPath $script:App -Recurse -Force -ErrorAction SilentlyContinue }

    It 'ohne offene Sitzung: nichts gemeldet' { @(Get-HMOpenSessions $script:App).Count | Should -Be 0 }
    It 'verwaiste Sperre wird entfernt und nicht gemeldet' {
        Set-Content -LiteralPath (Join-Path $script:Run 'alt_PC-9.lock') -Value 'alt|PC-9|2026-01-01T08:00:00|1'
        @(Get-HMOpenSessions $script:App).Count | Should -Be 0
        Test-Path (Join-Path $script:Run 'alt_PC-9.lock') | Should -BeFalse
    }
    It 'offene Sitzung wird mit Name und PC gemeldet' {
        $p = Join-Path $script:Run 'anna_PC-1.lock'
        $fs = [IO.File]::Open($p, 'CreateNew', 'ReadWrite', 'Read')
        try {
            $b = [Text.Encoding]::UTF8.GetBytes('anna|PC-1|2026-11-03T08:00:00|1'); $fs.Write($b, 0, $b.Length); $fs.Flush()
            $r = @(Get-HMOpenSessions $script:App)
            $r.Count | Should -Be 1
            $r[0] | Should -Be 'anna am PC PC-1'
        } finally { $fs.Dispose() }
    }
    It 'Pull prueft die Sitzungen, bevor etwas geladen wird' {
        $iCheck = $script:PullText.IndexOf('$open = @(Get-HMOpenSessions $Target)')
        $iLoad = $script:PullText.IndexOf('# --- 1. Version waehlen')
        $iCheck | Should -BeGreaterThan 0
        $iCheck | Should -BeLessThan $iLoad
    }
    It 'vor dem Ersetzen: waehrend des Ladens geoeffnet -> Meldung, vorbereitete Dateien entfernt (K10)' {
        $t1 = Join-Path $script:App 'Main.ps1.pulltmp'; $t2 = Join-Path $script:App 'Functions\X.ps1.pulltmp'
        [void](New-Item -ItemType Directory -Path (Join-Path $script:App 'Functions') -Force)
        Set-Content -LiteralPath $t1 -Value 'neu'; Set-Content -LiteralPath $t2 -Value 'neu'
        $staged = New-Object System.Collections.Generic.List[object]
        $staged.Add(@($t1, (Join-Path $script:App 'Main.ps1'))); $staged.Add(@($t2, (Join-Path $script:App 'Functions\X.ps1')))
        Set-Content -LiteralPath (Join-Path $script:Run 'alt_PC-9.lock') -Value 'alt|PC-9|2026-01-01T08:00:00|1'   # verwaist, zaehlt nicht
        $fs = [IO.File]::Open((Join-Path $script:Run 'berta_PC-2.lock'), 'CreateNew', 'ReadWrite', 'Read')
        try {
            $b = [Text.Encoding]::UTF8.GetBytes('berta|PC-2|2026-11-15T08:00:00|1'); $fs.Write($b, 0, $b.Length); $fs.Flush()
            $why = Get-HMReplaceBlock $script:App $staged
            $why | Should -Match 'berta am PC PC-2'
            $why | Should -Match 'Es wurde NICHTS veraendert\.$'
            $why | Should -Not -Match 'alt'
            Test-Path -LiteralPath $t1 | Should -BeFalse
            Test-Path -LiteralPath $t2 | Should -BeFalse
        } finally { $fs.Dispose() }
    }
    It 'vor dem Ersetzen: niemand offen -> keine Meldung, vorbereitete Dateien bleiben (K10)' {
        $t1 = Join-Path $script:App 'Main.ps1.pulltmp'; Set-Content -LiteralPath $t1 -Value 'neu'
        $staged = New-Object System.Collections.Generic.List[object]; $staged.Add(@($t1, (Join-Path $script:App 'Main.ps1')))
        Get-HMReplaceBlock $script:App $staged | Should -BeExactly ''
        Test-Path -LiteralPath $t1 | Should -BeTrue
    }
    It 'Pull prueft die Sitzungen unmittelbar vor dem Ersetzen noch einmal (nach dem Laden, vor dem Journal)' {
        $iFail = $script:PullText.IndexOf('Stop-HMPull "$fail Datei(en) nicht geladen')
        $iCheck = $script:PullText.IndexOf('$why = Get-HMReplaceBlock $Target $staged')
        $iJournal = $script:PullText.IndexOf('$journal = Join-Path $cfgDir')
        $iFail | Should -BeGreaterThan 0
        $iCheck | Should -BeGreaterThan $iFail
        $iCheck | Should -BeLessThan $iJournal
        $script:PullText.Substring($iCheck, $iJournal - $iCheck) | Should -Match 'if \(\$why\) \{ Stop-HMPull \$why \}'
    }
}

Describe 'Alle Releases signieren und freigeben: was offen ist (Get-HMReleasePlan, ab 1.0.5)' {
    BeforeAll {
        . (Join-Path $script:Root 'Functions\Core-Update.ps1')
        function New-R([string]$Tag, [bool]$Pre, [bool]$Man, [bool]$Sig) {
            [pscustomobject]@{ Version = [version]($Tag.TrimStart('v')); Tag = $Tag; Prerelease = $Pre; ManifestUrl = $(if ($Man) { 'm' } else { '' }); SignatureUrl = $(if ($Sig) { 's' } else { '' }) }
        }
    }
    It 'neues Vorab-Release ohne Signatur: signieren und freigeben' {
        $p = Get-HMReleasePlan @((New-R 'v1.0.5' $true $true $false), (New-R 'v1.0.4' $false $true $true))
        $p.ToSign | Should -Be @('v1.0.5')
        $p.Publish | Should -Be 'v1.0.5'
        $p.Stable | Should -Be 'v1.0.4'
    }
    It 'schon signiertes Vorab-Release: nur freigeben; mehrere: das neueste' {
        $p = Get-HMReleasePlan @((New-R 'v1.0.3' $true $true $true), (New-R 'v1.0.4' $true $true $true), (New-R 'v1.0.2' $false $true $true))
        @($p.ToSign).Count | Should -Be 0
        $p.Publish | Should -Be 'v1.0.4'
    }
    It 'alles freigegeben: nichts offen; Release ohne Pruefsummen-Datei wird gemeldet, nicht freigegeben' {
        $p = Get-HMReleasePlan @((New-R 'v1.0.4' $false $true $true), (New-R 'v1.0.3' $true $true $true))
        @($p.ToSign).Count | Should -Be 0
        $p.Publish | Should -Be ''
        $q = Get-HMReleasePlan @((New-R 'v1.0.5' $true $false $false), (New-R 'v1.0.4' $false $true $true))
        $q.Publish | Should -Be ''
        $q.NoManifest | Should -Be @('v1.0.5')
    }
    It 'ohne freigegebenes Release: neuestes Vorab-Release ist der Kandidat' {
        (Get-HMReleasePlan @((New-R 'v0.1.0' $true $true $false))).Publish | Should -Be 'v0.1.0'
        (Get-HMReleasePlan @()).Publish | Should -Be ''
    }
}

Describe 'Update-Knopf: ein Tooltip (KD1, ab 1.0.8)' {
    It 'MainWindow.xaml: btnUpdate hat den Tooltip mit Zeilenumbruch' {
        [xml]$x = [IO.File]::ReadAllText((Join-Path $script:Root 'XAML\MainWindow.xaml'), [Text.Encoding]::UTF8)
        $ns = New-Object Xml.XmlNamespaceManager $x.NameTable
        $ns.AddNamespace('x', 'http://schemas.microsoft.com/winfx/2006/xaml')
        $btn = $x.SelectSingleNode("//*[@x:Name='btnUpdate']", $ns)
        $btn | Should -Not -BeNullOrEmpty
        $want = "Update: l$([char]0xE4)dt eine neue Version und installiert sie (Pr$([char]0xFC)fsumme und Signatur). Gold = neue Version da. Ohne neue Version meldet der Knopf nur, dass alles aktuell ist.`nRechtsklick: jetzt pr$([char]0xFC)fen, Info. F$([char]0xFC)r Verwalter: andere Version, Kanal, Module."
        $btn.GetAttribute('ToolTip') | Should -BeExactly $want
    }
    It 'UI-Update.ps1 setzt keinen zweiten Tooltip an btnUpdate' {
        $t = [IO.File]::ReadAllText((Join-Path $script:Root 'Functions\UI-Update.ps1'), [Text.Encoding]::UTF8)
        $t | Should -Not -Match '\$b\.ToolTip\s*='
        $t | Should -Not -Match "btnUpdate'\]\.ToolTip"
    }
}

Describe 'Pull -NoDowngrade: keine Rueckstufung, wenn die App es verlangt (ab 1.0.8)' {
    BeforeAll {
        . (Join-Path $script:Root 'Functions\Core-Update.ps1')
        $script:Tmp = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtVer-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        [void](New-Item -ItemType Directory -Path $script:Tmp -Force)
    }
    AfterAll { Remove-Item -LiteralPath $script:Tmp -Recurse -Force -ErrorAction SilentlyContinue }

    It 'installierte Version aus Config\version.json; fehlt oder kaputt = $null' {
        [IO.File]::WriteAllText((Join-Path $script:Tmp 'version.json'), '{ "version": "1.0.8", "buildDate": "2026-11-15" }')
        Get-HMInstalledVersion $script:Tmp | Should -Be ([version]'1.0.8')
        [IO.File]::WriteAllText((Join-Path $script:Tmp 'version.json'), '{ "version": ')
        Get-HMInstalledVersion $script:Tmp | Should -BeNullOrEmpty
        [IO.File]::WriteAllText((Join-Path $script:Tmp 'version.json'), '{ "buildDate": "2026-11-15" }')
        Get-HMInstalledVersion $script:Tmp | Should -BeNullOrEmpty
        Remove-Item -LiteralPath (Join-Path $script:Tmp 'version.json') -Force
        Get-HMInstalledVersion $script:Tmp | Should -BeNullOrEmpty
    }
    It 'nicht hoeher = ablehnen; hoeher oder installierte Version unbekannt = installieren' {
        Test-HMNoUpgrade ([version]'1.0.8') ([version]'1.0.8') | Should -BeTrue
        Test-HMNoUpgrade ([version]'1.0.7') ([version]'1.0.8') | Should -BeTrue
        Test-HMNoUpgrade ([version]'1.0.8.0') ([version]'1.0.8') | Should -BeTrue
        Test-HMNoUpgrade ([version]'1.0.9') ([version]'1.0.8') | Should -BeFalse
        Test-HMNoUpgrade ([version]'1.0.10') ([version]'1.0.9') | Should -BeFalse
        Test-HMNoUpgrade ([version]'1.0.9') $null | Should -BeFalse
    }
    It 'Pull.ps1: Schalter vorhanden, Ablehnung nach der Auswahl und vor dem Laden, startet die App wieder' {
        $script:PullText | Should -Match '\[switch\]\$NoDowngrade'
        $iSel = $script:PullText.IndexOf('# --- 1. Version waehlen')
        $iChk = $script:PullText.IndexOf('if ($NoDowngrade -and -not $useBranch -and -not $Version)')
        $iLoad = $script:PullText.IndexOf('# --- 2. Pruefsumme und Signatur')
        $iChk | Should -BeGreaterThan $iSel
        $iChk | Should -BeLessThan $iLoad
        $block = $script:PullText.Substring($iChk, $iLoad - $iChk)
        $block | Should -Match 'Test-HMNoUpgrade \$rel\.Version'
        $block | Should -Match 'Start-HMApp \$Target'
        $block | Should -Not -Match 'Stop-HMPull'
        $block | Should -Match 'NICHTS veraendert'
        # derselbe Neustart wie nach einem erfolgreichen Update
        ([regex]::Matches($script:PullText, 'Start-HMApp \$Target')).Count | Should -Be 2
    }
}

Describe 'Update-Knopf: Klick entscheidet nach der letzten Pruefung (K4, ab 1.0.8)' {
    BeforeAll {
        # kuenstlicher Programmordner unter TEMP mit einem Platzhalter-Pull.ps1; Start-Process ist immer gemockt,
        # es wird nie ein echtes Pull.ps1 gestartet
        $script:AppRoot = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtUpd-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        [void](New-Item -ItemType Directory -Path (Join-Path $script:AppRoot 'Config') -Force)
        $script:UpdJson = Join-Path $script:AppRoot 'Config\update.json'
        $script:UpdJsonStd = '{ "SignerThumbprint": "' + ('A' * 40) + '" }'
        [IO.File]::WriteAllText($script:UpdJson, $script:UpdJsonStd)
        $script:PullDummy = Join-Path $script:AppRoot 'Pull.ps1'
        [IO.File]::WriteAllText($script:PullDummy, "param([string]`$Target, [int]`$WaitPid = 0, [string]`$Version = '', [switch]`$NoDowngrade)`r`nexit 0`r`n")
        $script:App = [pscustomobject]@{ id = 'HU-Amt'; title = 'HU-Amt'; signer = ('A' * 40); repo = [pscustomobject]@{ owner = 'x'; name = 'y' } }
        $script:AppTitle = 'HU-Amt'; $script:Version = '1.0.8'
        $script:DataRoot = $script:AppRoot; $script:UserRoot = $script:AppRoot; $script:UserName = 'anna'
        $script:Msgs = New-Object System.Collections.Generic.List[string]
        $script:Asked = New-Object System.Collections.Generic.List[string]
        $script:Answer = $false
        function Write-HULogInfo { param($m) }
        function Write-HULogWarn { param($m) }
        function Write-HULogOK { param($m) }
        function Write-HULogError { param($m) }
        function Read-HUJsonFile([string]$Path) { if (Test-Path -LiteralPath $Path) { return ([IO.File]::ReadAllText($Path) | ConvertFrom-Json) } return $null }
        function Show-HUMessage { param([string]$Text, [string]$Title, [string]$Icon = 'Info', $Owner = $null) $script:Msgs.Add("$Icon|$Text") }
        function Confirm-HU { param([string]$Text, [string]$Title, [switch]$Warning, $Owner = $null) $script:Asked.Add($Text); return $script:Answer }
        . (Join-Path $script:Root 'Functions\Core-Update.ps1')
        . (Join-Path $script:Root 'Functions\UI-Update.ps1')
        # Ergebnis einer Pruefung so setzen, wie Invoke-UpdateCheck es nach der Antwort von GitHub tut
        function Set-TestCheck([string]$State, [string]$Remote = '', [string]$Err = '') {
            $script:UpdateCheck = [pscustomobject]@{ State = $State; Remote = $Remote; Error = $Err; Key = (Get-HUUpdateCfgKey (Update-UpdateConfig)); Since = (Get-Date) }
        }
    }
    AfterAll { Remove-Item -LiteralPath $script:AppRoot -Recurse -Force -ErrorAction SilentlyContinue }
    BeforeEach {
        $script:Msgs.Clear(); $script:Asked.Clear(); $script:Answer = $false
        $script:IsVerwalter = $false
        $script:UpdateClickPending = $false
        $script:Window = [pscustomobject]@{ Closed = 0 } | Add-Member -MemberType ScriptMethod -Name Close -Value { $this.Closed++ } -PassThru
        [IO.File]::WriteAllText($script:UpdJson, $script:UpdJsonStd)
        Mock Start-Process { }
        Mock Invoke-UpdateCheck { }
    }

    Context 'Entscheidung ohne Fenster (Get-HUUpdateAction)' {
        It 'hoeher = install, gleich = current, niedriger = newer (Zahlen, nicht Text)' {
            Get-HUUpdateAction '1.0.8' '1.0.9' 'ok' | Should -Be 'install'
            Get-HUUpdateAction '1.0.8' '1.0.10' 'ok' | Should -Be 'install'
            Get-HUUpdateAction '1.0.8' '1.0.8' 'ok' | Should -Be 'current'
            Get-HUUpdateAction '1.0.8' 'v1.0.8.0' 'ok' | Should -Be 'current'
            Get-HUUpdateAction '1.0.8' '1.0.7' 'ok' | Should -Be 'newer'
            Get-HUUpdateAction '1.0.10' '1.0.9' 'ok' | Should -Be 'newer'
        }
        It 'keine Pruefung, Pruefung laeuft oder fehlgeschlagen = check' {
            foreach ($s in 'none', 'running', 'failed', '') { Get-HUUpdateAction '1.0.8' '1.0.9' $s | Should -Be 'check' }
        }
        It 'keine Version im Kanal = none' {
            Get-HUUpdateAction '1.0.8' '' 'empty' | Should -Be 'none'
        }
        It 'Antwort ohne lesbare Version = check; installierte Version unlesbar = install (wie bisher)' {
            Get-HUUpdateAction '1.0.8' 'kaputt' 'ok' | Should -Be 'check'
            Get-HUUpdateAction '' '1.0.9' 'ok' | Should -Be 'install'
        }
    }

    Context 'Klick auf den Knopf (Start-HUPull ohne -Version)' {
        It 'ohne Pruefung in dieser Sitzung: es wird zuerst geprueft, nichts gefragt und nichts installiert' {
            $script:UpdateCheck = [pscustomobject]@{ State = 'none'; Remote = ''; Error = ''; Key = ''; Since = [datetime]::MinValue }
            Start-HUPull
            $script:Asked.Count | Should -Be 0
            Should -Invoke Invoke-UpdateCheck -Times 1 -Exactly
            Should -Invoke Start-Process -Times 0 -Exactly
        }
        It 'zweiter Klick waehrend der Pruefung startet keinen zweiten Ablauf' {
            $script:UpdateCheck = [pscustomobject]@{ State = 'none'; Remote = ''; Error = ''; Key = ''; Since = [datetime]::MinValue }
            Start-HUPull; Start-HUPull
            Should -Invoke Invoke-UpdateCheck -Times 1 -Exactly
            $script:Asked.Count | Should -Be 0
        }
        It 'gleiche Version: nur Meldung "ist aktuell", nichts installiert' {
            Set-TestCheck 'ok' '1.0.8'
            Start-HUPull
            $script:Asked.Count | Should -Be 0
            $script:Msgs.Count | Should -Be 1
            $script:Msgs[0] | Should -BeExactly 'Info|HU-Amt ist aktuell: Version 1.0.8 (Kanal Stabil). Es gibt nichts zu laden.'
            Should -Invoke Start-Process -Times 0 -Exactly
            Should -Invoke Invoke-UpdateCheck -Times 0 -Exactly
        }
        It 'Verwalter bekommen den Hinweis auf den Rechtsklick' {
            $script:IsVerwalter = $true
            Set-TestCheck 'ok' '1.0.8'
            Start-HUPull
            $script:Msgs[0] | Should -BeLike '*Neu installieren oder eine andere Version: Rechtsklick auf Update, Andere Version / Vorversion installieren ...'
        }
        It 'installierte Version neuer als der Kanal: Meldung, keine Rueckstufung' {
            Set-TestCheck 'ok' '1.0.7'
            Start-HUPull
            $script:Asked.Count | Should -Be 0
            $script:Msgs[0] | Should -BeExactly 'Info|Installiert ist Version 1.0.8, im Kanal Stabil ist 1.0.7 die neueste. Es wird nichts geändert.'
            Should -Invoke Start-Process -Times 0 -Exactly
        }
        It 'keine Version im Kanal: Meldung, nichts installiert' {
            Set-TestCheck 'empty'
            Start-HUPull
            $script:Asked.Count | Should -Be 0
            $script:Msgs[0] | Should -BeLike 'Info|Im Kanal Stabil gibt es noch keine*Version. Es wird nichts geändert.'
            Should -Invoke Start-Process -Times 0 -Exactly
        }
        It 'neue Version: Rueckfrage wie bisher (abgelehnt = nichts gestartet)' {
            Set-TestCheck 'ok' '1.0.9'
            Start-HUPull
            $script:Asked.Count | Should -Be 1
            $script:Asked[0] | Should -BeLike '*v1.0.9 laden*'
            Should -Invoke Start-Process -Times 0 -Exactly
        }
        It 'fehlgeschlagene Pruefung: Prueft beim Klick zuerst; scheitert sie, klare Meldung und nichts installiert' {
            Set-TestCheck 'failed' '' 'GitHub-Abfragelimit erreicht (HTTP 403)'
            Start-HUPull
            Should -Invoke Invoke-UpdateCheck -Times 1 -Exactly
            $script:UpdateClickPending | Should -BeTrue
            # Antwort kommt: wieder fehlgeschlagen
            Set-TestCheck 'failed' '' 'GitHub-Abfragelimit erreicht (HTTP 403)'
            Complete-HUUpdateClick $script:UpdateCheck.Key
            $script:UpdateClickPending | Should -BeFalse
            $script:Asked.Count | Should -Be 0
            $script:Msgs[0] | Should -BeExactly 'Warning|Update-Prüfung nicht möglich: GitHub-Abfragelimit erreicht (HTTP 403). Bitte später noch einmal versuchen.'
            Should -Invoke Start-Process -Times 0 -Exactly
        }
        It 'Pruefung nach dem Klick findet eine neue Version: Rueckfrage' {
            $script:UpdateCheck = [pscustomobject]@{ State = 'none'; Remote = ''; Error = ''; Key = ''; Since = [datetime]::MinValue }
            Start-HUPull
            Set-TestCheck 'ok' '1.0.9'
            Complete-HUUpdateClick $script:UpdateCheck.Key
            $script:Asked.Count | Should -Be 1
        }
        It 'Ergebnis von vor dem Kanalwechsel zaehlt nicht: es wird neu geprueft' {
            Set-TestCheck 'ok' '1.0.9'   # gilt fuer Kanal Stabil
            [IO.File]::WriteAllText($script:UpdJson, '{ "SignerThumbprint": "' + ('A' * 40) + '", "Channel": "Test" }')
            Start-HUPull
            $script:Asked.Count | Should -Be 0
            Should -Invoke Invoke-UpdateCheck -Times 1 -Exactly
        }
        It 'Rechtsklick-Weg (Start-HUPull -Version): ohne Vergleich, Rueckfrage wie bisher' {
            $script:UpdateCheck = [pscustomobject]@{ State = 'none'; Remote = ''; Error = ''; Key = ''; Since = [datetime]::MinValue }
            Start-HUPull -Version '1.0.5'
            $script:Asked.Count | Should -Be 1
            $script:Asked[0] | Should -BeLike '*Version 1.0.5 laden*'
            Should -Invoke Invoke-UpdateCheck -Times 0 -Exactly
        }
        It 'neue Version bestaetigt: Pull bekommt -NoDowngrade, mit -Version nicht' {
            $script:Answer = $true
            Set-TestCheck 'ok' '1.0.9'
            Start-HUPull
            Should -Invoke Start-Process -Times 1 -Exactly -ParameterFilter { "$ArgumentList" -match ' -NoDowngrade' -and "$ArgumentList" -notmatch ' -Version ' }
            Start-HUPull -Version '1.0.5'
            Should -Invoke Start-Process -Times 1 -Exactly -ParameterFilter { "$ArgumentList" -match ' -Version 1\.0\.5' -and "$ArgumentList" -notmatch 'NoDowngrade' }
        }
        It 'altes Pull.ps1 ohne den Schalter: -NoDowngrade wird nicht uebergeben' {
            $script:Answer = $true
            [IO.File]::WriteAllText($script:PullDummy, "param([string]`$Target, [int]`$WaitPid = 0, [string]`$Version = '')`r`nexit 0`r`n")
            try {
                Set-TestCheck 'ok' '1.0.9'
                Start-HUPull
                Should -Invoke Start-Process -Times 1 -Exactly -ParameterFilter { "$ArgumentList" -notmatch 'NoDowngrade' }
            } finally { [IO.File]::WriteAllText($script:PullDummy, "param([string]`$Target, [int]`$WaitPid = 0, [string]`$Version = '', [switch]`$NoDowngrade)`r`nexit 0`r`n") }
        }
        It 'Branch-Modus: unveraendert, Rueckfrage ohne vorherige Pruefung' {
            [IO.File]::WriteAllText($script:UpdJson, '{ "SignerThumbprint": "' + ('A' * 40) + '", "AllowUnsigned": true, "UseBranch": true, "Branch": "dev" }')
            $script:UpdateCheck = [pscustomobject]@{ State = 'none'; Remote = ''; Error = ''; Key = ''; Since = [datetime]::MinValue }
            Start-HUPull
            $script:Asked.Count | Should -Be 1
            $script:Asked[0] | Should -BeLike '*Branch dev*'
            Should -Invoke Invoke-UpdateCheck -Times 0 -Exactly
        }
    }

    Context 'Alter der Pruefung: aelter als 10 Minuten wird beim Klick erneuert' {
        BeforeAll {
            # Pruefergebnis mit Zeitpunkt; $Since = $null legt das Feld gar nicht an
            function Set-AgedCheck([string]$State, [string]$Remote, $Since, [switch]$NoSince) {
                $o = [ordered]@{ State = $State; Remote = $Remote; Error = ''; Key = (Get-HUUpdateCfgKey (Update-UpdateConfig)) }
                if (-not $NoSince) { $o['Since'] = $Since }
                $script:UpdateCheck = [pscustomobject]$o
            }
        }
        It 'vor 11 Minuten: jedes Ergebnis gilt nicht mehr (check)' {
            $old = (Get-Date).AddMinutes(-11)
            Set-AgedCheck 'ok' '1.0.8' $old; Get-HUUpdateClickAction | Should -Be 'check'
            Set-AgedCheck 'ok' '1.0.9' $old; Get-HUUpdateClickAction | Should -Be 'check'
            Set-AgedCheck 'ok' '1.0.7' $old; Get-HUUpdateClickAction | Should -Be 'check'
            Set-AgedCheck 'empty' '' $old; Get-HUUpdateClickAction | Should -Be 'check'
        }
        It 'vor 9 Minuten: Entscheidung wie bisher' {
            $young = (Get-Date).AddMinutes(-9)
            Set-AgedCheck 'ok' '1.0.8' $young; Get-HUUpdateClickAction | Should -Be 'current'
            Set-AgedCheck 'ok' '1.0.9' $young; Get-HUUpdateClickAction | Should -Be 'install'
            Set-AgedCheck 'ok' '1.0.7' $young; Get-HUUpdateClickAction | Should -Be 'newer'
            Set-AgedCheck 'empty' '' $young; Get-HUUpdateClickAction | Should -Be 'none'
        }
        It 'fehlendes, leeres oder unbrauchbares Since gilt als veraltet' {
            Set-AgedCheck 'ok' '1.0.8' $null -NoSince; Get-HUUpdateClickAction | Should -Be 'check'
            Set-AgedCheck 'ok' '1.0.8' $null; Get-HUUpdateClickAction | Should -Be 'check'
            Set-AgedCheck 'ok' '1.0.8' ([datetime]::MinValue); Get-HUUpdateClickAction | Should -Be 'check'
            Set-AgedCheck 'ok' '1.0.8' 'heute'; Get-HUUpdateClickAction | Should -Be 'check'
            Set-AgedCheck 'ok' '1.0.8' ((Get-Date).AddHours(1)); Get-HUUpdateClickAction | Should -Be 'check'   # Uhr zurueckgestellt
        }
        It 'Klick mit veraltetem "aktuell": zuerst pruefen, keine Meldung; die neue Antwort entscheidet' {
            Set-AgedCheck 'ok' '1.0.8' ((Get-Date).AddMinutes(-11))
            Start-HUPull
            Should -Invoke Invoke-UpdateCheck -Times 1 -Exactly
            $script:Msgs.Count | Should -Be 0
            $script:Asked.Count | Should -Be 0
            $script:UpdateClickPending | Should -BeTrue
            # inzwischen freigegeben: die neue Pruefung findet 1.0.9
            Set-TestCheck 'ok' '1.0.9'
            Complete-HUUpdateClick $script:UpdateCheck.Key
            $script:UpdateClickPending | Should -BeFalse
            $script:Msgs.Count | Should -Be 0
            $script:Asked.Count | Should -Be 1
            $script:Asked[0] | Should -BeLike '*v1.0.9 laden*'
            Should -Invoke Start-Process -Times 0 -Exactly
        }
        It 'Klick mit veralteter "neuer Version": zuerst pruefen, Rueckfrage erst nach der neuen Antwort' {
            Set-AgedCheck 'ok' '1.0.9' ((Get-Date).AddMinutes(-11))
            Start-HUPull
            Should -Invoke Invoke-UpdateCheck -Times 1 -Exactly
            $script:Asked.Count | Should -Be 0
            Set-TestCheck 'ok' '1.0.10'
            Complete-HUUpdateClick $script:UpdateCheck.Key
            $script:Asked.Count | Should -Be 1
            $script:Asked[0] | Should -BeLike '*v1.0.10 laden*'
        }
    }
}
