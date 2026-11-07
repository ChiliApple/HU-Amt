#Requires -Version 5.1
# Pester 5: Absender fuer Mails aus Modulen (Functions\UI-Mail.ps1). Ohne Outlook und ohne Fenster: Konten und Dialog sind Stellvertreter.

BeforeAll {
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Import-Module (Join-Path $script:Root 'Core\HU.Office.psm1') -Force -DisableNameChecking
    . (Join-Path $script:Root 'Functions\Core-TestMode.ps1')
    . (Join-Path $script:Root 'Functions\UI-Mail.ps1')
    function Get-HUStateValue([string]$Name, $Default = $null) { if ($script:State.ContainsKey($Name)) { return $script:State[$Name] } return $Default }
    function Set-HUStateValue([string]$Name, $Value) { $script:State[$Name] = $Value }
    function Save-HUUIState { }
    function Show-HUMessage { param($Text, $Icon, $Owner) }
}

Describe 'Absender-Angabe' {
    It 'leer, Text oder Objekt werden vereinheitlicht' {
        (ConvertTo-HUMailSender $null).Address | Should -Be ''
        (ConvertTo-HUMailSender ' amt@x.at ').Address | Should -Be 'amt@x.at'
        (ConvertTo-HUMailSender @{ Address = 'a@x.at'; Mode = 'imAuftrag' }).Mode | Should -Be 'imAuftrag'
        (ConvertTo-HUMailSender @{ Address = 'a@x.at'; Mode = 'quatsch' }).Mode | Should -Be 'auto'
    }
}

Describe 'Absender bestimmen' {
    BeforeEach {
        $script:State = @{}; $script:Asked = 0; $script:Answer = $null
        $script:Accs = @([pscustomobject]@{ Address = 'anna@x.at'; Name = 'Anna'; Exchange = $true }, [pscustomobject]@{ Address = 'amt@x.at'; Name = 'Amt'; Exchange = $true })
        Mock Get-HUOutlookAccounts { $script:Accs }
        Mock Show-HUMailSenderDialog { $script:Asked++; $script:Answer }
    }
    It 'Einstellung des Moduls mit vorhandenem Konto: keine Frage' {
        $r = Resolve-HUMailSender -ModuleId 'Buchung' -Sender @{ Address = 'amt@x.at'; Mode = 'konto' }
        $r.Address | Should -Be 'amt@x.at'; $r.Mode | Should -Be 'konto'; $script:Asked | Should -Be 0
    }
    It 'gemeinsames Postfach (im Auftrag) braucht kein Konto im Profil' {
        $r = Resolve-HUMailSender -ModuleId 'Buchung' -Sender @{ Address = 'gemeinsam@x.at'; Mode = 'imAuftrag' }
        $r.Mode | Should -Be 'imAuftrag'; $script:Asked | Should -Be 0
    }
    It 'eingestelltes Konto fehlt auf dem PC: fragen, nie still ein anderes nehmen' {
        $script:Answer = [pscustomobject]@{ Address = 'anna@x.at'; Mode = 'konto'; Remember = $false }
        $r = Resolve-HUMailSender -ModuleId 'Buchung' -Sender @{ Address = 'fehlt@x.at'; Mode = 'konto' }
        $script:Asked | Should -Be 1; $r.Address | Should -Be 'anna@x.at'
    }
    It 'Abbruch der Frage: $null (keine Mail)' {
        $script:Answer = $null
        Resolve-HUMailSender -ModuleId 'Buchung' -Sender @{ Address = 'fehlt@x.at'; Mode = 'konto' } | Should -BeNullOrEmpty
    }
    It 'nichts eingestellt, ein Konto: ohne Frage das Standardkonto' {
        $script:Accs = @($script:Accs[0])
        $r = Resolve-HUMailSender -ModuleId 'Buchung'
        $r.Address | Should -Be ''; $script:Asked | Should -Be 0
    }
    It 'nichts eingestellt, mehrere Konten: fragen und auf Wunsch je Benutzer merken' {
        $script:Answer = [pscustomobject]@{ Address = 'amt@x.at'; Mode = 'konto'; Remember = $true }
        (Resolve-HUMailSender -ModuleId 'Buchung').Address | Should -Be 'amt@x.at'
        (Resolve-HUMailSender -ModuleId 'Buchung').Address | Should -Be 'amt@x.at'
        $script:Asked | Should -Be 1
        $script:State.ContainsKey('MailSender.Buchung') | Should -BeTrue
    }
    It 'gemerktes Konto nicht mehr vorhanden: wieder fragen' {
        $script:State['MailSender.Buchung'] = [pscustomobject]@{ Address = 'weg@x.at'; Mode = 'konto' }
        $script:Answer = [pscustomobject]@{ Address = 'anna@x.at'; Mode = 'konto'; Remember = $false }
        (Resolve-HUMailSender -ModuleId 'Buchung').Address | Should -Be 'anna@x.at'
        $script:Asked | Should -Be 1
    }
}

Describe 'Mail aus einem Modul' {
    AfterEach { Remove-Item -LiteralPath Env:\HUAMT_TESTMODE -ErrorAction SilentlyContinue }
    It 'im Testmodus wird nichts geoeffnet und nicht gefragt' {
        $env:HUAMT_TESTMODE = '1'
        Mock Show-HUMailSenderDialog { throw 'darf nicht fragen' }
        (Show-HUModuleMail -ModuleId 'Buchung' -To 'a@b.at' -Subject 'x').Way | Should -Be 'testmodus'
    }
    It 'Abbruch der Absender-Frage: keine Mail' {
        Mock Test-HUOutlookClassic { $true }
        Mock Get-HUOutlookAccounts { @([pscustomobject]@{ Address = 'a@x.at'; Name = 'A' }, [pscustomobject]@{ Address = 'b@x.at'; Name = 'B' }) }
        Mock Show-HUMailSenderDialog { $null }
        Mock Show-HUMail { throw 'darf nicht oeffnen' }
        $script:State = @{}
        (Show-HUModuleMail -ModuleId 'Buchung' -To 'x@y.at').Way | Should -Be 'abgebrochen'
    }
}

Describe 'Outlook-Dialog nach vorne' {
    It 'die Fensterhilfen lassen sich laden und finden ohne Outlook kein Fenster' {
        InModuleScope HU.Office { Initialize-HUOfficeWin32; [HUOffice.Front]::FindDialog([uint32]0) } | Should -Be ([IntPtr]::Zero)
    }
    It 'ohne Outlook-Prozess startet kein Waechter' {
        Mock -ModuleName HU.Office Get-Process { @() } -ParameterFilter { $Name -eq 'OUTLOOK' }
        InModuleScope HU.Office { Start-HUOutlookDialogFront -Session $null } | Should -BeNullOrEmpty
    }
}
