#Requires -Version 5.1
# Pester 5: Datenordner und Sperren (Core\HU.Paths.psm1). Nur kuenstliche Daten in einem Testordner unter TEMP.

BeforeAll {
    Set-StrictMode -Version 1.0   # wie Main.ps1: nicht gesetzte Variablen sind Fehler
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Import-Module (Join-Path $script:Root 'Core\HU.Paths.psm1') -Force -DisableNameChecking
}

Describe 'Datenordner' {
    BeforeEach {
        $script:App = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtTest-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        [void](New-Item -ItemType Directory -Path $script:App)
    }
    AfterEach { Remove-Item -LiteralPath $script:App -Recurse -Force -ErrorAction SilentlyContinue }

    It 'Testmodus: Datenstamm Data\_Test, laufende Apps bleiben in Data\Laufend' {
        $p = Initialize-HUDataPaths -AppRoot $script:App -UserName 'anna' -PcName 'PC-1' -TestMode
        $p.DataRoot | Should -Be (Join-Path $script:App 'Data\_Test')
        $p.UserRoot | Should -Be (Join-Path $script:App 'Data\_Test\Benutzer\anna')
        $p.RunDir | Should -Be (Join-Path $script:App 'Data\Laufend')
    }
    It 'legt Data, Benutzer und Laufend an' {
        $p = Initialize-HUDataPaths -AppRoot $script:App -UserName 'Test.Person' -PcName 'PC-1'
        $p.DataRoot | Should -Be (Join-Path $script:App 'Data')
        $p.UserRoot | Should -Be (Join-Path $script:App 'Data\Benutzer\Test.Person')
        Test-Path $p.UserRoot | Should -BeTrue
        Test-Path $p.RunDir | Should -BeTrue
    }
    It 'macht aus heiklen Zeichen im Benutzernamen einen sicheren Ordnernamen' {
        $p = Initialize-HUDataPaths -AppRoot $script:App -UserName 'a\b:c' -PcName 'PC-1'
        Split-Path $p.UserRoot -Leaf | Should -Be 'a_b_c'
        Split-Path (Split-Path $p.UserRoot -Parent) -Leaf | Should -Be 'Benutzer'
    }
    It 'Modulordner: gueltige Id wird angelegt' {
        $p = Initialize-HUDataPaths -AppRoot $script:App -UserName 'u' -PcName 'pc'
        $m = Get-HUModuleDataPath -DataRoot $p.DataRoot -ModuleId 'Beispiel'
        Test-Path $m | Should -BeTrue
    }
    It 'Modulordner: reservierte und ungueltige Ids werden abgelehnt' {
        $p = Initialize-HUDataPaths -AppRoot $script:App -UserName 'u' -PcName 'pc'
        { Get-HUModuleDataPath -DataRoot $p.DataRoot -ModuleId 'Benutzer' } | Should -Throw
        { Get-HUModuleDataPath -DataRoot $p.DataRoot -ModuleId 'Postfach' } | Should -Throw
        { Get-HUModuleDataPath -DataRoot $p.DataRoot -ModuleId 'Wiederherstellung' } | Should -Throw
        { Get-HUModuleDataPath -DataRoot $p.DataRoot -ModuleId '..\x' } | Should -Throw
        { Get-HUModuleDataPath -DataRoot $p.DataRoot -ModuleId '' } | Should -Throw
    }
    It 'Schreibprobe hinterlaesst keine Datei' {
        Assert-HUWritable $script:App
        @(Get-ChildItem -LiteralPath $script:App -Force).Count | Should -Be 0
    }
}

Describe 'Regel 2: laufende App' {
    BeforeEach {
        $script:Run = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtRun-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        [void](New-Item -ItemType Directory -Path $script:Run)
    }
    AfterEach { Remove-Item -LiteralPath $script:Run -Recurse -Force -ErrorAction SilentlyContinue }

    It 'meldet sich an und haelt die Datei offen' {
        $l = Open-HURunLock -RunDir $script:Run -UserName 'anna' -PcName 'PC-1'
        try {
            Test-Path $l.Path | Should -BeTrue
            Split-Path $l.Path -Leaf | Should -Be 'anna_PC-1.lock'
            { [IO.File]::Delete($l.Path) } | Should -Throw
        } finally { Close-HURunLock $l }
        Test-Path $l.Path | Should -BeFalse
    }
    It 'zweite Anmeldung mit gleichem Namen und PC liefert nichts' {
        $l = Open-HURunLock -RunDir $script:Run -UserName 'anna' -PcName 'PC-1'
        try { Open-HURunLock -RunDir $script:Run -UserName 'anna' -PcName 'PC-1' | Should -BeNullOrEmpty }
        finally { Close-HURunLock $l }
    }
    It 'zaehlt fremde Sitzungen, die eigene nicht' {
        $me = Open-HURunLock -RunDir $script:Run -UserName 'anna' -PcName 'PC-1'
        $other = Open-HURunLock -RunDir $script:Run -UserName 'bernd' -PcName 'PC-2'
        try {
            $r = @(Get-HURunningInstances -RunDir $script:Run -Self $me)
            $r.Count | Should -Be 1
            $r[0].UserName | Should -Be 'bernd'
            $r[0].ComputerName | Should -Be 'PC-2'
            $r[0].Since | Should -Not -BeNullOrEmpty
        } finally { Close-HURunLock $other; Close-HURunLock $me }
    }
    It 'entfernt verwaiste Lock-Dateien (Absturz)' {
        Set-Content -LiteralPath (Join-Path $script:Run 'alt_PC-9.lock') -Value 'alt|PC-9|2026-01-01T08:00:00|1'
        @(Get-HURunningInstances -RunDir $script:Run).Count | Should -Be 0
        Test-Path (Join-Path $script:Run 'alt_PC-9.lock') | Should -BeFalse
    }
    It 'Anmeldung ueberschreibt eine verwaiste eigene Datei' {
        Set-Content -LiteralPath (Join-Path $script:Run 'anna_PC-1.lock') -Value 'alt'
        $l = Open-HURunLock -RunDir $script:Run -UserName 'anna' -PcName 'PC-1'
        try { $l | Should -Not -BeNullOrEmpty } finally { Close-HURunLock $l }
    }
}

Describe 'Regel 3: gemeinsame Datei bearbeiten' {
    BeforeEach {
        $script:Dir = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtLock-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        [void](New-Item -ItemType Directory -Path $script:Dir)
        $script:File = Join-Path $script:Dir 'Liste.xlsx'
        Set-Content -LiteralPath $script:File -Value 'x'
    }
    AfterEach { Remove-Item -LiteralPath $script:Dir -Recurse -Force -ErrorAction SilentlyContinue }

    It 'der Erste bekommt die Sperre, der Zweite sieht, wer bearbeitet' {
        $a = Lock-HUDataFile -Path $script:File -UserName 'anna' -PcName 'PC-1'
        try {
            $a.Ok | Should -BeTrue
            $b = Lock-HUDataFile -Path $script:File -UserName 'bernd' -PcName 'PC-2'
            $b.Ok | Should -BeFalse
            $b.UserName | Should -Be 'anna'
            $b.ComputerName | Should -Be 'PC-1'
            Get-HULockInfoText $b | Should -Match 'anna \(PC-1\)'
            Get-HULockInfoText $b | Should -Match 'nur ansehen'
        } finally { Unlock-HUDataFile $a }
        $c = Lock-HUDataFile -Path $script:File -UserName 'bernd' -PcName 'PC-2'
        try { $c.Ok | Should -BeTrue } finally { Unlock-HUDataFile $c }
        Test-Path "$($script:File).lock" | Should -BeFalse
    }
    It 'eine verwaiste Sperre haelt niemanden auf' {
        Set-Content -LiteralPath "$($script:File).lock" -Value 'alt|PC-9|2026-01-01T08:00:00|1'
        $l = Lock-HUDataFile -Path $script:File -UserName 'anna' -PcName 'PC-1'
        try { $l.Ok | Should -BeTrue } finally { Unlock-HUDataFile $l }
    }
    It 'Unlock einer fremden Sperre tut nichts' {
        $a = Lock-HUDataFile -Path $script:File -UserName 'anna' -PcName 'PC-1'
        try {
            $b = Lock-HUDataFile -Path $script:File -UserName 'bernd' -PcName 'PC-2'
            Unlock-HUDataFile $b
            Test-Path "$($script:File).lock" | Should -BeTrue
        } finally { Unlock-HUDataFile $a }
    }
    It 'Kopie mit Zeitstempel, nur die letzten behalten' {
        1..4 | ForEach-Object { [void](Backup-HUDataFile -Path $script:File -Keep 2) }
        $k = @(Get-ChildItem -LiteralPath (Join-Path $script:Dir '_Kopien') -File)
        $k.Count | Should -Be 2
        $k[0].Name | Should -Match '^Liste_\d{8}-\d{6}(_\d+)?\.xlsx$'
    }
    It 'Aufraeumen der Kopien laesst fremde Dateien mit gleichem Anfang stehen' {
        $kd = Join-Path $script:Dir '_Kopien'
        [void](New-Item -ItemType Directory -Path $kd -Force)
        foreach ($n in 'Liste_alt.xlsx', 'Liste_Notizen.xlsx', 'Liste_20200101-000000.xlsx') { Set-Content -LiteralPath (Join-Path $kd $n) -Value 'y' }
        1..3 | ForEach-Object { [void](Backup-HUDataFile -Path $script:File -Keep 1) }
        Test-Path (Join-Path $kd 'Liste_alt.xlsx') | Should -BeTrue
        Test-Path (Join-Path $kd 'Liste_Notizen.xlsx') | Should -BeTrue
        Test-Path (Join-Path $kd 'Liste_20200101-000000.xlsx') | Should -BeFalse
        @(Get-ChildItem -LiteralPath $kd -File | Where-Object { $_.Name -match '^Liste_\d{8}-\d{6}(_\d+)?\.xlsx$' }).Count | Should -Be 1
    }
    It 'Kopie einer fehlenden Datei liefert leer' {
        Backup-HUDataFile -Path (Join-Path $script:Dir 'fehlt.xlsx') | Should -Be ''
    }
}

Describe 'Rechte der Verwalter-Liste (Entscheidung Verwalter.txt)' {
    BeforeAll {
        $script:Me = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
        $script:Inh = [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit'
        $script:NoInh = [System.Security.AccessControl.InheritanceFlags]::None
        $script:NoProp = [System.Security.AccessControl.PropagationFlags]::None
        # Nur die DACL lesen und schreiben (Set-Acl schreibt auch den Besitzer und braucht dafuer ein Privileg).
        # PowerShell 7: Erweiterungsmethoden von FileSystemAclExtensions, 5.1: Methoden von FileInfo/DirectoryInfo.
        function Get-AclTestDacl([string]$Path) {
            $i = Get-Item -LiteralPath $Path -Force
            $sec = [System.Security.AccessControl.AccessControlSections]::Access
            if ($PSVersionTable.PSVersion.Major -ge 6) { return [System.IO.FileSystemAclExtensions]::GetAccessControl($i, $sec) }
            return $i.GetAccessControl($sec)
        }
        function Set-AclTestDacl([string]$Path, $Acl) {
            $i = Get-Item -LiteralPath $Path -Force
            if ($PSVersionTable.PSVersion.Major -ge 6) { [System.IO.FileSystemAclExtensions]::SetAccessControl($i, $Acl) } else { $i.SetAccessControl($Acl) }
        }
        # nur am eigenen Testordner unter TEMP: Vererbung aus TEMP abschalten, nur der aktuelle Benutzer mit Vollzugriff
        function New-AclTestDir {
            $d = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtAcl-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
            [void](New-Item -ItemType Directory -Path $d)
            $acl = Get-AclTestDacl $d
            $acl.SetAccessRuleProtection($true, $false)
            foreach ($r in @($acl.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]))) { [void]$acl.RemoveAccessRuleSpecific($r) }
            $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($script:Me, 'FullControl', $script:Inh, $script:NoProp, 'Allow')))
            Set-AclTestDacl $d $acl
            return $d
        }
        function Add-AclTestRule([string]$Path, [string]$Sid, [string]$Rights, [switch]$Folder) {
            $acl = Get-AclTestDacl $Path
            $inh = $(if ($Folder) { $script:Inh } else { $script:NoInh })
            $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule((New-Object System.Security.Principal.SecurityIdentifier $Sid), $Rights, $inh, $script:NoProp, 'Allow')))
            Set-AclTestDacl $Path $acl
        }
        function Set-AclTestFileReadOnlyForUsers([string]$File) {
            # Vererbung an der Datei unterbrechen (geerbte Eintraege als Kopie), dann breite Schreibrechte entfernen, Lesen lassen
            $acl = Get-AclTestDacl $File
            $acl.SetAccessRuleProtection($true, $true)
            Set-AclTestDacl $File $acl
            $acl = Get-AclTestDacl $File
            foreach ($r in @($acl.GetAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier]))) {
                if ($r.IdentityReference.Value -ne $script:Me.Value) { [void]$acl.RemoveAccessRuleSpecific($r) }
            }
            $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule((New-Object System.Security.Principal.SecurityIdentifier 'S-1-5-32-545'), 'ReadAndExecute', 'Allow')))
            Set-AclTestDacl $File $acl
        }
    }
    BeforeEach {
        $script:Ad = New-AclTestDir
        $script:Af = Join-Path $script:Ad 'Verwalter.txt'
        [IO.File]::WriteAllText($script:Af, "anna`r`n")
    }
    AfterEach {
        if ($script:Ad -and (Test-Path -LiteralPath $script:Ad)) { Remove-Item -LiteralPath $script:Ad -Recurse -Force -ErrorAction SilentlyContinue }
    }

    It 'nur der eigene Benutzer: Ok, und der Hinweis sagt, dass Freigaberechte nicht geprueft sind' {
        $r = Test-HUVerwalterAcl -DataRoot $script:Ad
        $r.Status | Should -Be 'Ok'
        @($r.Funde).Count | Should -Be 0
        $r.ToolTip | Should -Match 'Freigaberechte \(SMB\).*nicht geprüft'
    }
    It 'Schreibrecht <Name> an der Datei (direkt): Warnung' -TestCases @(
        @{ Name = 'Jeder'; Sid = 'S-1-1-0'; Rights = 'Write' }
        @{ Name = 'Authentifizierte Benutzer'; Sid = 'S-1-5-11'; Rights = 'Modify' }
        @{ Name = 'Benutzer'; Sid = 'S-1-5-32-545'; Rights = 'FullControl' }
        @{ Name = 'Domänen-Benutzer'; Sid = 'S-1-5-21-1111111111-2222222222-3333333333-513'; Rights = 'Write' }
    ) {
        param($Name, $Sid, $Rights)
        Add-AclTestRule $script:Af $Sid $Rights
        $r = Test-HUVerwalterAcl -DataRoot $script:Ad
        $r.Status | Should -Be 'Warnung'
        @($r.Funde | Where-Object { $_ -like "Verwalter.txt*$Name ($Sid, direkt)*" }).Count | Should -Be 1
        $r.ToolTip | Should -Match 'nicht geprüft'
    }
    It 'Aendern fuer Benutzer am Ordner wird an die Datei vererbt: Warnung (geerbt)' {
        Add-AclTestRule $script:Ad 'S-1-5-32-545' 'Modify' -Folder
        $r = Test-HUVerwalterAcl -DataRoot $script:Ad
        $r.Status | Should -Be 'Warnung'
        @($r.Funde | Where-Object { $_ -like 'Verwalter.txt*Benutzer (S-1-5-32-545, geerbt)*' }).Count | Should -Be 1
        # Aendern enthaelt kein "Unterordner und Dateien loeschen": kein Fund am Ordner
        @($r.Funde | Where-Object { $_ -like 'Ordner Data*' }).Count | Should -Be 0
    }
    It 'empfohlene Rechte: Data mit Aendern, Datei nur Lesen und Vererbung unterbrochen: Ok' {
        Add-AclTestRule $script:Ad 'S-1-5-32-545' 'Modify' -Folder
        Set-AclTestFileReadOnlyForUsers $script:Af
        $r = Test-HUVerwalterAcl -DataRoot $script:Ad
        $r.Status | Should -Be 'Ok'
        (Get-Acl -LiteralPath $script:Af).AreAccessRulesProtected | Should -BeTrue
    }
    It 'Vollzugriff am Ordner Data: Warnung, auch wenn die Datei geschuetzt ist' {
        Set-AclTestFileReadOnlyForUsers $script:Af
        Add-AclTestRule $script:Ad 'S-1-5-11' 'FullControl' -Folder
        $r = Test-HUVerwalterAcl -DataRoot $script:Ad
        $r.Status | Should -Be 'Warnung'
        @($r.Funde | Where-Object { $_ -like 'Ordner Data mit Vollzugriff*Authentifizierte Benutzer (S-1-5-11, direkt)*' }).Count | Should -Be 1
        @($r.Funde | Where-Object { $_ -like 'Verwalter.txt*' }).Count | Should -Be 0
    }
    It 'nur "Unterordner und Dateien loeschen" am Ordner reicht fuer eine Warnung' {
        Set-AclTestFileReadOnlyForUsers $script:Af
        Add-AclTestRule $script:Ad 'S-1-1-0' 'DeleteSubdirectoriesAndFiles, ReadAndExecute' -Folder
        (Test-HUVerwalterAcl -DataRoot $script:Ad).Status | Should -Be 'Warnung'
    }
    It 'Liste fehlt und Benutzer duerfen im Ordner Dateien anlegen: Warnung' {
        Remove-Item -LiteralPath $script:Af -Force
        Add-AclTestRule $script:Ad 'S-1-5-32-545' 'Modify' -Folder
        $r = Test-HUVerwalterAcl -DataRoot $script:Ad
        $r.Status | Should -Be 'Warnung'
        @($r.Funde | Where-Object { $_ -like 'Verwalter.txt fehlt*' }).Count | Should -Be 1
    }
    It 'Lese-Rechte fuer breite Gruppen sind kein Fund' {
        Add-AclTestRule $script:Ad 'S-1-1-0' 'ReadAndExecute' -Folder
        (Test-HUVerwalterAcl -DataRoot $script:Ad).Status | Should -Be 'Ok'
    }
    It 'ACL nicht lesbar: still nicht pruefbar' {
        (Test-HUVerwalterAcl -DataRoot (Join-Path $script:Ad 'gibt-es-nicht')).Status | Should -Be 'NichtPruefbar'
    }
}
