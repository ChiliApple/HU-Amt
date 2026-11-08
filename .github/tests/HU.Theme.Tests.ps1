#Requires -Version 5.1
# Pester 5: Farbschema Dunkel/Hell (Functions\UI-Theme.ps1). Keine Fenster, XAML wird nur geparst.

BeforeAll {
    Set-StrictMode -Version 1.0   # wie Main.ps1
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    $script:AppRoot = $script:Root
    . (Join-Path $script:Root 'Functions\UI-Theme.ps1')
    Add-Type -AssemblyName PresentationFramework
    $theme = [IO.File]::ReadAllText((Join-Path $script:Root 'XAML\Theme.xaml'), [Text.Encoding]::UTF8)
    $script:ThemeInner = [regex]::Match($theme, '(?s)<ResourceDictionary[^>]*>(.*)</ResourceDictionary>').Groups[1].Value
    function Get-TestXaml([string]$Name) { [IO.File]::ReadAllText((Join-Path $script:Root "XAML\$Name.xaml"), [Text.Encoding]::UTF8).Replace('<!--HU:THEME-->', $script:ThemeInner) }
}
AfterAll { Set-HUTheme 'Dunkel' }

Describe 'Farbschema' {
    AfterEach { Set-HUTheme 'Dunkel' }

    It 'Dunkel aendert nichts (Standard)' {
        Set-HUTheme 'Dunkel'
        $x = Get-TestXaml 'MainWindow'
        ConvertTo-HUThemeXaml $x | Should -BeExactly $x
        ConvertTo-HUThemeColor '#1E1E1E' 'Bg' | Should -Be '#1E1E1E'
    }
    It 'unbekannter Wert gilt als Dunkel' {
        Set-HUTheme 'Lila'
        Get-HUTheme | Should -Be 'Dunkel'
    }
    It 'setzt Farben je Rolle um' {
        Set-HUTheme 'Hell'
        ConvertTo-HUThemeColor '#1E1E1E' 'Bg' | Should -Be '#F2EFE9'
        ConvertTo-HUThemeColor '#CCCCCC' 'Fg' | Should -Be '#2E2A24'
        ConvertTo-HUThemeColor 'White' 'Fg' | Should -Be '#1F1C18'
        ConvertTo-HUThemeColor 'White' 'BtnFg' | Should -Be 'White'
        ConvertTo-HUThemeColor '#555555' 'Btn' | Should -Be '#6B655A'
        ConvertTo-HUThemeColor '#3E3E42' 'Border' | Should -Be '#CFC8BA'
        ConvertTo-HUThemeColor '#1976D2' 'Btn' | Should -Be '#1976D2'   # farbige Knoepfe bleiben
        ConvertTo-HUThemeColor '#801E1E1E' 'Bg' | Should -Be '#80F2EFE9' # Deckkraft bleibt
        ConvertTo-HUThemeColor '{StaticResource X}' 'Bg' | Should -Be '{StaticResource X}'
    }
    It 'kein Zielwert ist selbst ein Schluessel (doppelt anwenden aendert nichts)' {
        foreach ($r in $script:HUThemeMaps.Keys) {
            $m = $script:HUThemeMaps[$r]
            foreach ($k in $m.Keys) { $m.ContainsKey($m[$k]) | Should -BeFalse -Because "$r $k -> $($m[$k])" }
        }
    }
    It 'Knopf mit weisser Schrift bekommt mittelgraue Flaeche, Text daneben wird dunkel' {
        Set-HUTheme 'Hell'
        $x = '<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"><Button Background="#555555" Foreground="White"/><TextBlock Foreground="White"/><Border Background="#3E3E42"/></Grid>'
        $h = ConvertTo-HUThemeXaml $x
        $h | Should -Match '<Button Background="#6B655A" Foreground="White"/>'
        $h | Should -Match '<TextBlock Foreground="#1F1C18"/>'
        $h | Should -Match '<Border Background="#FFFDF9"/>'
    }
    It 'Knopf mit eigener dunkler Schrift wird wie ein Feld behandelt' {
        Set-HUTheme 'Hell'
        $h = ConvertTo-HUThemeXaml '<Button Background="#3E3E42" Foreground="#CCCCCC"/>'
        $h | Should -Be '<Button Background="#FFFDF9" Foreground="#2E2A24"/>'
    }
    It 'Setter: Knopf-Style ja, Hover mit TargetName nicht' {
        Set-HUTheme 'Hell'
        $x = '<Style TargetType="Button"><Setter Property="Background" Value="#555555"/><Setter Property="Foreground" Value="White"/><Setter TargetName="tb" Property="Background" Value="#3E3E42"/></Style>'
        $h = ConvertTo-HUThemeXaml $x
        $h | Should -Match 'Property="Background" Value="#6B655A"'
        $h | Should -Match 'Property="Foreground" Value="White"'
        $h | Should -Match 'TargetName="tb" Property="Background" Value="#FFFDF9"'
    }
    It 'Kommentare bleiben unveraendert' {
        Set-HUTheme 'Hell'
        ConvertTo-HUThemeXaml '<Grid><!-- Background="#1E1E1E" --></Grid>' | Should -Be '<Grid><!-- Background="#1E1E1E" --></Grid>'
    }
    It 'alle Fenster des Kerns laden auch in Hell und aendern sich' {
        Set-HUTheme 'Hell'
        foreach ($f in @(Get-ChildItem (Join-Path $script:Root 'XAML') -Filter *.xaml | Where-Object Name -ne 'Theme.xaml')) {
            $x = Get-TestXaml $f.BaseName
            $h = ConvertTo-HUThemeXaml $x
            $h | Should -Not -BeExactly $x -Because $f.Name
            $r = [System.Xml.XmlReader]::Create([IO.StringReader]::new($h))
            try { { [void][System.Windows.Markup.XamlReader]::Load($r) } | Should -Not -Throw -Because $f.Name } finally { $r.Close() }
        }
    }
    It 'Startbildschirm: Hell mit heller Flaeche, Dunkel unveraendert, beide laden' {
        . (Join-Path $script:Root 'Functions\UI-Splash.ps1')
        $d = Get-HUSplashXaml '1.0.0' 'HU-Amt' 'Test'
        $l = Get-HUSplashXaml '1.0.0' 'HU-Amt' 'Test' -Light
        $d | Should -Match 'Background="#FF1E1E1E"'
        $l | Should -Match 'Background="#FFF2EFE9"'
        $l | Should -Not -Match '#FF1E1E1E'
        foreach ($x in $d, $l) { { [void][System.Windows.Markup.XamlReader]::Parse($x) } | Should -Not -Throw }
    }
    It 'Theme-Teil wird zwischengespeichert und gleich umgesetzt wie ohne Zwischenspeicher' {
        Set-HUTheme 'Hell'
        $x = Get-TestXaml 'SettingsWindow'
        $a = ConvertTo-HUThemeXaml $x
        $b = ConvertTo-HUThemeXamlCore $x
        $a | Should -BeExactly $b
    }
}
