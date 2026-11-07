#Requires -Version 5.1
# Pester 5: verschiebbare Bereiche (GridSplitter) merken, Functions\UI-State.ps1. WPF-Objekte ohne Fenster, nur im STA-Thread.

BeforeAll {
    Set-StrictMode -Version 1.0   # wie Main.ps1: nicht gesetzte Variablen sind Fehler
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    $script:Sta = ([Threading.Thread]::CurrentThread.GetApartmentState() -eq 'STA')
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    $script:UserRoot = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtSp-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    [void](New-Item -ItemType Directory -Path $script:UserRoot)
    function Read-HUJsonFile([string]$Path) { if (Test-Path -LiteralPath $Path) { return ([IO.File]::ReadAllText($Path) | ConvertFrom-Json) } return $null }
    function Write-HUJsonFile([string]$Path, $Object, [int]$Depth = 5) { [IO.File]::WriteAllText($Path, ($Object | ConvertTo-Json -Depth $Depth)) }
    . (Join-Path $script:Root 'Functions\UI-State.ps1')

    function New-SplitPanel {
        $p = New-Object System.Windows.Controls.DockPanel
        $g = New-Object System.Windows.Controls.Grid; $g.Name = 'grdHaupt'
        foreach ($w in @(([System.Windows.GridLength]::new(2, 'Star')), ([System.Windows.GridLength]::new(6)), ([System.Windows.GridLength]::new(1, 'Star')))) { $cd = New-Object System.Windows.Controls.ColumnDefinition; $cd.Width = $w; $g.ColumnDefinitions.Add($cd) }
        $sp = New-Object System.Windows.Controls.GridSplitter; [System.Windows.Controls.Grid]::SetColumn($sp, 1); [void]$g.Children.Add($sp)
        [void]$p.Children.Add($g)
        return $p
    }
}
AfterAll { Remove-Item -LiteralPath $script:UserRoot -Recurse -Force -ErrorAction SilentlyContinue }

Describe 'Verschiebbare Bereiche' {
    BeforeEach { $script:UIState = $null; $script:HUSplitRoots = @{}; Remove-Item -LiteralPath $script:StateFilePath -ErrorAction SilentlyContinue }

    It 'findet das Grid mit Splitter und merkt Spaltenbreiten' {
        if (-not $script:Sta) { Set-ItResult -Skipped -Because 'WPF braucht STA'; return }
        $p = New-SplitPanel
        @(Get-HUSplitterGrids $p).Count | Should -Be 1
        $snap = @(Get-HUSplitterSnapshot $p)
        $snap[0].G | Should -Be 'grdHaupt'
        $snap[0].C | Should -Be @('Star:2', 'Pixel:6', 'Star:1')
    }
    It 'Speichern und beim naechsten Mal wiederherstellen (wie Schliessen und Neustart)' {
        if (-not $script:Sta) { Set-ItResult -Skipped -Because 'WPF braucht STA'; return }
        $p = New-SplitPanel
        Register-HUSplitterState $p 'Reiter.Buchung'
        $g = @(Get-HUSplitterGrids $p)[0]
        $g.ColumnDefinitions[0].Width = [System.Windows.GridLength]::new(5, 'Star')   # so als haette man den Splitter gezogen
        Save-HUSplitterState 'Bereiche.Reiter.Buchung'
        Test-Path $script:StateFilePath | Should -BeTrue
        # neuer Start: frischer Zustand aus der Datei, neue Oberflaeche
        $script:UIState = $null; $script:HUSplitRoots = @{}
        $p2 = New-SplitPanel
        Register-HUSplitterState $p2 'Reiter.Buchung'
        @(Get-HUSplitterGrids $p2)[0].ColumnDefinitions[0].Width.Value | Should -Be 5
        @(Get-HUSplitterGrids $p2)[0].ColumnDefinitions[0].Width.IsStar | Should -BeTrue
    }
    It 'geaenderte Oberflaeche (andere Spaltenzahl): nichts wiederherstellen' {
        if (-not $script:Sta) { Set-ItResult -Skipped -Because 'WPF braucht STA'; return }
        Set-HUStateValue 'Bereiche.Reiter.X' @([pscustomobject]@{ G = 'grdHaupt'; R = @(); C = @('Star:9', 'Pixel:6') })
        $p = New-SplitPanel
        Register-HUSplitterState $p 'Reiter.X'
        @(Get-HUSplitterGrids $p)[0].ColumnDefinitions[0].Width.Value | Should -Be 2
    }
    It 'Schliessen des Hauptfensters speichert alle Reiter' {
        if (-not $script:Sta) { Set-ItResult -Skipped -Because 'WPF braucht STA'; return }
        $p = New-SplitPanel
        Register-HUSplitterState $p 'Reiter.Y'
        @(Get-HUSplitterGrids $p)[0].ColumnDefinitions[2].Width = [System.Windows.GridLength]::new(3, 'Star')
        Save-HUAllSplitterState
        @((Get-HUStateValue 'Bereiche.Reiter.Y')[0].C)[2] | Should -Be 'Star:3'
    }
    It 'Fenster mit Restore-/Save-HUDialogState bekommen das automatisch' {
        if (-not $script:Sta) { Set-ItResult -Skipped -Because 'WPF braucht STA'; return }
        function Get-HUUiScaleValue { 1.0 }
        $w = New-Object System.Windows.Window; $w.Content = New-SplitPanel
        Restore-HUDialogState $w 'TestFenster'
        @(Get-HUSplitterGrids $w)[0].ColumnDefinitions[0].Width = [System.Windows.GridLength]::new(7, 'Star')
        Save-HUDialogState $w 'TestFenster'
        @((Get-HUStateValue 'Bereiche.Fenster.TestFenster')[0].C)[0] | Should -Be 'Star:7'
    }
}
