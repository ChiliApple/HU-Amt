#Requires -Version 5.1
<#
.SYNOPSIS
    Fenster "Neu in dieser Version" (ab 1.0.4): zeigt die CHANGELOG-Abschnitte von HU-Amt und den Modulen,
    gezeichnet wie auf GitHub (Ueberschriften, fette Titel, Punkte mit Unterpunkten, Code), ohne Zusatzprogramm.
.DESCRIPTION
    - Beim ersten Start nach einem Update, je Benutzer (ui-state.json, Schluessel seenVersions): alle Versionen seit der
      zuletzt gesehenen, vom Programm und von jedem aktualisierten Modul (Invoke-HUChangelogOnStart).
    - Neuer Benutzer: nichts zeigen, nur die Versionen merken. Benutzer von vor 1.0.4 (noch ohne seenVersions): einmal die
      aktuelle Version zeigen.
    - Nicht im Testmodus, nicht in der Demo, nicht im Starttest.
    - Jederzeit: Rechtsklick auf "?" > "Neu in dieser Version" (Show-HUChangelog).
#>

# Programm und geladene Module mit Version und Pfad zum CHANGELOG
function Get-HUChangelogCurrent {
    $out = @([pscustomobject]@{ Key = 'HU-Amt'; Title = "$($script:AppTitle)"; Version = "$($script:Version)"; Path = (Join-Path $script:AppRoot 'CHANGELOG.md') })
    foreach ($a in @($script:LoadedAddons)) {
        if (-not $a -or -not "$($a.Id)") { continue }
        $t = $(if ("$($a.Header)") { "$($a.Header)" } elseif ("$($a.Title)") { "$($a.Title)" } else { "$($a.Id)" })
        $out += [pscustomobject]@{ Key = "addon:$($a.Id)"; Title = $t; Version = "$($a.Version)"; Path = (Join-Path "$($a.Dir)" 'CHANGELOG.md') }
    }
    return $out
}

# Aktueller Abschnitt je Eintrag (fuer den Aufruf von Hand und fuer Benutzer von vor 1.0.4)
function Get-HUChangelogLatest([object[]]$Current) {
    return @(foreach ($c in @($Current)) {
            if (-not "$($c.Path)" -or -not (Test-Path -LiteralPath "$($c.Path)")) { continue }
            $txt = ''; try { $txt = [IO.File]::ReadAllText("$($c.Path)", [Text.Encoding]::UTF8) } catch { continue }
            $sec = @(Select-HUChangelogSections (Get-HUChangelogSections $txt) $null "$($c.Version)")
            if ($sec.Count) { [pscustomobject]@{ Title = "$($c.Title)"; Key = "$($c.Key)"; Version = "$($c.Version)"; Sections = $sec } }
        })
}

function Invoke-HUChangelogOnStart {
    if ((Get-Command Test-HUDemoMode -ErrorAction SilentlyContinue) -and (Test-HUDemoMode)) { return }
    $cur = @(Get-HUChangelogCurrent)
    $raw = Get-HUStateValue 'seenVersions' $null
    $seen = @{}
    if ($raw) { foreach ($p in $raw.PSObject.Properties) { $seen[$p.Name] = "$($p.Value)" } }
    $news = @()
    if ($raw) { $news = @(Get-HUChangelogNews -Current $cur -Seen $seen) }
    elseif (-not $script:FirstRun) { $news = @(Get-HUChangelogLatest $cur) }   # Benutzer von vor 1.0.4: einmal die aktuelle Version
    # merken: aktuelle Versionen; Eintraege von Modulen, die gerade fehlen, bleiben stehen
    $o = [ordered]@{}
    foreach ($k in @($seen.Keys | Sort-Object)) { $o[$k] = $seen[$k] }
    foreach ($c in $cur) { $o["$($c.Key)"] = "$($c.Version)" }
    try { Set-HUStateValue 'seenVersions' ([pscustomobject]$o); Save-HUUIState } catch { Write-HULogWarn "Neuigkeiten: Stand nicht gespeichert: $($_.Exception.Message)" }
    if ($news.Count) {
        Write-HULogInfo "Neu in dieser Version: $(@($news | ForEach-Object { "$($_.Title) $($_.Version)" }) -join ', ')"
        Show-HUChangelogWindow -News $news -Heading 'Neu seit deinem letzten Start'
    }
}

function Show-HUChangelog {
    $news = @(Get-HUChangelogLatest @(Get-HUChangelogCurrent))
    if (-not $news.Count) { Show-HUMessage 'Zu dieser Version gibt es keine Beschreibung.' -Icon Info; return }
    Show-HUChangelogWindow -News $news -Heading 'Neu in dieser Version'
}

function Show-HUChangelogWindow([object[]]$News, [string]$Heading = 'Neu in dieser Version') {
    $x = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Neu in dieser Version" Width="860" Height="660" MinWidth="520" MinHeight="360" WindowStartupLocation="CenterOwner"
        ResizeMode="CanResizeWithGrip" Background="#0D1117" ShowInTaskbar="False">
  <Window.Resources>
    <!--HU:THEME-->
  </Window.Resources>
  <DockPanel>
    <Border DockPanel.Dock="Bottom" BorderBrush="#30363D" BorderThickness="0,1,0,0" Padding="16,10,16,14">
      <DockPanel>
        <Button x:Name="btnClose" DockPanel.Dock="Right" Content="Schließen" Width="110" Style="{StaticResource DarkButton}" IsCancel="True" IsDefault="True"
                ToolTip="Fenster schließen. Wieder öffnen: Rechtsklick auf ? und Neu in dieser Version"/>
        <TextBlock x:Name="hint" Foreground="#8B949E" VerticalAlignment="Center" TextWrapping="Wrap" Margin="0,0,12,0"
                   Text="Wieder öffnen: Rechtsklick auf ? und &quot;Neu in dieser Version&quot;."/>
      </DockPanel>
    </Border>
    <FlowDocumentScrollViewer x:Name="doc" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" IsToolBarVisible="False"
                              Background="#0D1117" ToolTip="Änderungen von HU-Amt und den Modulen"/>
  </DockPanel>
</Window>
'@
    $d = New-HUWindow -XamlText (Get-HUThemedXaml $x)
    $w = $d.Window; $c = $d.C
    $light = [bool](Test-HULightTheme)
    # Farben wie GitHub (dunkel bzw. hell)
    $pal = $(if ($light) { @{ Bg = '#FFFFFF'; Fg = '#1F2328'; Muted = '#59636E'; Line = '#D1D9E0'; CodeBg = '#EFF1F3'; Accent = '#0969DA' } }
        else { @{ Bg = '#0D1117'; Fg = '#E6EDF3'; Muted = '#8B949E'; Line = '#30363D'; CodeBg = '#262C36'; Accent = '#4493F8' } })
    $bc = New-Object System.Windows.Media.BrushConverter
    $br = { param($h) $bc.ConvertFromString($h) }
    $fd = New-Object System.Windows.Documents.FlowDocument
    $fd.FontFamily = New-Object System.Windows.Media.FontFamily('Segoe UI, Segoe UI Emoji, Segoe UI Symbol')
    $fd.FontSize = 14; $fd.Foreground = & $br $pal.Fg; $fd.Background = & $br $pal.Bg
    $fd.PagePadding = New-Object System.Windows.Thickness(28, 18, 28, 18); $fd.TextAlignment = 'Left'
    $inl = {
        param($Target, [object[]]$Parts)
        foreach ($p in @($Parts)) {
            $r = New-Object System.Windows.Documents.Run("$($p.Text)")
            if ($p.Kind -eq 'Bold') { $r.FontWeight = 'SemiBold' }
            if ($p.Kind -eq 'Code') {
                $r.FontFamily = New-Object System.Windows.Media.FontFamily('Consolas'); $r.FontSize = 13; $r.Background = & $br $pal.CodeBg
            }
            [void]$Target.Inlines.Add($r)
        }
    }
    $sep = {
        param([double]$Size, [string]$Text, [string]$Color, [double]$Top)
        $p = New-Object System.Windows.Documents.Paragraph
        $p.FontSize = $Size; $p.FontWeight = 'SemiBold'; $p.Margin = New-Object System.Windows.Thickness(0, $Top, 0, 10)
        $p.Padding = New-Object System.Windows.Thickness(0, 0, 0, 6); $p.BorderBrush = & $br $pal.Line; $p.BorderThickness = New-Object System.Windows.Thickness(0, 0, 0, 1)
        $p.Foreground = & $br $Color
        [void]$p.Inlines.Add((New-Object System.Windows.Documents.Run($Text)))
        return $p
    }
    $intro = New-Object System.Windows.Documents.Paragraph
    $intro.FontSize = 24; $intro.FontWeight = 'SemiBold'; $intro.Margin = New-Object System.Windows.Thickness(0, 0, 0, 4)
    [void]$intro.Inlines.Add((New-Object System.Windows.Documents.Run($Heading)))
    [void]$fd.Blocks.Add($intro)
    $first = $true
    foreach ($n in @($News)) {
        [void]$fd.Blocks.Add((& $sep 20 "$($n.Title) $($n.Version)" $pal.Accent $(if ($first) { 10 } else { 26 })))
        $first = $false
        foreach ($s in @($n.Sections)) {
            [void]$fd.Blocks.Add((& $sep 17 "$($s.Heading)" $pal.Fg 8))
            $stack = New-Object System.Collections.Generic.List[object]   # offene Listen je Ebene
            foreach ($b in @(ConvertTo-HUChangelogBlocks @($s.Lines))) {
                if ($b.Kind -ne 'Bullet') { $stack.Clear() }
                switch ($b.Kind) {
                    'Heading' {
                        $p = New-Object System.Windows.Documents.Paragraph
                        $p.FontSize = 15.5; $p.FontWeight = 'SemiBold'; $p.Margin = New-Object System.Windows.Thickness(0, 14, 0, 6)
                        & $inl $p $b.Parts; [void]$fd.Blocks.Add($p)
                    }
                    'Para' {
                        $p = New-Object System.Windows.Documents.Paragraph
                        $p.Margin = New-Object System.Windows.Thickness(0, 4, 0, 8); $p.Foreground = & $br $pal.Muted
                        & $inl $p $b.Parts; [void]$fd.Blocks.Add($p)
                    }
                    'Bullet' {
                        $lvl = [Math]::Min([int]$b.Level, $stack.Count)
                        while ($stack.Count -gt $lvl + 1) { $stack.RemoveAt($stack.Count - 1) }
                        if ($stack.Count -le $lvl) {
                            $l = New-Object System.Windows.Documents.List
                            $l.MarkerStyle = $(if ($lvl -eq 0) { 'Disc' } elseif ($lvl -eq 1) { 'Circle' } else { 'Square' })
                            $l.Margin = New-Object System.Windows.Thickness(0, $(if ($lvl -eq 0) { 2 } else { 0 }), 0, $(if ($lvl -eq 0) { 6 } else { 0 }))
                            $l.Padding = New-Object System.Windows.Thickness(22, 0, 0, 0)
                            if ($lvl -eq 0) { [void]$fd.Blocks.Add($l) }
                            else { $parentItem = $stack[$lvl - 1].ListItems.LastListItem; if ($parentItem) { [void]$parentItem.Blocks.Add($l) } else { [void]$fd.Blocks.Add($l) } }
                            $stack.Add($l)
                        }
                        $p = New-Object System.Windows.Documents.Paragraph
                        $p.Margin = New-Object System.Windows.Thickness(0, 2, 0, 2)
                        & $inl $p $b.Parts
                        $li = New-Object System.Windows.Documents.ListItem($p)
                        [void]$stack[$lvl].ListItems.Add($li)
                    }
                }
            }
        }
    }
    $c.doc.Document = $fd
    try { $c.doc.Background = & $br $pal.Bg; $w.Background = & $br $pal.Bg } catch { }
    $c.btnClose.Add_Click({ $w.Close() })
    [void]$w.ShowDialog()
}
