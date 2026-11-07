#Requires -Version 5.1
<#
.SYNOPSIS
    Kontakt uebernehmen (fuer alle Module, z. B. Mieter in Buchung): aus Outlook-Kontakten oder vCard-Dateien (.vcf).
.DESCRIPTION
    Select-HUContact oeffnet ein Fenster mit Suche, "In Outlook suchen", "vCard-Datei öffnen ..." und Hineinziehen von .vcf-Dateien.
    Rueckgabe: ausgewaehlte Kontakt-Objekte (Name, Firma, Mail, Telefon, Mobil, Strasse, Plz, Ort, Land, Quelle) oder leer bei Abbruch.
    Das Modul uebernimmt die Felder selbst in seine Daten (nur gewaehlte Kontakte, nichts automatisch).
    Im Testmodus ist Outlook gesperrt, vCard-Dateien gehen.
.NOTES
    Zielmaschine: der PC, auf dem die App laeuft. PowerShell 5.1 und 7. Nur im Fenster-Thread aufrufen.
#>

function Select-HUContact {
    param([switch]$Multiple, [string]$Title = 'Kontakt übernehmen', $Owner = $null)
    $x = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Kontakt übernehmen" Width="900" Height="560" MinWidth="640" MinHeight="380" Background="#1E1E1E" ResizeMode="CanResizeWithGrip"
        WindowStartupLocation="CenterOwner" AllowDrop="True">
    <Window.Resources><ResourceDictionary><!--HU:THEME--></ResourceDictionary></Window.Resources>
    <DockPanel Margin="12">
        <DockPanel DockPanel.Dock="Top" Margin="0,0,0,8">
            <Button x:Name="btnVcf" DockPanel.Dock="Right" Content="vCard-Datei öffnen ..." Style="{StaticResource ToolButton}" Margin="8,0,0,0" ToolTip="Kontakte aus einer oder mehreren .vcf-Dateien lesen (auch hineinziehen)"/>
            <Button x:Name="btnOutlook" DockPanel.Dock="Right" Content="In Outlook suchen" Background="#1565C0" Style="{StaticResource DarkButton}" Margin="8,0,0,0" ToolTip="Kontakte aus dem eigenen Outlook (Ordner Kontakte) nach dem Suchtext laden"/>
            <TextBox x:Name="txtSearch" Style="{StaticResource DarkTextBox}" ToolTip="Name, Firma, Mail oder Ort; filtert die Liste sofort"/>
        </DockPanel>
        <DockPanel DockPanel.Dock="Bottom" Margin="0,10,0,0">
            <StackPanel Orientation="Horizontal" DockPanel.Dock="Right">
                <Button x:Name="btnOk" Content="Übernehmen" Width="120" Background="#388E3C" Style="{StaticResource DarkButton}" IsDefault="True" Margin="0,0,8,0" ToolTip="Markierte Kontakte übernehmen"/>
                <Button x:Name="btnCancel" Content="Abbrechen" Width="100" Background="#555555" Style="{StaticResource DarkButton}" IsCancel="True" ToolTip="Nichts übernehmen"/>
            </StackPanel>
            <TextBlock x:Name="txtInfo" Foreground="#BDBDBD" VerticalAlignment="Center" TextWrapping="Wrap" ToolTip="Quelle und Anzahl"/>
        </DockPanel>
        <DataGrid x:Name="grd" Style="{StaticResource DarkDataGrid}" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column"
                  ToolTip="Kontakt markieren (mehrere mit Strg) und übernehmen, Doppelklick übernimmt sofort">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Name" Binding="{Binding Name}"/>
                <DataGridTextColumn Header="Firma / Verein" Binding="{Binding Firma}"/>
                <DataGridTextColumn Header="Mail" Binding="{Binding Mail}"/>
                <DataGridTextColumn Header="Telefon" Binding="{Binding Telefon}"/>
                <DataGridTextColumn Header="Mobil" Binding="{Binding Mobil}"/>
                <DataGridTextColumn Header="Straße" Binding="{Binding Strasse}"/>
                <DataGridTextColumn Header="PLZ" Binding="{Binding Plz}"/>
                <DataGridTextColumn Header="Ort" Binding="{Binding Ort}"/>
                <DataGridTextColumn Header="Quelle" Binding="{Binding Quelle}"/>
            </DataGrid.Columns>
        </DataGrid>
    </DockPanel>
</Window>
'@
    $d = New-HUWindow -XamlText (Get-HUThemedXaml $x)
    $w = $d.Window; $c = $d.C
    if ($Owner) { try { $w.Owner = $Owner } catch { } }
    $w.Title = $Title
    $c.grd.SelectionMode = $(if ($Multiple) { 'Extended' } else { 'Single' })
    $st = @{ All = @(); R = @() }
    $outlookOk = (-not (Test-HUTestMode)) -and (Test-HUOutlookClassic)
    if (-not $outlookOk) { $c.btnOutlook.IsEnabled = $false; $c.btnOutlook.ToolTip = $(if (Test-HUTestMode) { 'Im Testmodus ist Outlook gesperrt; vCard-Dateien gehen.' } else { 'Kein klassisches Outlook auf diesem PC; vCard-Dateien gehen.' }) }
    $c.txtInfo.Text = 'Suchtext eingeben und „In Outlook suchen“ oder eine vCard-Datei öffnen bzw. hineinziehen.'
    $refresh = {
        $list = @(Find-HUContact -Contacts $st.All -Search $c.txtSearch.Text)
        $c.grd.ItemsSource = $list
        $c.txtInfo.Text = "$($list.Count) von $(@($st.All).Count) Kontakten"
    }
    $addVcf = {
        param([string[]]$Paths)
        $n = 0; $err = @()
        foreach ($p in @($Paths | Where-Object { $_ -match '(?i)\.vcf$' })) {
            try { $k = @(Import-HUVCardFile -Path $p); $st.All = @($st.All) + $k; $n += $k.Count } catch { $err += "$([IO.Path]::GetFileName($p)): $($_.Exception.Message)" }
        }
        & $refresh
        if ($err.Count) { Show-HUMessage ("Nicht gelesen:`n" + ($err -join "`n")) -Icon Warning -Owner $w }
        elseif ($n -eq 0) { Show-HUMessage 'In der Datei war kein Kontakt.' -Icon Information -Owner $w }
    }
    $c.txtSearch.Add_TextChanged({ & $refresh })
    $c.btnOutlook.Add_Click({
            $w.Cursor = [System.Windows.Input.Cursors]::Wait
            try {
                $k = @(Get-HUOutlookContacts -Search $c.txtSearch.Text -Max 500)
                $st.All = @(@($st.All | Where-Object { $_.Quelle -ne 'Outlook' }) + $k)
                & $refresh
                if ($k.Count -ge 500) { $c.txtInfo.Text += ' (höchstens 500 aus Outlook, Suchtext genauer fassen)' }
            } catch { Show-HUMessage "Outlook-Kontakte nicht lesbar:`n$($_.Exception.Message)" -Icon Warning -Owner $w }
            finally { $w.Cursor = $null }
        })
    $c.btnVcf.Add_Click({
            $dlg = New-Object Microsoft.Win32.OpenFileDialog
            $dlg.Filter = 'vCard (*.vcf)|*.vcf|Alle Dateien (*.*)|*.*'; $dlg.Multiselect = $true; $dlg.Title = 'vCard-Datei öffnen'
            if ($dlg.ShowDialog($w)) { & $addVcf $dlg.FileNames }
        })
    $w.Add_DragOver({ param($s, $e) $e.Effects = $(if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) { 'Copy' } else { 'None' }); $e.Handled = $true })
    $w.Add_Drop({ param($s, $e) if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) { & $addVcf @($e.Data.GetData([System.Windows.DataFormats]::FileDrop)) } })
    $take = {
        $sel = @($c.grd.SelectedItems)
        if (-not $sel.Count) { Show-HUMessage 'Bitte einen Kontakt markieren.' -Icon Information -Owner $w; return }
        $st.R = $sel; $w.Close()
    }
    $c.btnOk.Add_Click({ & $take })
    $c.grd.Add_MouseDoubleClick({ if ($c.grd.SelectedItem) { & $take } })
    [void]$w.ShowDialog()
    return @($st.R)
}
