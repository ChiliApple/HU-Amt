#Requires -Version 5.1
<#
.SYNOPSIS
    Absender fuer Mails aus Modulen: einstellen (Einstellungsseite des Moduls) oder beim Versenden fragen.
.DESCRIPTION
    Absender-Angabe eines Moduls (in seinen Einstellungen gespeichert): @{ Address = 'amt@...'; Mode = 'auto' | 'konto' | 'imAuftrag' }.
      - leer: Bei einem Konto im Outlook wird es genommen, bei mehreren fragt HU-Amt einmal (Haken "merken", je Benutzer).
      - konto: ueber dieses Konto im Outlook-Profil (MailItem.SendUsingAccount). Fehlt es auf dem PC: Hinweis und Auswahl.
      - imAuftrag: gemeinsames Postfach ohne eigenes Konto (MailItem.SentOnBehalfOfName). Braucht das Recht im Exchange (unsicher, nicht in der Doku).
      - auto: Konto, falls vorhanden, sonst im Auftrag.
    Die Mail wird immer nur angezeigt; im Feld "Von" sieht man den Absender und kann ihn vor dem Senden noch aendern.
.NOTES
    Zielmaschine: der PC, auf dem die App laeuft. PowerShell 5.1 und 7. Nur im Fenster-Thread aufrufen.
#>

function ConvertTo-HUMailSender($InputObject) {
    $a = ''; $m = 'auto'
    if ($InputObject -is [string]) { $a = "$InputObject".Trim() }
    elseif ($InputObject) {
        $g = { param($n) if ($InputObject -is [hashtable]) { $InputObject[$n] } elseif ($InputObject.PSObject.Properties[$n]) { $InputObject.$n } else { $null } }
        $a = "$(& $g 'Address')".Trim(); $mm = "$(& $g 'Mode')"
        if ($mm -in 'auto', 'konto', 'imAuftrag') { $m = $mm }
    }
    return [pscustomobject]@{ Address = $a; Mode = $m }
}

function Get-HUMailSenderStateKey([string]$ModuleId) { return "MailSender.$ModuleId" }

<#
.SYNOPSIS
    Fenster "Von welchem Konto?". -ForSettings: fuer die Einstellungsseite eines Moduls (ohne "merken", mit leerer Wahl).
    Rueckgabe: @{ Address; Mode; Remember } oder $null bei Abbruch.
#>
function Show-HUMailSenderDialog {
    param([object[]]$Accounts = @(), [string]$Text = '', [switch]$ForSettings, $Current = $null, $Owner = $null)
    $x = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Absender" Width="560" SizeToContent="Height" ResizeMode="NoResize" Background="#1E1E1E" WindowStartupLocation="CenterOwner">
    <Window.Resources><ResourceDictionary><!--HU:THEME--></ResourceDictionary></Window.Resources>
    <StackPanel Margin="16">
        <TextBlock x:Name="txtInfo" Foreground="White" TextWrapping="Wrap" Margin="0,0,0,10"/>
        <TextBlock Text="Mailkonten im Outlook dieses PCs" Style="{StaticResource FieldLabel}"/>
        <ListBox x:Name="lstAcc" Style="{StaticResource DarkListBox}" Height="130" ToolTip="Über dieses Konto wird die Mail geschickt"/>
        <TextBlock Text="oder gemeinsames Postfach (im Auftrag von)" Style="{StaticResource FieldLabel}" Margin="0,10,0,2"/>
        <TextBox x:Name="txtOther" Style="{StaticResource DarkTextBox}" ToolTip="Adresse eines gemeinsamen Postfachs ohne eigenes Konto. Dafür braucht der Benutzer im Exchange das Recht „Senden im Auftrag“ oder „Senden als“."/>
        <CheckBox x:Name="chkRemember" Style="{StaticResource DarkCheckBox}" Margin="0,10,0,0" IsChecked="True" ToolTip="Nur für dich und dieses Modul, nicht mehr fragen">
            <TextBlock Text="Für dieses Modul merken" Foreground="White"/>
        </CheckBox>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
            <Button x:Name="btnNone" Content="Keine feste Wahl" Background="#555555" Style="{StaticResource DarkButton}" Margin="0,0,8,0" Visibility="Collapsed" ToolTip="Leer lassen: bei mehreren Konten wird jeder Benutzer beim Versenden gefragt"/>
            <Button x:Name="btnOk" Content="OK" Width="110" Background="#388E3C" Style="{StaticResource DarkButton}" IsDefault="True" Margin="0,0,8,0" ToolTip="Absender übernehmen"/>
            <Button x:Name="btnCancel" Content="Abbrechen" Width="100" Background="#555555" Style="{StaticResource DarkButton}" IsCancel="True" ToolTip="Abbrechen, keine Mail"/>
        </StackPanel>
    </StackPanel>
</Window>
'@
    $d = New-HUWindow -XamlText (Get-HUThemedXaml $x)
    $w = $d.Window; $c = $d.C
    if ($Owner) { try { $w.Owner = $Owner } catch { } }
    $c.txtInfo.Text = $(if ($Text) { $Text } elseif ($ForSettings) { 'Von welchem Konto sollen Mails aus diesem Modul kommen? Leer lassen, wenn jeder selbst wählen soll.' } else { 'Von welchem Konto soll die Mail kommen?' })
    $list = @($Accounts | Where-Object { "$($_.Address)" })
    foreach ($a in $list) { [void]$c.lstAcc.Items.Add("$($a.Name) <$($a.Address)>") }
    $cur = ConvertTo-HUMailSender $Current
    if ($cur.Address) {
        $i = [Array]::FindIndex([object[]]$list, [Predicate[object]] { param($o) "$($o.Address)" -ieq $cur.Address })
        if ($i -ge 0 -and $cur.Mode -ne 'imAuftrag') { $c.lstAcc.SelectedIndex = $i } else { $c.txtOther.Text = $cur.Address }
    } elseif ($list.Count) { $c.lstAcc.SelectedIndex = 0 }
    if ($ForSettings) { $c.chkRemember.Visibility = 'Collapsed'; $c.btnNone.Visibility = 'Visible' }
    $c.txtOther.Add_TextChanged({ if ("$($c.txtOther.Text)".Trim()) { $c.lstAcc.SelectedIndex = -1 } })
    $c.lstAcc.Add_SelectionChanged({ if ($c.lstAcc.SelectedIndex -ge 0) { $c.txtOther.Text = '' } })
    $st = @{ R = $null }
    $c.btnOk.Add_Click({
            $other = "$($c.txtOther.Text)".Trim()
            if ($other) {
                if ($other -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { Show-HUMessage 'Bitte eine gültige Mailadresse eingeben.' -Icon Warning -Owner $w; return }
                $st.R = [pscustomobject]@{ Address = $other; Mode = 'imAuftrag'; Remember = [bool]$c.chkRemember.IsChecked }
            } elseif ($c.lstAcc.SelectedIndex -ge 0) {
                $st.R = [pscustomobject]@{ Address = "$($list[$c.lstAcc.SelectedIndex].Address)"; Mode = 'konto'; Remember = [bool]$c.chkRemember.IsChecked }
            } else { Show-HUMessage 'Bitte ein Konto wählen oder eine Adresse eingeben.' -Icon Warning -Owner $w; return }
            $w.Close()
        })
    $c.btnNone.Add_Click({ $st.R = [pscustomobject]@{ Address = ''; Mode = 'auto'; Remember = $false }; $w.Close() })
    [void]$w.ShowDialog()
    return $st.R
}

<#
.SYNOPSIS
    Fuer die Einstellungsseite eines Moduls: Absender waehlen. Rueckgabe: @{ Address; Mode } (Address leer = keine feste Wahl) oder $null bei Abbruch.
#>
function Select-HUMailSender {
    param($Current = $null, $Owner = $null)
    if (Test-HUTestMode) { Show-HUMessage 'Im Testmodus ist Outlook gesperrt. Den Absender bitte außerhalb des Testmodus einstellen.' -Icon Info -Owner $Owner; return $null }
    $accs = @(); try { if (Test-HUOutlookClassic) { $accs = @(Get-HUOutlookAccounts) } } catch { }
    $r = Show-HUMailSenderDialog -Accounts $accs -ForSettings -Current $Current -Owner $Owner
    if (-not $r) { return $null }
    return [pscustomobject]@{ Address = $r.Address; Mode = $r.Mode }
}

<#
.SYNOPSIS
    Absender fuer eine Mail bestimmen. Reihenfolge: Einstellung des Moduls, gemerkte Wahl des Benutzers, einziges Konto, Frage.
    Rueckgabe: @{ Address; Mode } (Address leer = Standardkonto) oder $null, wenn der Benutzer abbricht.
#>
function Resolve-HUMailSender {
    param([Parameter(Mandatory)][string]$ModuleId, $Sender = $null, $Owner = $null)
    $cfg = ConvertTo-HUMailSender $Sender
    $accs = @(); try { $accs = @(Get-HUOutlookAccounts) } catch { }
    $has = { param($addr) @($accs | Where-Object { "$($_.Address)" -ieq $addr }).Count -gt 0 }
    if ($cfg.Address) {
        if ($cfg.Mode -eq 'imAuftrag' -or (& $has $cfg.Address)) { return $cfg }
        if ($cfg.Mode -eq 'auto') { return [pscustomobject]@{ Address = $cfg.Address; Mode = 'imAuftrag' } }
        # eingestelltes Konto fehlt auf diesem PC: nie still ein anderes nehmen
        $r = Show-HUMailSenderDialog -Accounts $accs -Owner $Owner -Text "Das eingestellte Mailkonto $($cfg.Address) gibt es im Outlook dieses PCs nicht. Von welchem Konto soll die Mail kommen?"
        if (-not $r) { return $null }
        return [pscustomobject]@{ Address = $r.Address; Mode = $r.Mode }
    }
    $key = Get-HUMailSenderStateKey $ModuleId
    $saved = ConvertTo-HUMailSender (Get-HUStateValue $key $null)
    if ($saved.Address -and ($saved.Mode -eq 'imAuftrag' -or (& $has $saved.Address))) { return $saved }
    if ($accs.Count -le 1) { return [pscustomobject]@{ Address = ''; Mode = 'auto' } }
    $r = Show-HUMailSenderDialog -Accounts $accs -Owner $Owner
    if (-not $r) { return $null }
    if ($r.Remember) { Set-HUStateValue $key ([pscustomobject]@{ Address = $r.Address; Mode = $r.Mode }); try { Save-HUUIState } catch { } }
    return [pscustomobject]@{ Address = $r.Address; Mode = $r.Mode }
}

<#
.SYNOPSIS
    Mail aus einem Modul anzeigen, mit Absender nach Einstellung bzw. Rueckfrage. Parameter wie Show-HUMail, dazu -ModuleId und -Sender.
    Rueckgabe wie Show-HUMail; Way = 'abgebrochen', wenn der Benutzer die Absender-Frage abbricht.
#>
function Show-HUModuleMail {
    param([Parameter(Mandatory)][string]$ModuleId, $Sender = $null, [string[]]$To = @(), [string[]]$Cc = @(), [string]$Subject = '', [string]$Body = '',
        [string[]]$Attachments = @(), $Owner = $null)
    if (Test-HUTestMode) { return [pscustomobject]@{ Way = 'testmodus'; Message = 'Testmodus: Es wurde keine Mail geöffnet.'; From = '' } }
    if (-not (Test-HUOutlookClassic)) { return (Show-HUMail -To $To -Cc $Cc -Subject $Subject -Body $Body -Attachments $Attachments) }
    $snd = Resolve-HUMailSender -ModuleId $ModuleId -Sender $Sender -Owner $Owner
    if (-not $snd) { return [pscustomobject]@{ Way = 'abgebrochen'; Message = 'Keine Mail: Absender nicht gewählt.'; From = '' } }
    return (Show-HUMail -To $To -Cc $Cc -Subject $Subject -Body $Body -Attachments $Attachments -From $snd.Address -FromMode $(if ($snd.Address) { $snd.Mode } else { 'auto' }))
}
