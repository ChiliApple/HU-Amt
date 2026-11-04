#Requires -Version 5.1
<#
.SYNOPSIS
    Beispielmodul: Dateien hineinziehen oder auswaehlen, der Reiter listet sie auf und schreibt einen Verlauf nach Data\Beispiel.
.NOTES
    Nur fuer Tests und Entwicklung (AddonsDev\Beispiel). Zielmaschine: der PC, auf dem die App laeuft.
#>

$script:BspRows = $null
$script:BspDir = ''
$script:BspTimer = $null

function Initialize-HUAddonBeispiel {
    $c = $script:Controls
    $script:BspDir = Get-HUAddonDataDir 'Beispiel'
    $script:BspRows = New-Object System.Collections.ObjectModel.ObservableCollection[object]
    $c['bspGrid'].ItemsSource = $script:BspRows
    $c['bspTxtFolder'].Text = "Daten: $($script:BspDir)"
    $c['bspBtnMain'].Add_Click({
            $dlg = New-Object Microsoft.Win32.OpenFileDialog
            $dlg.Multiselect = $true
            $dlg.Filter = 'Unterstützte Dateien|*.pdf;*.xlsx;*.txt|Alle Dateien|*.*'
            if ($dlg.ShowDialog($script:Window)) { Invoke-HUAddonDropBeispiel -Paths @($dlg.FileNames) }
        })
    $c['bspBtnFolder'].Add_Click({ Open-HUPath $script:BspDir })
    $c['bspBtnSettings'].Add_Click({ Show-HUMessage 'Das Beispiel hat keine Einstellungen.' -Icon Info })
    $script:BspTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:BspTimer.Interval = [TimeSpan]::FromMilliseconds(1200)
    $script:BspTimer.Add_Tick({ $this.Stop(); Set-HUBusyDot $script:Controls['bspDot'] $false })
}

# Dateien uebernehmen (Hineinziehen ins Fenster oder Hauptknopf)
function Invoke-HUAddonDropBeispiel([string[]]$Paths) {
    Set-HUBusyDot $script:Controls['bspDot'] $true
    $n = 0
    foreach ($p in @($Paths)) {
        $f = Get-Item -LiteralPath $p -ErrorAction SilentlyContinue
        if (-not $f) { continue }
        $script:BspRows.Insert(0, [pscustomobject]@{
                Zeit    = (Get-Date).ToString('HH:mm:ss')
                Datei   = $f.Name
                Groesse = '{0:N0} KB' -f [Math]::Ceiling($f.Length / 1024.0)
                Ordner  = $f.DirectoryName
            })
        Add-Content -LiteralPath (Join-Path $script:BspDir 'verlauf.txt') -Value ('{0:s}|{1}|{2}' -f (Get-Date), $script:UserName, $f.Name) -Encoding UTF8
        $n++
    }
    $script:Controls['bspTxtStatus'].Text = "$n Datei(en) übernommen"
    Write-HULogOK "Beispiel: $n Datei(en) übernommen"
    $script:BspTimer.Stop(); $script:BspTimer.Start()
}

function Stop-HUAddonBeispiel {
    if ($script:BspTimer) { $script:BspTimer.Stop() }
}
