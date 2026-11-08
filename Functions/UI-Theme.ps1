#Requires -Version 5.1
<#
.SYNOPSIS
    Farbschema Dunkel (Standard) oder Hell (warm, nicht reinweiss) fuer Kern und Module.
.DESCRIPTION
    Alle Oberflaechen sind dunkel gestaltet. Fuer "Hell" werden die dunklen Farben beim Laden umgesetzt, je nach Rolle:
      Hintergrund (Background, Zeilen), Text (Foreground, Cursor), Rahmen (BorderBrush, Linien), Knopf (graue Knopfflaeche).
    Wirkt an allen zentralen Ladestellen: New-HUWindow (alle Fenster, auch der Module), Reiter der Module, Toast, Get-HUBrush.
    Farbige Knoepfe (blau, gruen, rot, orange) bleiben, ihre weisse Schrift auch.
    Persoenliche Einstellung ui.theme ('Dunkel' | 'Hell'), wirkt nach dem Neustart.
.NOTES
    Dot-Source aus Main.ps1 vor UI-Common. Zielmaschine: der PC, auf dem die App laeuft.
    Regel fuer die Tabellen: kein Zielwert darf selbst ein Schluessel sein (doppelt anwenden aendert nichts).
#>

$script:HUTheme = 'Dunkel'
$script:HUThemeCache = @{}
$script:HUThemeBlock = $null

$script:HUThemeMaps = @{
    # Flaechen: Fenster, Panels, Eingabefelder, Hinweisflaechen, Auswahl
    Bg     = @{
        '1E1E1E' = 'F2EFE9'; '1A1A1A' = 'F2EFE9'; '202020' = 'F0EDE6'; '232323' = 'EDE9E2'; '252526' = 'E9E5DD'; '262626' = 'E9E5DD'
        '2A2D2E' = 'E6E1D8'; '2B2B2B' = 'E6E1D8'; '2D2D30' = 'E3DED5'; '2D2D2D' = 'E3DED5'; '303030' = 'E1DCD2'; '333333' = 'DDD7CC'
        '333337' = 'DDD7CC'; '383838' = 'DAD4C8'; '3C3C3C' = 'FFFDF9'; '3E3E42' = 'FFFDF9'; '3F3F46' = 'FFFDF9'; '404040' = 'FFFDF9'
        '505058' = 'E8E2D6'; '555555' = 'E3DED5'; '5A5A5A' = 'BFB8AA'; '787878' = 'A39C8F'; '9A9A9A' = '8A8377'
        '094771' = '3D6FB0'; '264F78' = 'A9C4E8'
        '3A2E12' = 'FFF3D6'; '3A321F' = 'FFF3D6'; '1B3A1E' = 'E3F2E1'; '1B2A3A' = 'E1ECF7'; '3A1B1B' = 'F9E1E1'; '2E2A1A' = 'FFF3D6'; '1E2A1E' = 'E3F2E1'
    }
    # Schrift
    Fg     = @{
        'CCCCCC' = '2E2A24'; 'D4D4D4' = '2E2A24'; 'C8C8C8' = '2E2A24'; 'E0E0E0' = '24211C'; 'DDDDDD' = '24211C'; 'EEEEEE' = '24211C'
        'F0F0F0' = '24211C'; 'FFFFFF' = '1F1C18'
        'AAAAAA' = '5E584E'; '9E9E9E' = '5E584E'; 'BDBDBD' = '5E584E'; 'B0B0B0' = '5E584E'; 'BBBBBB' = '5E584E'; 'A0A0A0' = '5E584E'
        '858585' = '6E675C'; '6A6A6A' = '7A7366'; '707070' = '6E675C'; '808080' = '6E675C'; '888888' = '6E675C'; '757575' = '6E675C'
        '999999' = '6E675C'
        'FFB74D' = 'B45309'; 'FFA726' = 'B45309'; 'FFCC80' = 'B45309'; 'FFE082' = '8D6E00'; 'FFD54F' = '8D6E00'; 'F9A825' = '8D6E00'
        'FFEB3B' = '8D6E00'; 'FF9800' = 'C25E00'
        '4CAF50' = '2E7D32'; '81C784' = '2E7D32'; '66BB6A' = '2E7D32'; 'A5D6A7' = '2E7D32'; '8BC34A' = '2E7D32'
        '9CDCFE' = '1F5FA8'; '4FC3F7' = '0277BD'; '64B5F6' = '1F5FA8'; '90CAF9' = '1F5FA8'; '42A5F5' = '1F5FA8'; '29B6F6' = '0277BD'
        'EF5350' = 'C62828'; 'FF5252' = 'C62828'; 'E57373' = 'C62828'; 'EF9A9A' = 'C62828'; 'F44336' = 'C62828'; 'FF8A80' = 'C62828'
        'CE93D8' = '7B1FA2'; 'BA68C8' = '7B1FA2'; '87CEFA' = '0277BD'
    }
    # Rahmen und Linien
    Border = @{
        '3E3E42' = 'CFC8BA'; '555555' = 'BFB8AA'; '2D2D30' = 'DDD6C9'; '2D2D2D' = 'DDD6C9'; '333337' = 'D6CFC2'; '3F3F46' = 'CFC8BA'
        '505058' = 'C7C0B2'; '444444' = 'CFC8BA'; '666666' = 'B5AE9F'; '1E1E1E' = 'DDD6C9'; '252526' = 'DDD6C9'
    }
    # graue Knopfflaeche mit weisser Schrift: mittleres warmes Grau, die weisse Schrift bleibt lesbar
    Btn    = @{ '3E3E42' = '6B655A'; '555555' = '6B655A'; '505058' = '6B655A'; '2D2D30' = '6B655A'; '444444' = '6B655A'; '616161' = '6B655A'; '666666' = '6B655A' }
}
# Schrift auf Knoepfen: Weiss bleibt (Knopf bleibt farbig bzw. mittelgrau), alles andere wie Text
$script:HUThemeMaps.BtnFg = @{}
foreach ($k in $script:HUThemeMaps.Fg.Keys) { if ($k -ne 'FFFFFF') { $script:HUThemeMaps.BtnFg[$k] = $script:HUThemeMaps.Fg[$k] } }
# Ohne bekannte Rolle (Farben im Code, Color=, Fill=): dunkle Grautoene wie Flaechen, helle wie Schrift, kraeftige Farben bleiben
$script:HUThemeMaps.Auto = @{}
foreach ($k in $script:HUThemeMaps.Bg.Keys) { $script:HUThemeMaps.Auto[$k] = $script:HUThemeMaps.Bg[$k] }
foreach ($k in $script:HUThemeMaps.Fg.Keys) {
    if ($k -in 'FFFFFF', '4CAF50', 'FF9800', 'EF5350', 'F44336', '66BB6A', '8BC34A', '42A5F5', '29B6F6') { continue }
    $script:HUThemeMaps.Auto[$k] = $script:HUThemeMaps.Fg[$k]
}
$script:HUThemeMaps.Auto['555555'] = '857D70'   # oft Knopfflaeche oder Rahmen: dunkel genug fuer weisse Schrift

# Bild eines Fensters als PNG (Sichtpruefung im Starttest). Wirkt auch ausserhalb des Bildschirms. Fehler werden nur protokolliert.
function Save-HUWindowShot($Window, [string]$Path) {
    try {
        $Window.UpdateLayout()
        $Window.Dispatcher.Invoke([Action]{ }, [System.Windows.Threading.DispatcherPriority]::Render)
        $el = $Window.Content
        if (-not $el -or $el.ActualWidth -lt 1) { return }
        $name = ([IO.Path]::GetFileName($Path)) -replace '[\\/:*?"<>|]', '_'
        $dir = Split-Path -Path $Path -Parent
        if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
        $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap ([int][Math]::Ceiling($el.ActualWidth)), ([int][Math]::Ceiling($el.ActualHeight)), 96, 96, ([System.Windows.Media.PixelFormats]::Pbgra32)
        $dv = New-Object System.Windows.Media.DrawingVisual
        $dc = $dv.RenderOpen()
        $dc.DrawRectangle($Window.Background, $null, [System.Windows.Rect]::new(0, 0, $el.ActualWidth, $el.ActualHeight))
        $dc.DrawRectangle([System.Windows.Media.VisualBrush]::new($el), $null, [System.Windows.Rect]::new(0, 0, $el.ActualWidth, $el.ActualHeight))
        $dc.Close()
        $rtb.Render($dv)
        $enc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
        $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
        $fs = [IO.File]::Create((Join-Path $dir $name))
        try { $enc.Save($fs) } finally { $fs.Close() }
    } catch { try { Write-HULogWarn "Bild $Path`: $($_.Exception.Message)" } catch { } }
}

function Set-HUTheme([string]$Name) {
    $script:HUTheme = $(if ("$Name" -eq 'Hell') { 'Hell' } else { 'Dunkel' })
    $script:HUThemeCache = @{}
    $script:HUThemeBlock = $null
    $script:HUDlgTheme = $null
}
function Get-HUTheme { return $script:HUTheme }
function Test-HULightTheme { return ($script:HUTheme -eq 'Hell') }

<#
.SYNOPSIS
    Eine Farbe ins aktive Farbschema umsetzen. Im dunklen Schema unveraendert.
.PARAMETER Role
    Bg | Fg | Border | Btn | BtnFg | Auto (Standard). Farben im Code am besten mit Rolle: Text = Fg, Flaeche = Bg.
.OUTPUTS
    [string] '#RRGGBB' bzw. '#AARRGGBB' oder die Eingabe, wenn nichts zu tun ist.
#>
function ConvertTo-HUThemeColor([string]$Color, [string]$Role = 'Auto') {
    if ($script:HUTheme -ne 'Hell' -or -not $Color) { return $Color }
    $v = $Color.Trim()
    $alpha = ''
    if ($v -eq 'White') { $hex = 'FFFFFF' }
    elseif ($v -match '^#([0-9A-Fa-f]{2})?([0-9A-Fa-f]{6})$') { $alpha = "$($Matches[1])"; $hex = $Matches[2].ToUpperInvariant() }
    else { return $Color }
    $map = $script:HUThemeMaps[$Role]
    if (-not $map) { $map = $script:HUThemeMaps.Auto }
    if ($map.ContainsKey($hex)) { return "#$alpha$($map[$hex])" }
    return $Color
}

# Rolle einer Eigenschaft. $Btn: Knopf-Umfeld; $WhiteText: der Knopf hat weisse (oder keine eigene) Schrift
function Get-HUThemeRole([string]$Property, [bool]$Btn, [bool]$WhiteText) {
    $p = $Property; $i = $p.LastIndexOf('.'); if ($i -ge 0) { $p = $p.Substring($i + 1) }
    switch -regex ($p) {
        '^(Background|RowBackground|AlternatingRowBackground)$' { if ($Btn -and $WhiteText) { return 'Btn' } return 'Bg' }
        '^(Foreground|CaretBrush)$' { if ($Btn) { return 'BtnFg' } return 'Fg' }
        '^(BorderBrush|HorizontalGridLinesBrush|VerticalGridLinesBrush|Stroke)$' { return 'Border' }
        default { return 'Auto' }
    }
}

<#
.SYNOPSIS
    XAML-Text ins aktive Farbschema umsetzen (nur Farben in Attributen und Settern). Im dunklen Schema unveraendert.
.DESCRIPTION
    Geht die Tags der Reihe nach durch und merkt sich das Umfeld (Knopf, Style mit TargetType Button). Setter mit TargetName
    (Hover in Vorlagen) zaehlen nicht als Knopfflaeche. Ergebnis wird je XAML-Text zwischengespeichert.
#>
function ConvertTo-HUThemeXaml([string]$Xaml) {
    if ($script:HUTheme -ne 'Hell' -or -not $Xaml) { return $Xaml }
    if ($script:HUThemeCache.ContainsKey($Xaml)) { return $script:HUThemeCache[$Xaml] }
    # Der eingesetzte Theme-Teil (XAML\Theme.xaml) steckt in fast jedem Fenster: nur einmal umsetzen, dann einsetzen
    if (-not $script:HUThemeBlock) {
        $script:HUThemeBlock = @{ Dark = ''; Light = '' }
        try {
            $tf = Join-Path (Join-Path $script:AppRoot 'XAML') 'Theme.xaml'
            $tm = [regex]::Match([IO.File]::ReadAllText($tf, [Text.Encoding]::UTF8), '(?s)<ResourceDictionary[^>]*>(.*)</ResourceDictionary>')
            if ($tm.Success -and $tm.Groups[1].Value.Length -gt 100) {
                $script:HUThemeBlock.Dark = $tm.Groups[1].Value
                $script:HUThemeBlock.Light = ConvertTo-HUThemeXamlCore $tm.Groups[1].Value
            }
        } catch { }
    }
    $tb = $script:HUThemeBlock
    if ($tb.Dark -and $Xaml.Contains($tb.Dark)) {
        $out = (ConvertTo-HUThemeXamlCore ($Xaml.Replace($tb.Dark, '<!--HU:THEME-HELL-->'))).Replace('<!--HU:THEME-HELL-->', $tb.Light)
    } else { $out = ConvertTo-HUThemeXamlCore $Xaml }
    if ($script:HUThemeCache.Count -gt 200) { $script:HUThemeCache = @{} }
    $script:HUThemeCache[$Xaml] = $out
    return $out
}
function ConvertTo-HUThemeXamlCore([string]$Xaml) {
    $tagRx = [regex]'<!--[\s\S]*?-->|<(/?)([A-Za-z_][\w:.]*)((?:"[^"]*"|''[^'']*''|[^''">])*?)(/?)>'
    $attRx = [regex]'(\s)([\w:.]+)(\s*=\s*)"([^"]*)"'
    $sb = New-Object System.Text.StringBuilder ($Xaml.Length + 256)
    $stack = New-Object System.Collections.ArrayList
    $pos = 0
    foreach ($m in $tagRx.Matches($Xaml)) {
        [void]$sb.Append($Xaml, $pos, $m.Index - $pos); $pos = $m.Index + $m.Length
        if ($m.Value.StartsWith('<!--')) { [void]$sb.Append($m.Value); continue }
        $name = $m.Groups[2].Value
        if ($m.Groups[1].Value -eq '/') {
            for ($i = $stack.Count - 1; $i -ge 0; $i--) { if ($stack[$i].N -eq $name) { $stack.RemoveRange($i, $stack.Count - $i); break } }
            [void]$sb.Append($m.Value); continue
        }
        $attrs = $m.Groups[3].Value
        $tt = ''; $mt = [regex]::Match($attrs, 'TargetType\s*=\s*"(?:\{x:Type\s+)?([\w:]+)\}?"'); if ($mt.Success) { $tt = $mt.Groups[1].Value }
        $btnHere = ($name -eq 'Button' -or $tt -match '(^|:)Button$')
        $btnAnc = $btnHere; if (-not $btnAnc) { foreach ($s in $stack) { if ($s.B) { $btnAnc = $true; break } } }
        if ($attrs -match '#[0-9A-Fa-f]{6}|"White"') {
            $isSetter = ($name -eq 'Setter')
            $setRole = ''
            $wT = $true
            if ($isSetter) {
                $prop = ''; $mp = [regex]::Match($attrs, '\sProperty\s*=\s*"([^"]*)"'); if ($mp.Success) { $prop = $mp.Groups[1].Value }
                $hasTarget = $attrs -match '\sTargetName\s*='
                # Knopf-Style: Standard DarkButton hat weisse Schrift
                $setRole = Get-HUThemeRole $prop ($btnAnc -and -not $hasTarget) $true
            } else {
                $mf = [regex]::Match($attrs, '\sForeground\s*=\s*"([^"]*)"')
                if ($mf.Success -and $mf.Groups[1].Value -notmatch '^(White|#(FF)?FFFFFF)$') { $wT = $false }
            }
            $ab = New-Object System.Text.StringBuilder
            $ap = 0
            foreach ($a in $attRx.Matches($attrs)) {
                [void]$ab.Append($attrs, $ap, $a.Index - $ap); $ap = $a.Index + $a.Length
                $an = $a.Groups[2].Value; $av = $a.Groups[4].Value
                if ($isSetter) { if ($an -eq 'Value') { $av = ConvertTo-HUThemeColor $av $setRole } }
                else { $av = ConvertTo-HUThemeColor $av (Get-HUThemeRole $an $btnAnc $wT) }
                [void]$ab.Append($a.Groups[1].Value).Append($an).Append($a.Groups[3].Value).Append('"').Append($av).Append('"')
            }
            [void]$ab.Append($attrs, $ap, $attrs.Length - $ap)
            $attrs = $ab.ToString()
        }
        $self = ($m.Groups[4].Value -eq '/')
        [void]$sb.Append('<').Append($name).Append($attrs).Append($(if ($self) { '/>' } else { '>' }))
        if (-not $self) { [void]$stack.Add(@{ N = $name; B = $btnHere }) }
    }
    [void]$sb.Append($Xaml, $pos, $Xaml.Length - $pos)
    return $sb.ToString()
}
