#Requires -Version 5.1
<#
.SYNOPSIS
    HU.Tools - externe Programme aufrufen und Dateinamen fuer Ergebnisse bilden (gemeinsam fuer alle Module).
.DESCRIPTION
    Prozesse:
      ConvertTo-HUProcessArgument maskiert ein Argument nach den Regeln der Microsoft-C-Laufzeit
      (CommandLineToArgvW): Leerzeichen/Tab/Anfuehrungszeichen -> in "..." setzen, \ vor " verdoppeln und " als \",
      \ am Ende verdoppeln. Gilt fuer Programme, die ihre Befehlszeile so zerlegen (Python, Ghostscript, qpdf, 7-Zip,
      .NET-Programme). cmd.exe zerlegt anders: fuer cmd.exe /c keine fremden Texte als Argumente uebergeben.
      Invoke-HUNativeProcess startet ein Programm ohne Fenster, liest stdout/stderr (UTF-8), schreibt wahlweise Text
      als UTF-8 in die Eingabe, setzt Umgebung und Arbeitsordner und beendet den Prozessbaum bei Zeitueberschreitung.
      -State (Hashtable) bekommt waehrend des Laufs .Pid, damit ein anderer Thread mit Stop-HUProcessTree abbrechen kann.
    Dateinamen:
      ConvertTo-HUSafeFileName macht aus einem Text einen gueltigen Windows-Dateinamen (Leerzeichen und Umlaute bleiben).
      Get-HUFreeFilePath / Get-HUFreeDirectoryPath liefern <Name><Endung>, sonst <Name>_2, _3 ... (nie ueberschreiben).
      Fuer Dateinamen mit Zeitstempel im Uebergabeordner gibt es New-HUHandoverPath in HU.Exchange.
    Laeuft ohne Fenster und ohne Functions\ (auch in einem Hintergrund-Runspace).
.NOTES
    Zielmaschine: der PC, auf dem die App laeuft. PowerShell 5.1 und 7.
#>

Set-StrictMode -Version 1.0

# ---------------------------------------------------------------- Prozesse ----------------------------------------------------------------

<#
.SYNOPSIS
    Ein Argument fuer ProcessStartInfo.Arguments maskieren (Regeln der Microsoft-C-Laufzeit).
    Leerer Text wird zu "" (sonst fiele das Argument weg).
#>
function ConvertTo-HUProcessArgument([string]$Argument) {
    if ($Argument -and $Argument -notmatch '[\s"]') { return $Argument }
    $s = "$Argument" -replace '(\\*)"', '$1$1\"'
    $s = $s -replace '(\\+)$', '$1$1'
    return '"' + $s + '"'
}

<#
.SYNOPSIS
    Mehrere Argumente maskieren und mit Leerzeichen verbinden (Text fuer ProcessStartInfo.Arguments).
#>
function Join-HUProcessArguments([string[]]$Arguments = @()) {
    return (@($Arguments | ForEach-Object { ConvertTo-HUProcessArgument "$_" }) -join ' ')
}

<#
.SYNOPSIS
    Prozessbaum beenden (taskkill /T /F). Fehler werden verschluckt (Prozess kann schon beendet sein).
#>
function Stop-HUProcessTree([int]$ProcessId) {
    if ($ProcessId -le 0) { return }
    try { & "$env:SystemRoot\System32\taskkill.exe" /PID $ProcessId /T /F 2>&1 | Out-Null } catch { }
}

<#
.SYNOPSIS
    Programm ohne Fenster starten, Ausgabe lesen, Zeitlimit.
.DESCRIPTION
    -Arguments: einzelne Argumente, werden mit ConvertTo-HUProcessArgument maskiert.
    -Env: zusaetzliche Umgebungsvariablen. Schluessel PATHVORNE: Text wird vorne an PATH gesetzt.
    -WorkDir: Arbeitsordner (leer = aktueller Prozessordner).
    -StdIn: Text, der als UTF-8 ohne BOM an die Eingabe geht (z. B. Auftraege mit Kennwort, die so nie auf der Platte
            landen). Die Eingabe wird immer umgeleitet und danach geschlossen, auch ohne Text: ein Programm, das
            unerwartet auf Eingabe wartet, bekommt Dateiende und haengt nicht bis zum Zeitlimit.
    -State: Hashtable; waehrend des Laufs steht dort .Pid (danach 0), zum Abbrechen aus einem anderen Thread.
    Rueckgabe: [pscustomobject] @{ ExitCode; Out; Err; TimedOut }. Bei Zeitueberschreitung ExitCode -1, Out leer,
    Err "Zeitüberschreitung nach <n> s", der Prozessbaum ist beendet.
    Laesst sich das Programm nicht starten (Datei fehlt), wirft die Funktion (wie Process.Start).
#>
function Invoke-HUNativeProcess {
    param(
        [Parameter(Mandatory)][string]$Exe,
        [string[]]$Arguments = @(),
        [int]$TimeoutSec = 600,
        [hashtable]$Env = @{},
        [string]$WorkDir = '',
        [string]$StdIn = '',
        $State = $null
    )
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $Exe
    $psi.Arguments = Join-HUProcessArguments $Arguments
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.RedirectStandardInput = $true
    $psi.StandardOutputEncoding = [Text.Encoding]::UTF8
    $psi.StandardErrorEncoding = [Text.Encoding]::UTF8
    if ($WorkDir) { $psi.WorkingDirectory = $WorkDir }
    foreach ($k in @($Env.Keys)) {
        if ($k -eq 'PATHVORNE') { $psi.EnvironmentVariables['PATH'] = "$($Env[$k]);$($psi.EnvironmentVariables['PATH'])" }
        else { $psi.EnvironmentVariables[$k] = "$($Env[$k])" }
    }
    $ms = [int][Math]::Min([double]$TimeoutSec * 1000, [double][int]::MaxValue)
    if ($ms -lt 0) { $ms = 0 }
    $p = [System.Diagnostics.Process]::Start($psi)
    if ($State) { $State.Pid = $p.Id }
    try {
        $o = $p.StandardOutput.ReadToEndAsync(); $e = $p.StandardError.ReadToEndAsync()
        # direkt als Bytes: StandardInputEncoding gibt es unter .NET Framework (PS 5.1) nicht
        if ($StdIn) {
            $b = (New-Object System.Text.UTF8Encoding $false).GetBytes($StdIn)
            try { $p.StandardInput.BaseStream.Write($b, 0, $b.Length); $p.StandardInput.BaseStream.Flush() } catch { }
        }
        try { $p.StandardInput.Close() } catch { }
        if (-not $p.WaitForExit($ms)) {
            Stop-HUProcessTree $p.Id
            return [pscustomobject]@{ ExitCode = -1; Out = ''; Err = "Zeitüberschreitung nach $TimeoutSec s"; TimedOut = $true }
        }
        # ohne Zeitangabe wartet WaitForExit auch auf das Ende der umgeleiteten Ausgabe
        $p.WaitForExit()
        return [pscustomobject]@{ ExitCode = $p.ExitCode; Out = $o.Result; Err = $e.Result; TimedOut = $false }
    } finally {
        if ($State) { $State.Pid = 0 }
        $p.Dispose()
    }
}

# ---------------------------------------------------------------- Dateinamen ----------------------------------------------------------------

<#
.SYNOPSIS
    Gueltiger Windows-Dateiname aus einem Text (ohne Endung uebergeben).
.DESCRIPTION
    Ersetzt \ / : * ? " < > | und Steuerzeichen durch _, fasst Leerraum und __ zusammen, kuerzt auf -MaxLength,
    schneidet Leerzeichen, Punkte, _ und - am Rand ab, setzt _ vor reservierte Geraetenamen (CON, PRN, AUX, NUL, COM1-9, LPT1-9).
    Standard: "A _ B" wird zu "A_B".
    -KeepSeparators: einfachere Regel - Leerzeichen um _ bleiben, - am Rand bleibt, nur Doppel-Leerzeichen werden eins.
    -Fallback: Name, wenn nichts uebrig bleibt (wird ebenfalls bereinigt), sonst "Dokument".
#>
function ConvertTo-HUSafeFileName([string]$Name, [string]$Fallback = 'Dokument', [int]$MaxLength = 150, [switch]$KeepSeparators) {
    $bad = '[\\/:*?"<>|\x00-\x1F]'
    $n = "$Name" -replace $bad, '_'
    if ($KeepSeparators) {
        $n = $n -replace '_{2,}', '_' -replace '\s{2,}', ' '
        $trim = [char[]]@(' ', '.', '_')
    } else {
        $n = $n -replace '\s+', ' '
        $n = $n -replace '_{2,}', '_'
        $n = $n -replace '(\s*_\s*)+', '_'
        $trim = [char[]]@(' ', '.', '_', '-')
    }
    $n = $n.Trim($trim)
    if ($MaxLength -gt 0 -and $n.Length -gt $MaxLength) { $n = $n.Substring(0, $MaxLength).Trim($trim) }
    if (-not $n) { $n = ("$Fallback" -replace $bad, '_').Trim(' ', '.') }
    if (-not $n) { $n = 'Dokument' }
    if ($n -match '^(?i)(CON|PRN|AUX|NUL|COM\d|LPT\d)$') { $n = "_$n" }
    return $n
}

<#
.SYNOPSIS
    Freier Dateipfad im Ordner: <Name><Endung>, sonst <Name>_2<Endung>, _3 ... (nie ueberschreiben).
    -Extension mit oder ohne Punkt; leer = ohne Endung. Der Pfad wird nur gesucht, nicht angelegt: wer gleichzeitig
    schreiben kann (Freigabe), legt die Datei mit FileMode.CreateNew an und sucht bei einem Fehler neu.
#>
function Get-HUFreeFilePath([Parameter(Mandatory)][string]$Dir, [Parameter(Mandatory)][string]$Name, [string]$Extension = '.pdf') {
    $ext = "$Extension"
    if ($ext -and -not $ext.StartsWith('.')) { $ext = ".$ext" }
    $p = Join-Path $Dir "$Name$ext"; $n = 1
    while (Test-Path -LiteralPath $p) { $n++; $p = Join-Path $Dir ("{0}_{1}{2}" -f $Name, $n, $ext) }
    return $p
}

<#
.SYNOPSIS
    Freier Ordnerpfad: <Name>, sonst <Name>_2, _3 ... (wird nur gesucht, nicht angelegt).
#>
function Get-HUFreeDirectoryPath([Parameter(Mandatory)][string]$Dir, [Parameter(Mandatory)][string]$Name) {
    $p = Join-Path $Dir $Name; $n = 1
    while (Test-Path -LiteralPath $p) { $n++; $p = Join-Path $Dir ("{0}_{1}" -f $Name, $n) }
    return $p
}

Export-ModuleMember -Function ConvertTo-HUProcessArgument, Join-HUProcessArguments, Stop-HUProcessTree, Invoke-HUNativeProcess,
    ConvertTo-HUSafeFileName, Get-HUFreeFilePath, Get-HUFreeDirectoryPath
