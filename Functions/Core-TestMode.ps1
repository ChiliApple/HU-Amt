#Requires -Version 5.1
<#
.SYNOPSIS
    Testmodus: nur fuer Verwalter, nur fuer die laufende Sitzung auf diesem PC.
.DESCRIPTION
    - HU-Amt startet neu mit -TestMode. Der Datenstamm ist dann Data\_Test\ statt Data\: alle Module bekommen ueber
      Get-HUAddonDataDir bzw. $script:DataRoot automatisch einen getrennten Testbestand. Echte Daten bleiben unberuehrt.
    - Outlook, Word und Mail sind gesperrt (Core\HU.Office.psm1 prueft $env:HUAMT_TESTMODE, gilt auch in Hintergrund-Auftraegen).
    - Keine automatische Sicherung, orange Leiste im Hauptfenster, Fenstertitel mit "TEST".
    - "Test zurücksetzen" loescht beim naechsten Start alles unter Data\_Test\ (nur wenn dort die Kennungsdatei liegt).
    - Ein normaler Neustart beendet den Testmodus; Neustarts aus dem Testmodus heraus (z. B. nach Modul-Installation) bleiben im Test.
.NOTES
    Zielmaschine: der PC, auf dem die App laeuft. PowerShell 5.1 und 7.
#>

$script:HUTestMarker = '.hu-testbestand'
$script:HUTestResetMarker = '.zuruecksetzen'

<#
.SYNOPSIS
    $true, wenn die App im Testmodus laeuft. Module pruefen das z. B. vor Outlook oder Word und zeigen dann eine Simulation.
#>
function Test-HUTestMode { return ($env:HUAMT_TESTMODE -eq '1') }

<#
.SYNOPSIS
    $true im Demo-Modus (fuer alle Benutzer, Beispieldaten, bei jedem Start frisch). Demo ist eine Form des Testmodus:
    Test-HUTestMode ist dann ebenfalls $true (Outlook, Word, Mail gesperrt). Pruefungen wie der gefuehrte Test gibt es nicht.
#>
function Test-HUDemoMode { return ($env:HUAMT_DEMO -eq '1') }

# Unter dem eingestellten Datenordner (Get-HUDataLocation, Standard <App>\Data); im Demo-Modus _Demo\<Benutzer>
function Get-HUTestDataRoot([string]$AppRoot = $script:AppRoot) {
    $base = (Get-HUDataLocation -AppRoot $AppRoot).Path
    if (Test-HUDemoMode) { return (Join-Path (Join-Path $base '_Demo') (Get-HUSafeName $env:USERNAME)) }
    return (Join-Path $base '_Test')
}

<#
.SYNOPSIS
    Beim Start im Demo-Modus: eigenen Demo-Bestand des Benutzers leeren (nur unter _Demo\<Benutzer> mit Kennung),
    Verwalter-Liste uebernehmen. Die Beispieldaten legt danach Initialize-HUAddonTestData je Modul hinein.
#>
function Initialize-HUDemoData([string]$AppRoot = $script:AppRoot) {
    if (-not (Test-HUDemoMode)) { throw 'Demo-Modus ist nicht aktiv.' }
    $root = Get-HUTestDataRoot $AppRoot
    if (-not (Test-Path -LiteralPath $root)) { [void](New-Item -ItemType Directory -Path $root -Force) }
    $mk = Join-Path $root $script:HUTestMarker
    if (-not (Test-Path -LiteralPath $mk)) {
        [IO.File]::WriteAllText($mk, "Demo-Bestand von HU-Amt. Wird bei jedem Demo-Start geleert. Angelegt $(Get-Date -Format 'yyyy-MM-dd HH:mm').`r`n")
    }
    foreach ($i in @(Get-ChildItem -LiteralPath $root -Force)) {
        if ($i.Name -ceq $script:HUTestMarker) { continue }
        if (-not (Test-HUTestDataPath $i.FullName $AppRoot)) { throw "Pfad liegt nicht im Demo-Bestand: $($i.FullName)" }
        Remove-Item -LiteralPath $i.FullName -Recurse -Force
    }
    $real = Join-Path (Get-HUDataLocation -AppRoot $AppRoot).Path 'Verwalter.txt'
    if (Test-Path -LiteralPath $real) { Copy-Item -LiteralPath $real -Destination (Join-Path $root 'Verwalter.txt') -Force }
}

<#
.SYNOPSIS
    Demo starten (alle Benutzer) oder beenden. Beides startet die App neu.
#>
function Switch-HUDemoMode {
    if (Test-HUDemoMode) {
        if (Confirm-HU "Demo beenden?`n`n$($script:AppTitle) startet neu und arbeitet wieder mit den echten Daten." 'Demo') { Restart-HUApp -LeaveDemo }
        return
    }
    $txt = "Demo starten?`n`n$($script:AppTitle) startet neu mit Beispieldaten in allen Modulen. Probier alles aus, es ist nichts davon echt.`n`n" +
    "Outlook, Word und Mails sind in der Demo gesperrt. Beim nächsten Demo-Start ist alles wieder wie neu. Deine echten Daten bleiben unberührt."
    if (Confirm-HU $txt 'Demo') { Restart-HUApp -EnterDemo }
}

<#
.SYNOPSIS
    Sicherheitspruefung vor jedem Loeschen von Testdaten: Pfad liegt unter Data\_Test\ und der Testbestand ist gekennzeichnet.
#>
function Test-HUTestDataPath([string]$Path, [string]$AppRoot = $script:AppRoot) {
    if (-not $Path) { return $false }
    $root = [IO.Path]::GetFullPath((Get-HUTestDataRoot $AppRoot)).TrimEnd('\')
    $full = [IO.Path]::GetFullPath($Path).TrimEnd('\')
    if (-not ($full.Equals($root, [StringComparison]::OrdinalIgnoreCase) -or $full.StartsWith($root + '\', [StringComparison]::OrdinalIgnoreCase))) { return $false }
    return (Test-Path -LiteralPath (Join-Path $root $script:HUTestMarker))
}

<#
.SYNOPSIS
    Beim Start im Testmodus, vor Protokoll und Modulen: Testbestand anlegen bzw. zuruecksetzen, Verwalter-Liste uebernehmen
    (bei jedem Start neu).
    Rueckgabe: $true, wenn zurueckgesetzt wurde.
#>
function Initialize-HUTestData([string]$AppRoot = $script:AppRoot) {
    $root = Get-HUTestDataRoot $AppRoot
    if (-not (Test-Path -LiteralPath $root)) { [void](New-Item -ItemType Directory -Path $root -Force) }
    $mk = Join-Path $root $script:HUTestMarker
    if (-not (Test-Path -LiteralPath $mk)) {
        [IO.File]::WriteAllText($mk, "Testbestand von HU-Amt. Darf jederzeit geloescht werden. Angelegt $(Get-Date -Format 'yyyy-MM-dd HH:mm').`r`n")
    }
    $reset = $false
    if (Test-Path -LiteralPath (Join-Path $root $script:HUTestResetMarker)) {
        if (-not (Test-HUTestDataPath $root $AppRoot)) { throw 'Der Testbestand ist nicht gekennzeichnet. Zurücksetzen abgebrochen.' }
        foreach ($i in @(Get-ChildItem -LiteralPath $root -Force)) {
            if ($i.Name -ceq $script:HUTestMarker) { continue }
            if (-not (Test-HUTestDataPath $i.FullName $AppRoot)) { throw "Pfad liegt nicht im Testbestand: $($i.FullName)" }
            Remove-Item -LiteralPath $i.FullName -Recurse -Force
        }
        $reset = $true
    }
    # Verwalter bleiben im Test Verwalter (sonst fehlte der Reiter Verwalter und der Weg zurueck).
    # Bei jedem Start neu kopieren (KERN-B17): Aenderungen der echten Liste gelten auch im Testmodus.
    $real = Join-Path (Get-HUDataLocation -AppRoot $AppRoot).Path 'Verwalter.txt'
    $test = Join-Path $root 'Verwalter.txt'
    if (Test-Path -LiteralPath $real) { Copy-Item -LiteralPath $real -Destination $test -Force }
    return $reset
}

<#
.SYNOPSIS
    Beispieldaten eines Moduls in den Testbestand legen: <Modul>\Testdaten\ -> Data\_Test\<Id>\.
.DESCRIPTION
    Nur im Testmodus, nur unter dem gekennzeichneten Testbestand und nur, wenn dort fuer das Modul noch nichts liegt
    (beim ersten Start im Testmodus und nach "Test zuruecksetzen"). Echte Daten werden nie beruehrt.
    Verknuepfungen werden uebersprungen. Rueckgabe: Anzahl kopierter Dateien.
#>
function Initialize-HUAddonTestData($Addon, [string]$AppRoot = $script:AppRoot) {
    if (-not (Test-HUTestMode) -or -not $Addon -or -not "$($Addon.Dir)" -or -not "$($Addon.Id)") { return 0 }
    $src = [IO.Path]::GetFullPath((Join-Path $Addon.Dir 'Testdaten')).TrimEnd('\')
    if (-not (Test-Path -LiteralPath $src -PathType Container)) { return 0 }
    $dst = Join-Path (Get-HUTestDataRoot $AppRoot) $Addon.Id
    if (-not (Test-HUTestDataPath $dst $AppRoot)) { throw 'Der Testbestand ist nicht gekennzeichnet.' }
    # leer = keine Dateien (leere Ordner, die das Modul bei einem frueheren Start angelegt hat, zaehlen nicht)
    if ((Test-Path -LiteralPath $dst) -and @(Get-ChildItem -LiteralPath $dst -Recurse -Force -File).Count) { return 0 }
    $n = 0
    foreach ($i in @(Get-ChildItem -LiteralPath $src -Recurse -Force)) {
        if ($i.Attributes -band [IO.FileAttributes]::ReparsePoint) { continue }
        $to = [IO.Path]::Combine($dst, $i.FullName.Substring($src.Length + 1))
        if ($i.PSIsContainer) { [void][IO.Directory]::CreateDirectory($to); continue }
        [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($to))
        [IO.File]::Copy($i.FullName, $to, $false); $n++
    }
    return $n
}

<#
.SYNOPSIS
    Testmodus starten (nur Verwalter) oder beenden. Beides startet die App neu.
#>
function Switch-HUTestMode {
    if (Test-HUTestMode) {
        if (Confirm-HU "Testmodus beenden?`n`n$($script:AppTitle) startet neu und arbeitet wieder mit den echten Daten. Der Testbestand bleibt bis zum Zurücksetzen erhalten." 'Testmodus') { Restart-HUApp -LeaveTestMode }
        return
    }
    if (-not (Test-HUVerwalter)) { Show-HUMessage 'Den Testmodus können nur Verwalter starten.' -Icon Warning; return }
    $txt = "Testmodus starten?`n`n$($script:AppTitle) startet neu und arbeitet bis zum nächsten normalen Neustart mit einem eigenen Testbestand (Data\_Test).`n`n" +
    "Echte Daten bleiben unberührt. Outlook, Word und Mails sind im Test gesperrt, die automatische Sicherung ruht."
    if (Confirm-HU $txt 'Testmodus') { Restart-HUApp -EnterTestMode }
}

<#
.SYNOPSIS
    Testdaten loeschen: Kennzeichnung fuer den naechsten Start setzen und im Testmodus neu starten.
    Geloescht wird erst beim Start, bevor ein Modul eine Datei offen hat.
#>
function Request-HUTestReset {
    if (-not (Test-HUTestMode)) { return }
    $root = Get-HUTestDataRoot
    if (-not (Test-HUTestDataPath $root)) { Show-HUMessage 'Der Testbestand ist nicht gekennzeichnet. Es wurde nichts gelöscht.' -Icon Error; return }
    if (-not (Confirm-HU "Alle Testdaten löschen und mit leerem Testbestand neu starten?`n`nEchte Daten bleiben unberührt." 'Testmodus' -Warning)) { return }
    [IO.File]::WriteAllText((Join-Path $root $script:HUTestResetMarker), (Get-Date -Format 'o'))
    Restart-HUApp
}
