#Requires -Version 5.1
<#
.SYNOPSIS
    HU.Office - Outlook (klassisch) und Word per COM fuer alle Module: Termine, Aufgaben, Mail anzeigen, Word-Vorlagen, .ics.
.DESCRIPTION
    Eine gepruefte Stelle fuer COM, damit nicht jedes Modul eigene Outlook- und Word-Aufrufe baut.
    Grundsaetze:
      - Nur klassisches Outlook. Das neue Outlook hat kein Objektmodell (Microsoft: "Feature comparison between new Outlook
        and classic Outlook"). Rueckfall: .ics-Datei bzw. mailto:.
      - Nie automatisch senden: Mails werden nur angezeigt (MailItem.Display).
      - Kein Prozess bleibt haengen: Alle COM-Objekte werden freigegeben. Outlook wird nur beendet, wenn es vorher nicht lief
        und kein Fenster offen ist. Word laeuft immer als eigene, unsichtbare Instanz und wird am Ende beendet.
      - Doppelte Eintraege gibt es nicht: Jeder Termin und jede Aufgabe traegt einen Schluessel in der Benutzereigenschaft
        'HUAmtSchluessel' (z. B. 'Fristen:F-2026-0012'). Set-* sucht zuerst ueber EntryID, dann ueber den Schluessel,
        und aendert den gefundenen Eintrag statt einen zweiten anzulegen.
    Doku (Microsoft Learn, learn.microsoft.com/office/vba/api/...):
      outlook.namespace.pickfolder (Nothing bei Abbruch), outlook.namespace.getfolderfromid, outlook.namespace.getitemfromid,
      outlook.items.add (legt im Ordner der Items-Auflistung an), outlook.items.restrict (benutzerdefinierte Felder muessen
      im Ordner existieren, sonst Fehler), outlook.userproperties.add (olText = 1, AddToFolderFields Standard True,
      Name ohne [ ] _ #), outlook.taskitem.markcomplete, outlook.olitemtype, outlook.oldefaultfolders,
      outlook.explorers.count, outlook.inspectors.count,
      word.document.exportasfixedformat (wdExportFormatPDF = 17), word.wdsaveoptions (wdDoNotSaveChanges = 0),
      word.wdalertlevel (wdAlertsNone = 0), word.document.storyranges, word.range.nextstoryrange, word.tables.add,
      word.wdfindwrap (wdFindStop = 0). RFC 5545 fuer .ics (CRLF, Falten bei 75 Oktetten, DTSTAMP Pflicht, TRIGGER:-PT15M).
    Unsicher (nicht in der Doku belegt) und darum mit Rueckfall gebaut:
      - Find/Restrict ueber einen Schluessel, der nicht als Ordnerfeld existiert: Fehler wird abgefangen, dann Durchsuchen.
      - Laenge des Ersetzungstexts in Find.Execute: wird nicht genutzt, der Fundbereich bekommt den Text direkt (Range.Text).
.NOTES
    Zielmaschine: der PC, auf dem die App laeuft (Windows PowerShell 5.1 und PowerShell 7). COM braucht einen STA-Thread:
    den Fenster-Thread oder den Runspace-Pool des Kerns (Invoke-AsyncCommand, STA). PickFolder und Display nur im Fenster-Thread.
#>

Set-StrictMode -Version 1.0

$script:HUOfKeyProperty = 'HUAmtSchluessel'
$script:HUOfKeyPattern = '^[A-Za-z0-9][A-Za-z0-9:._-]{0,119}$'
$script:HUOfItemType = @{ Mail = 0; Appointment = 1; Task = 3 }          # OlItemType
$script:HUOfDefaultFolder = @{ Calendar = 9; Tasks = 13 }                # OlDefaultFolders
$script:HUOfBusy = @{ Free = 0; Tentative = 1; Busy = 2; OutOfOffice = 3 } # OlBusyStatus

# ---------------------------------------------------------------- Allgemein ----------------------------------------------------------------

function Test-HUComProgId([string]$ProgId) {
    try {
        $k = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey("$ProgId\CLSID")
        if (-not $k) { return $false }
        try { return [bool]"$($k.GetValue(''))" } finally { $k.Close() }
    } catch { return $false }
}

<#
.SYNOPSIS
    $true, wenn das klassische Outlook installiert ist (registrierte ProgID). Das neue Outlook zaehlt nicht.
#>
function Test-HUOutlookClassic { return (Test-HUComProgId 'Outlook.Application') }

<#
.SYNOPSIS
    $true, wenn Word installiert ist (registrierte ProgID).
#>
function Test-HUWord { return (Test-HUComProgId 'Word.Application') }

<#
.SYNOPSIS
    COM-Objekte freigeben. $null und Nicht-COM-Objekte werden uebersprungen.
#>
function Remove-HUComObject {
    param([object[]]$InputObject)
    foreach ($o in @($InputObject)) {
        if ($null -eq $o) { continue }
        try { if ([Runtime.InteropServices.Marshal]::IsComObject($o)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) } } catch { }
    }
}

<#
.SYNOPSIS
    Schluessel pruefen: Buchstaben, Ziffern und : . _ -, hoechstens 120 Zeichen (kommt in einen Filter, darum streng).
#>
function Test-HUOfficeKey([string]$Key) { return ("$Key" -match $script:HUOfKeyPattern) }

<#
.SYNOPSIS
    $true im Testmodus der App ($env:HUAMT_TESTMODE = 1, gilt auch in Hintergrund-Auftraegen).
    Dann sind Outlook, Word und Mail gesperrt: Open-HUOutlook und Invoke-HUWordTemplate werfen, Show-HUMail oeffnet nichts.
#>
function Test-HUOfficeTestMode { return ($env:HUAMT_TESTMODE -eq '1') }

function Assert-HUOfficeLive([string]$What) {
    if (Test-HUOfficeTestMode) { throw "Testmodus: $What ist gesperrt (keine echten Outlook-, Word- oder Mail-Aktionen)." }
}

# COM-Eigenschaft mit einem Objekt setzen. Direkte Zuweisung klappt nicht in jeder PowerShell-Version zuverlaessig, darum InvokeMember.
function Set-HUComProperty($Target, [string]$Name, $Value) {
    try { [void]$Target.GetType().InvokeMember($Name, [Reflection.BindingFlags]::SetProperty, $null, $Target, @($Value)) }
    catch { $Target.$Name = $Value }
}

function Assert-HUOfficeKey([string]$Key) {
    if (-not (Test-HUOfficeKey $Key)) { throw "Ungueltiger Schluessel '$Key' (erlaubt: Buchstaben, Ziffern und : . _ -, hoechstens 120 Zeichen)." }
}

# ---------------------------------------------------------------- Outlook: Sitzung ----------------------------------------------------------------

# RPC-Fehler, bei denen Outlook nur kurz nicht annimmt: beim kalten Start oder waehrend es sich nach Quit beendet.
# 0x80010001 RPC_E_CALL_REJECTED, 0x800706BA RPC-Server nicht verfuegbar, 0x800706BE Remoteprozeduraufruf fehlgeschlagen.
$script:HUOfRetryHResults = @(-2147418111, -2147023174, -2147023170)
$script:HUOfCreateApp = { New-Object -ComObject Outlook.Application -ErrorAction Stop }   # im Test ersetzbar
$script:HUOfRetryDelayMs = 2000

function Get-HUComHResult($ErrorRecordOrException) {
    $ex = $(if ($ErrorRecordOrException -is [System.Management.Automation.ErrorRecord]) { $ErrorRecordOrException.Exception } else { $ErrorRecordOrException })
    while ($ex) {
        if ($ex.HResult -and $script:HUOfRetryHResults -contains [int]$ex.HResult) { return [int]$ex.HResult }
        $ex = $ex.InnerException
    }
    return 0
}

# Skriptblock ausfuehren und bei den RPC-Fehlern oben bis -TimeoutSeconds wiederholen. Andere Fehler gehen sofort weiter.
function Invoke-HUComRetry([scriptblock]$Action, [int]$TimeoutSeconds = 30) {
    $until = (Get-Date).AddSeconds($TimeoutSeconds)
    while ($true) {
        try { return (& $Action) }
        catch {
            if ((Get-HUComHResult $_) -eq 0 -or (Get-Date) -ge $until) { throw }
            Start-Sleep -Milliseconds $script:HUOfRetryDelayMs
        }
    }
}

<#
.SYNOPSIS
    Outlook-Sitzung oeffnen. Rueckgabe: Application, Namespace, WasRunning, ProcessId. Immer mit Close-HUOutlook schliessen.
    Startet Outlook kalt oder beendet es sich gerade, lehnt es Aufrufe kurz ab (RPC_E_CALL_REJECTED u. a.):
    dann wird bis -TimeoutSeconds wiederholt (im Echttest aufgefallen).
#>
function Open-HUOutlook {
    param([int]$TimeoutSeconds = 30)
    Assert-HUOfficeLive 'Outlook'
    if (-not (Test-HUOutlookClassic)) { throw 'Das klassische Outlook ist nicht installiert.' }
    $before = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $was = ($before.Count -gt 0)
    $app = Invoke-HUComRetry -TimeoutSeconds $TimeoutSeconds -Action $script:HUOfCreateApp
    if ($null -eq $app) { throw 'Outlook konnte nicht gestartet werden.' }
    $ns = $null
    try { $ns = Invoke-HUComRetry -TimeoutSeconds $TimeoutSeconds -Action { $app.GetNamespace('MAPI') } }
    catch { Remove-HUComObject $app; throw }
    $procId = 0
    if (-not $was) {
        $new = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id })
        if ($new.Count -eq 1) { $procId = [int]$new[0].Id }
    }
    return [pscustomobject]@{ Application = $app; Namespace = $ns; WasRunning = $was; ProcessId = $procId }
}

<#
.SYNOPSIS
    Outlook-Sitzung schliessen. Beendet Outlook nur, wenn es vorher nicht lief und kein Fenster offen ist (z. B. eine angezeigte Mail).
    Nach Quit wird bis zu -WaitSeconds gewartet, bis der selbst gestartete Prozess weg ist. Sonst hielte die naechste Sitzung
    das sterbende Outlook fuer „lief schon“ und wuerde es nie beenden. Outlook wird nie hart beendet.
#>
function Close-HUOutlook {
    param($Session, [int]$WaitSeconds = 15)
    if (-not $Session) { return }
    $app = $Session.Application
    $quit = $false
    try {
        if ($app -and -not $Session.WasRunning) {
            $ex = $null; $in = $null
            try {
                $ex = $app.Explorers; $in = $app.Inspectors
                if ([int]$ex.Count -eq 0 -and [int]$in.Count -eq 0) { Invoke-HUComRetry -TimeoutSeconds 10 -Action { $app.Quit() }; $quit = $true }
            } catch { } finally { Remove-HUComObject @($ex, $in) }
        }
    } finally {
        Remove-HUComObject @($Session.Namespace, $app)
        $Session.Namespace = $null; $Session.Application = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
    $pidProp = $Session.PSObject.Properties['ProcessId']
    if ($quit -and $pidProp -and [int]$pidProp.Value -gt 0) {
        # WaitForExit statt Get-Process: ein Prozess, der sich gerade beendet, ist per Id oft schon nicht mehr zu finden (PS 7)
        $pr = $null
        try { $pr = [System.Diagnostics.Process]::GetProcessById([int]$pidProp.Value); [void]$pr.WaitForExit($WaitSeconds * 1000) }
        catch { } finally { if ($pr) { $pr.Dispose() } }
    }
}

<#
.SYNOPSIS
    Skriptblock mit einer Outlook-Sitzung ausfuehren: & $ScriptBlock $Session. Die Sitzung wird immer geschlossen.
#>
function Invoke-HUOutlook {
    param([Parameter(Mandatory)][scriptblock]$ScriptBlock)
    $s = Open-HUOutlook
    try { return (& $ScriptBlock $s) } finally { Close-HUOutlook $s }
}

# Eigene Sitzung nur, wenn keine mitgegeben wurde.
# Ungewoehnliche Namen mit Absicht: Der Skriptblock sieht die Variablen dieser Funktion (dynamischer Gueltigkeitsbereich).
# Hiesse ein Parameter z. B. $Body, bekaeme der Termin den Skriptblock statt des Texts (im Echttest aufgefallen).
function Use-HUOutlookSession($__HUOfSession, [scriptblock]$__HUOfScript) {
    if ($__HUOfSession) { return (& $__HUOfScript $__HUOfSession) }
    $__HUOfOwn = Open-HUOutlook
    try { return (& $__HUOfScript $__HUOfOwn) } finally { Close-HUOutlook $__HUOfOwn }
}

# ---------------------------------------------------------------- Outlook: Ordner ----------------------------------------------------------------

<#
.SYNOPSIS
    Ordnerangabe vereinheitlichen: Name, Path, EntryId, StoreId, Type (Calendar | Tasks). Leer = Standardordner des Benutzers.
    So wird sie in den Einstellungen eines Moduls gespeichert.
#>
function ConvertTo-HUOutlookFolderSpec {
    param($InputObject, [ValidateSet('Calendar', 'Tasks')][string]$Type = 'Calendar')
    $g = { param($n) if ($InputObject -and $InputObject.PSObject.Properties[$n]) { "$($InputObject.$n)" } else { '' } }
    $t = & $g 'Type'
    if ($t -notin @('Calendar', 'Tasks')) { $t = $Type }
    return [pscustomobject]@{ Name = (& $g 'Name'); Path = (& $g 'Path'); EntryId = (& $g 'EntryId'); StoreId = (& $g 'StoreId'); Type = $t }
}

<#
.SYNOPSIS
    Ordner oeffnen: mit EntryId ueber GetFolderFromID (z. B. gemeinsamer Kalender), sonst der Standardordner (Kalender oder Aufgaben).
    Das zurueckgegebene COM-Objekt gibt der Aufrufer mit Remove-HUComObject frei.
#>
function Get-HUOutlookFolder {
    param([Parameter(Mandatory)]$Session, $Folder, [ValidateSet('Calendar', 'Tasks')][string]$Type = 'Calendar')
    $spec = ConvertTo-HUOutlookFolderSpec $Folder -Type $Type
    if ($spec.EntryId) {
        try {
            if ($spec.StoreId) { return $Session.Namespace.GetFolderFromID($spec.EntryId, $spec.StoreId) }
            return $Session.Namespace.GetFolderFromID($spec.EntryId)
        } catch { throw "Outlook-Ordner '$($spec.Name)' nicht gefunden. Bitte in den Einstellungen neu waehlen. ($($_.Exception.Message))" }
    }
    return $Session.Namespace.GetDefaultFolder($script:HUOfDefaultFolder[$spec.Type])
}

# Dialoge von Outlook (z. B. PickFolder) gehen oft hinter dem HU-Amt-Fenster auf, waehrend HU-Amt auf sie wartet.
# Darum: Outlook darf in den Vordergrund (AllowSetForegroundWindow, wir sind gerade vorne), und ein kurzer Waechter in einem
# eigenen Runspace holt das erste sichtbare Dialogfenster (#32770) des Outlook-Prozesses nach vorne (SetForegroundWindow).
function Initialize-HUOfficeWin32 {
    if ('HUOffice.Front' -as [type]) { return }
    Add-Type -Namespace HUOffice -Name Front -MemberDefinition @'
public delegate bool EnumProc(System.IntPtr h, System.IntPtr l);
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc f, System.IntPtr l);
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint pid);
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool IsWindowVisible(System.IntPtr h);
[System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)] public static extern int GetClassName(System.IntPtr h, System.Text.StringBuilder s, int n);
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool SetForegroundWindow(System.IntPtr h);
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool AllowSetForegroundWindow(uint pid);
public static System.IntPtr FindDialog(uint pid) {
    System.IntPtr found = System.IntPtr.Zero;
    EnumWindows(delegate (System.IntPtr h, System.IntPtr l) {
        uint p; GetWindowThreadProcessId(h, out p);
        if (p != pid || !IsWindowVisible(h)) return true;
        var c = new System.Text.StringBuilder(64); GetClassName(h, c, 64);
        if (c.ToString() == "#32770") { found = h; return false; }
        return true; }, System.IntPtr.Zero);
    return found;
}
'@
}

function Start-HUOutlookDialogFront {
    param($Session, [int]$Seconds = 30)
    try {
        $procId = 0
        if ($Session -and $Session.PSObject.Properties['ProcessId'] -and [int]$Session.ProcessId -gt 0) { $procId = [int]$Session.ProcessId }
        if (-not $procId) { $p = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue); if ($p.Count -ge 1) { $procId = [int]$p[0].Id } }
        if (-not $procId) { return $null }
        Initialize-HUOfficeWin32
        [void][HUOffice.Front]::AllowSetForegroundWindow([uint32]$procId)
        $ps = [powershell]::Create()
        [void]$ps.AddScript({
                param($procId, $sec)
                for ($i = 0; $i -lt ($sec * 5); $i++) {
                    $h = [HUOffice.Front]::FindDialog([uint32]$procId)
                    if ($h -ne [IntPtr]::Zero) { [void][HUOffice.Front]::SetForegroundWindow($h); return }
                    Start-Sleep -Milliseconds 200
                }
            }).AddArgument($procId).AddArgument($Seconds)
        return [pscustomobject]@{ PS = $ps; Handle = $ps.BeginInvoke() }
    } catch { return $null }
}

function Stop-HUOutlookDialogFront($Guard) {
    if (-not $Guard) { return }
    try { $Guard.PS.Stop() } catch { }
    try { $Guard.PS.Dispose() } catch { }
}

<#
.SYNOPSIS
    Mailkonten im Outlook-Profil dieses PCs (Namespace.Accounts): Address (SMTP), Name, Exchange.
    Gemeinsame Postfaecher ohne eigenes Konto stehen hier nicht; dafuer gibt es "im Auftrag" (Show-HUMail -FromMode imAuftrag).
#>
function Get-HUOutlookAccounts {
    param($Session = $null)
    return @(Use-HUOutlookSession $Session {
            param($s)
            $accs = $null
            try {
                $accs = $s.Namespace.Accounts
                foreach ($a in $accs) {
                    try { [pscustomobject]@{ Address = "$($a.SmtpAddress)"; Name = "$($a.DisplayName)"; Exchange = ([int]$a.AccountType -eq 0) } } finally { Remove-HUComObject $a }
                }
            } finally { Remove-HUComObject $accs }
        })
}

<#
.SYNOPSIS
    Kontakte aus dem Kontakte-Ordner des eigenen Outlook (olFolderContacts = 10), im Format von Core\HU.Contacts.psm1:
    Name, Firma, Mail, Telefon, Mobil, Strasse, Plz, Ort, Land, Quelle. Verteilerlisten werden uebersprungen (nur Class 40 = olContact).
    -Search filtert nach Name, Firma, Mail und Ort. -Max begrenzt die Anzahl.
#>
function Get-HUOutlookContacts {
    param([string]$Search = '', [int]$Max = 500, $Session = $null)
    $q = "$Search".Trim()
    return @(Use-HUOutlookSession $Session {
            param($s)
            $f = $null; $items = $null; $n = 0
            try {
                $f = $s.Namespace.GetDefaultFolder(10)
                $items = $f.Items
                foreach ($it in $items) {
                    try {
                        if ([int]$it.Class -ne 40) { continue }
                        $street = "$($it.MailingAddressStreet)"; $plz = "$($it.MailingAddressPostalCode)"; $city = "$($it.MailingAddressCity)"; $land = "$($it.MailingAddressCountry)"
                        if (-not ($street + $plz + $city)) { $street = "$($it.BusinessAddressStreet)"; $plz = "$($it.BusinessAddressPostalCode)"; $city = "$($it.BusinessAddressCity)"; $land = "$($it.BusinessAddressCountry)" }
                        if (-not ($street + $plz + $city)) { $street = "$($it.HomeAddressStreet)"; $plz = "$($it.HomeAddressPostalCode)"; $city = "$($it.HomeAddressCity)"; $land = "$($it.HomeAddressCountry)" }
                        $tel = "$($it.BusinessTelephoneNumber)"; if (-not $tel) { $tel = "$($it.HomeTelephoneNumber)" }
                        $name = ("$($it.FullName)" -replace '\s+', ' ').Trim(); $firma = ("$($it.CompanyName)").Trim()
                        if (-not $name) { $name = $firma }
                        $mail = "$($it.Email1Address)".Trim()
                        # IndexOf statt -like: [ ] * ? im Suchtext sind kein Platzhalter (KERN-B7)
                        if ($q -and "$name $firma $mail $city".IndexOf($q, [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
                        [pscustomobject]@{ Name = $name; Firma = $firma; Mail = $mail; Telefon = $tel.Trim(); Mobil = "$($it.MobileTelephoneNumber)".Trim()
                            Strasse = ($street -replace '\s+', ' ').Trim(); Plz = $plz.Trim(); Ort = $city.Trim(); Land = $land.Trim(); Quelle = 'Outlook'
                        }
                        $n++
                        if ($n -ge $Max) { break }
                    } finally { Remove-HUComObject $it }
                }
            } finally { Remove-HUComObject @($items, $f) }
        })
}

function Find-HUOutlookAccount($Session, [string]$Address) {
    $accs = $Session.Namespace.Accounts
    try {
        foreach ($a in $accs) {
            if ("$($a.SmtpAddress)".Trim().Equals($Address.Trim(), [StringComparison]::OrdinalIgnoreCase)) { return $a }
            Remove-HUComObject $a
        }
        return $null
    } finally { Remove-HUComObject $accs }
}

<#
.SYNOPSIS
    Ordner im Outlook-Dialog waehlen (nur im Fenster-Thread). Rueckgabe: Ordnerangabe oder $null bei Abbruch.
    Passt der Ordnertyp nicht (z. B. Posteingang statt Kalender), wirft die Funktion mit einer verstaendlichen Meldung.
#>
function Select-HUOutlookFolder {
    param([ValidateSet('Calendar', 'Tasks')][string]$Type = 'Calendar', $Session = $null)
    return (Use-HUOutlookSession $Session {
            param($s)
            $guard = Start-HUOutlookDialogFront -Session $s
            try { $f = $s.Namespace.PickFolder() } finally { Stop-HUOutlookDialogFront $guard }
            if ($null -eq $f) { return $null }
            try {
                $want = $(if ($Type -eq 'Tasks') { $script:HUOfItemType.Task } else { $script:HUOfItemType.Appointment })
                if ([int]$f.DefaultItemType -ne $want) { throw "Der Ordner '$($f.Name)' ist kein $(if ($Type -eq 'Tasks') { 'Aufgaben' } else { 'Kalender' })ordner." }
                return [pscustomobject]@{ Name = "$($f.Name)"; Path = "$($f.FolderPath)"; EntryId = "$($f.EntryID)"; StoreId = "$($f.StoreID)"; Type = $Type }
            } finally { Remove-HUComObject $f }
        })
}

# ---------------------------------------------------------------- Outlook: Eintraege ----------------------------------------------------------------

function Get-HUOutlookItemKey($Item) {
    $ups = $null; $p = $null
    try {
        $ups = $Item.UserProperties
        $p = $ups.Find($script:HUOfKeyProperty)
        if ($null -eq $p) { return '' }
        return "$($p.Value)"
    } catch { return '' } finally { Remove-HUComObject @($p, $ups) }
}

function Set-HUOutlookItemKey($Item, [string]$Key) {
    $ups = $null; $p = $null
    try {
        $ups = $Item.UserProperties
        $p = $ups.Find($script:HUOfKeyProperty)
        if ($null -eq $p) { $p = $ups.Add($script:HUOfKeyProperty, 1, $true) }   # olText, als Ordnerfeld (fuer Restrict)
        $p.Value = $Key
    } finally { Remove-HUComObject @($p, $ups) }
}

<#
.SYNOPSIS
    Eintrag ueber den Schluessel finden. Reihenfolge: EntryID (falls bekannt und Schluessel passt), Restrict auf das Ordnerfeld,
    zuletzt Durchsuchen. Rueckgabe: COM-Objekt oder $null (Aufrufer gibt es mit Remove-HUComObject frei).
#>
function Find-HUOutlookItem {
    param([Parameter(Mandatory)]$Session, [Parameter(Mandatory)]$FolderObject, [Parameter(Mandatory)][string]$Key, [string]$EntryId = '')
    Assert-HUOfficeKey $Key
    if ($EntryId) {
        $it = $null
        try {
            $sid = "$($FolderObject.StoreID)"
            $it = $(if ($sid) { $Session.Namespace.GetItemFromID($EntryId, $sid) } else { $Session.Namespace.GetItemFromID($EntryId) })
        } catch { $it = $null }
        if ($it) {
            if ((Get-HUOutlookItemKey $it) -eq $Key) { return $it }
            Remove-HUComObject $it
        }
    }
    $items = $null; $hits = $null
    try {
        $items = $FolderObject.Items
        $found = $null
        try {
            $hits = $items.Restrict("[$($script:HUOfKeyProperty)] = '$Key'")
            foreach ($h in $hits) {
                if ($null -eq $found -and (Get-HUOutlookItemKey $h) -eq $Key) { $found = $h } else { Remove-HUComObject $h }
            }
            if ($found) { return $found }
        } catch { }
        # Rueckfall: Feld existiert im Ordner noch nicht (Restrict wirft) oder Restrict findet nichts
        foreach ($h in $items) {
            if ($null -eq $found -and (Get-HUOutlookItemKey $h) -eq $Key) { $found = $h } else { Remove-HUComObject $h }
        }
        return $found
    } finally { Remove-HUComObject @($hits, $items) }
}

function New-HUOutlookItemResult($Item, [bool]$Created) {
    return [pscustomobject]@{ EntryId = "$($Item.EntryID)"; StoreId = $(try { "$($Item.Parent.StoreID)" } catch { '' }); Created = $Created }
}

<#
.SYNOPSIS
    Termin anlegen oder denselben aendern (Schluessel). Rueckgabe: EntryId, StoreId, Created.
    -Folder: Ordnerangabe (leer = eigener Kalender). -ReminderMinutes -1 = keine Erinnerung. -End leer = Start + 1 Stunde.
    Ganztaegig: Start und End sind Tage, End ist der letzte Tag (inklusive).
#>
function Set-HUOutlookAppointment {
    param(
        $Session = $null, $Folder = $null,
        [Parameter(Mandatory)][string]$Key, [string]$EntryId = '',
        [Parameter(Mandatory)][string]$Subject, [Parameter(Mandatory)][datetime]$Start, [Nullable[datetime]]$End = $null,
        [switch]$AllDay, [string]$Location = '', [string]$Body = '', [int]$ReminderMinutes = -1,
        [ValidateSet('Free', 'Tentative', 'Busy', 'OutOfOffice')][string]$BusyStatus = 'Busy', [string]$Categories = ''
    )
    Assert-HUOfficeKey $Key
    return (Use-HUOutlookSession $Session {
            param($s)
            $fo = Get-HUOutlookFolder -Session $s -Folder $Folder -Type Calendar
            $it = $null; $items = $null
            try {
                $it = Find-HUOutlookItem -Session $s -FolderObject $fo -Key $Key -EntryId $EntryId
                $created = $false
                if (-not $it) { $items = $fo.Items; $it = $items.Add($script:HUOfItemType.Appointment); $created = $true }
                $it.Subject = $Subject
                if ($AllDay) {
                    $last = $(if ($End) { ([datetime]$End).Date } else { $Start.Date })
                    $it.AllDayEvent = $true; $it.Start = $Start.Date; $it.End = $last.AddDays(1)
                } else {
                    $e = $(if ($End) { ([datetime]$End) } else { $Start.AddHours(1) })
                    if ($e -le $Start) { throw 'Das Ende liegt nicht nach dem Beginn.' }
                    $it.AllDayEvent = $false; $it.Start = $Start; $it.End = $e
                }
                $it.Location = $Location
                $it.Body = $Body
                $it.BusyStatus = $script:HUOfBusy[$BusyStatus]
                $it.Categories = $Categories
                if ($ReminderMinutes -ge 0) { $it.ReminderSet = $true; $it.ReminderMinutesBeforeStart = $ReminderMinutes } else { $it.ReminderSet = $false }
                Set-HUOutlookItemKey $it $Key
                $it.Save()
                return (New-HUOutlookItemResult $it $created)
            } finally { Remove-HUComObject @($it, $items, $fo) }
        })
}

<#
.SYNOPSIS
    Aufgabe anlegen oder dieselbe aendern (Schluessel). Rueckgabe: EntryId, StoreId, Created.
    -Folder leer = eigene Aufgaben. -ReminderTime leer = keine Erinnerung.
#>
function Set-HUOutlookTask {
    param(
        $Session = $null, $Folder = $null,
        [Parameter(Mandatory)][string]$Key, [string]$EntryId = '',
        [Parameter(Mandatory)][string]$Subject, [Parameter(Mandatory)][datetime]$DueDate, [Nullable[datetime]]$StartDate = $null,
        [Nullable[datetime]]$ReminderTime = $null, [string]$Body = '', [string]$Categories = ''
    )
    Assert-HUOfficeKey $Key
    return (Use-HUOutlookSession $Session {
            param($s)
            $fo = Get-HUOutlookFolder -Session $s -Folder $Folder -Type Tasks
            $it = $null; $items = $null
            try {
                $it = Find-HUOutlookItem -Session $s -FolderObject $fo -Key $Key -EntryId $EntryId
                $created = $false
                if (-not $it) { $items = $fo.Items; $it = $items.Add($script:HUOfItemType.Task); $created = $true }
                $it.Subject = $Subject
                if ($StartDate) { $it.StartDate = ([datetime]$StartDate).Date }
                $it.DueDate = $DueDate.Date
                $it.Body = $Body
                $it.Categories = $Categories
                if ($ReminderTime) { $it.ReminderSet = $true; $it.ReminderTime = ([datetime]$ReminderTime) } else { $it.ReminderSet = $false }
                Set-HUOutlookItemKey $it $Key
                $it.Save()
                return (New-HUOutlookItemResult $it $created)
            } finally { Remove-HUComObject @($it, $items, $fo) }
        })
}

<#
.SYNOPSIS
    Eintrag als erledigt markieren. Aufgabe: MarkComplete. Termin: Erinnerung aus, Betreff mit -Prefix davor.
    Rueckgabe: $true, wenn der Eintrag gefunden wurde.
#>
function Complete-HUOutlookItem {
    param($Session = $null, $Folder = $null, [ValidateSet('Task', 'Appointment')][string]$Type = 'Task',
        [Parameter(Mandatory)][string]$Key, [string]$EntryId = '', [string]$Prefix = 'Erledigt: ')
    Assert-HUOfficeKey $Key
    return (Use-HUOutlookSession $Session {
            param($s)
            $fo = Get-HUOutlookFolder -Session $s -Folder $Folder -Type $(if ($Type -eq 'Task') { 'Tasks' } else { 'Calendar' })
            $it = $null
            try {
                $it = Find-HUOutlookItem -Session $s -FolderObject $fo -Key $Key -EntryId $EntryId
                if (-not $it) { return $false }
                if ($Type -eq 'Task') { $it.MarkComplete() }
                else {
                    $it.ReminderSet = $false
                    if ($Prefix -and -not "$($it.Subject)".StartsWith($Prefix)) { $it.Subject = $Prefix + "$($it.Subject)" }
                }
                $it.Save()
                return $true
            } finally { Remove-HUComObject @($it, $fo) }
        })
}

<#
.SYNOPSIS
    Eintrag loeschen (z. B. Storno einer Buchung). Rueckgabe: $true, wenn er gefunden und geloescht wurde.
#>
function Remove-HUOutlookItem {
    param($Session = $null, $Folder = $null, [ValidateSet('Task', 'Appointment')][string]$Type = 'Appointment',
        [Parameter(Mandatory)][string]$Key, [string]$EntryId = '')
    Assert-HUOfficeKey $Key
    return (Use-HUOutlookSession $Session {
            param($s)
            $fo = Get-HUOutlookFolder -Session $s -Folder $Folder -Type $(if ($Type -eq 'Task') { 'Tasks' } else { 'Calendar' })
            $it = $null
            try {
                $it = Find-HUOutlookItem -Session $s -FolderObject $fo -Key $Key -EntryId $EntryId
                if (-not $it) { return $false }
                $it.Delete()
                return $true
            } finally { Remove-HUComObject @($it, $fo) }
        })
}

# ---------------------------------------------------------------- Mail ----------------------------------------------------------------

<#
.SYNOPSIS
    mailto:-Adresse mit kodiertem Betreff und Text. Ist sie laenger als -MaxLength, wird der Text gekuerzt und -Note angehaengt.
    Die Grenze ist nicht dokumentiert, 1800 Zeichen sind vorsichtig gewaehlt.
#>
function New-HUMailtoUri {
    param([string]$To = '', [string]$Subject = '', [string]$Body = '', [int]$MaxLength = 1800, [string]$Note = '')
    $enc = { param($t) [Uri]::EscapeDataString("$t") }
    $head = 'mailto:' + $(if ($To) { (& $enc $To.Trim()) -replace '%40', '@' } else { '' }) + '?subject=' + (& $enc $Subject) + '&body='
    $full = $head + (& $enc $Body)
    if ($full.Length -le $MaxLength) { return $full }
    $tail = "`r`n`r`n" + $Note
    $lo = 0; $hi = $Body.Length; $best = ''
    while ($lo -le $hi) {
        $mid = [int][Math]::Floor(($lo + $hi) / 2)
        $u = $head + (& $enc ($Body.Substring(0, $mid).TrimEnd() + ' ...' + $tail))
        if ($u.Length -le $MaxLength) { $best = $u; $lo = $mid + 1 } else { $hi = $mid - 1 }
    }
    if (-not $best) { $best = $head + (& $enc $Note) }
    if ($best.Length -gt $MaxLength) { $best = $head }
    return $best
}

<#
.SYNOPSIS
    Posteingang eines anderen Postfachs nach Mailadresse oeffnen (COM-Objekt, der Aufrufer gibt es frei).
    1. Konto im Outlook-Profil mit dieser Adresse: Posteingang seines Postfachs (Account.DeliveryStore).
    2. sonst freigegebenes Postfach (Exchange, Leserecht noetig): Namespace.GetSharedDefaultFolder.
    Wirft mit verstaendlicher Meldung, wenn die Adresse unbekannt ist oder das Recht fehlt.
#>
function Get-HUOutlookMailboxInbox {
    param([Parameter(Mandatory)]$Session, [Parameter(Mandatory)][string]$Mailbox)
    $addr = $Mailbox.Trim()
    if ($addr -notmatch '^[^@\s;,]+@[^@\s;,]+\.[^@\s;,]+$') { throw "'$addr' ist keine gültige Mailadresse." }
    $acc = $null; $store = $null
    try {
        $acc = Find-HUOutlookAccount $Session $addr
        if ($acc) {
            try { $store = $acc.DeliveryStore } catch { $store = $null }
            if ($store) { return $store.GetDefaultFolder(6) }   # olFolderInbox
        }
    } catch { throw "Posteingang von $addr nicht lesbar: $($_.Exception.Message)" }
    finally { Remove-HUComObject @($store, $acc) }
    $rcp = $null
    try {
        # Resolve wirft selbst, wenn Outlook offline ist oder das Adressbuch nicht erreicht: verstaendliche Meldung statt COM-Text
        try { $rcp = $Session.Namespace.CreateRecipient($addr); [void]$rcp.Resolve() }
        catch { throw "Das Postfach $addr lässt sich gerade nicht prüfen (Outlook offline oder Adressbuch nicht erreichbar). ($($_.Exception.Message))" }
        if (-not $rcp.Resolved) { throw "Das Postfach $addr ist in Outlook nicht bekannt." }
        try { return $Session.Namespace.GetSharedDefaultFolder($rcp, 6) }
        catch { throw "Auf den Posteingang von $addr gibt es kein Leserecht oder er ist nicht erreichbar. ($($_.Exception.Message))" }
    } finally { Remove-HUComObject $rcp }
}

<#
.SYNOPSIS
    Anhaenge aus Mails holen (nur lesen): Posteingang, neueste zuerst, hoechstens -Days alt. Die Mails bleiben unberuehrt
    (nicht verschoben, nicht als gelesen markiert). Passende Anhaenge werden nach -TargetDir gespeichert (eindeutige Namen).
    Rueckgabe: je Anhang EntryId, StoreId, Subject, Sender, SenderAddress, Received (DateTime), FileName, Path.
    Anhangsnamen kommen von aussen: nur der Dateiname zaehlt, unzulaessige Zeichen werden ersetzt, gespeichert wird nur
    innerhalb von -TargetDir. Ein Fehler bei einem Anhang bricht die anderen nicht ab.
.PARAMETER Problems
    Optional eine Liste ([System.Collections.Generic.List[string]]), in die Fehler je Anhang und der Hinweis
    "Hoechstzahl an Mails erreicht" geschrieben werden (fuer die Anzeige im Modul).
.PARAMETER Folder
    Wie bei Get-HUOutlookFolder eine gespeicherte Ordnerangabe mit EntryId; leer = Posteingang des Standardprofils.
.PARAMETER Mailbox
    Ab 1.0.4: Mailadresse eines anderen Postfachs (zweites Konto im Profil oder freigegebenes Postfach), gelesen wird
    dessen Posteingang (Get-HUOutlookMailboxInbox). -Folder hat Vorrang. StoreId der Ergebnisse ist die des Postfachs.
.NOTES
    Outlook fragt beim Zugriff auf Anhaenge nur ohne aktuellen Virenschutz nach (Object Model Guard).
#>
function Get-HUOutlookAttachments {
    param($Session = $null, $Folder = $null, [int]$Days = 30, [string[]]$Extensions = @('.pdf'), [Parameter(Mandatory)][string]$TargetDir, [int]$MaxMails = 500,
        $Problems = $null, [string]$Mailbox = '')
    if (-not (Test-Path -LiteralPath $TargetDir)) { [void](New-Item -ItemType Directory -Path $TargetDir -Force) }
    $rootDir = [IO.Path]::GetFullPath($TargetDir).TrimEnd('\') + '\'
    $bad = @([IO.Path]::GetInvalidFileNameChars())
    $ext = @($Extensions | ForEach-Object { "$_".ToLowerInvariant() })
    $limit = (Get-Date).AddDays(-[Math]::Max(1, $Days))
    return (Use-HUOutlookSession $Session {
            param($s)
            $out = New-Object 'System.Collections.Generic.List[object]'
            $fld = $null; $items = $null
            try {
                $eid = $(if ($Folder -and $Folder.PSObject.Properties['EntryId']) { "$($Folder.EntryId)" } else { '' })
                if ($eid) {
                    $sid = $(if ($Folder.PSObject.Properties['StoreId']) { "$($Folder.StoreId)" } else { '' })
                    $fld = $(if ($sid) { $s.Namespace.GetFolderFromID($eid, $sid) } else { $s.Namespace.GetFolderFromID($eid) })
                } elseif ("$Mailbox".Trim()) { $fld = Get-HUOutlookMailboxInbox -Session $s -Mailbox $Mailbox }
                else { $fld = $s.Namespace.GetDefaultFolder(6) }   # olFolderInbox
                $items = $fld.Items
                $items.Sort('[ReceivedTime]', $true)
                $n = 0
                $it = $items.GetFirst()
                while ($null -ne $it) {
                    if ($n -ge $MaxMails) { if ($null -ne $Problems) { [void]$Problems.Add("Es wurden nur die neuesten $MaxMails Mails durchsucht.") }; break }
                    $n++
                    $atts = $null
                    try {
                        if ([int]$it.Class -ne 43) { continue }   # olMail
                        $rcv = [datetime]$it.ReceivedTime
                        if ($rcv -lt $limit) { break }
                        $atts = $it.Attachments
                        for ($k = 1; $k -le [int]$atts.Count; $k++) {
                            $a = $null; $fn = ''
                            try {
                                $a = $atts.Item($k)
                                $fn = "$($a.FileName)"
                                # nur der letzte Namensteil, unzulaessige Zeichen ersetzen, erst danach die Endung bestimmen
                                $safe = (($fn -split '[\\/]')[-1]).Trim()
                                foreach ($ch in $bad) { $safe = $safe.Replace([string]$ch, '_') }
                                if ($safe.Length -gt 120) { $safe = $safe.Substring($safe.Length - 120) }
                                $e = [IO.Path]::GetExtension($safe).ToLowerInvariant()
                                if (-not $safe -or $ext -notcontains $e) { continue }
                                $base = [IO.Path]::GetFileNameWithoutExtension($safe); if (-not $base.Trim('.', ' ')) { $base = 'Anhang' }
                                $dst = Join-Path $TargetDir "$base$e"; $i = 2
                                while (Test-Path -LiteralPath $dst) { $dst = Join-Path $TargetDir "$($base)_$i$e"; $i++ }
                                if (-not ([IO.Path]::GetFullPath($dst)).StartsWith($rootDir, [StringComparison]::OrdinalIgnoreCase)) { throw 'unzulässiger Dateiname' }
                                $a.SaveAsFile($dst)
                                $addr = ''; try { $addr = "$($it.SenderEmailAddress)" } catch { }
                                $out.Add([pscustomobject]@{ EntryId = "$($it.EntryID)"; StoreId = "$($fld.StoreID)"; Subject = "$($it.Subject)"; Sender = "$($it.SenderName)"
                                        SenderAddress = $addr; Received = $rcv; FileName = $fn; Path = $dst })
                            } catch {
                                if ($null -ne $Problems) { [void]$Problems.Add("Anhang $(if ($fn) { "'$fn' " })in einer Mail vom $($rcv.ToString('dd.MM.yyyy HH:mm')) nicht lesbar: $($_.Exception.Message)") }
                            } finally { Remove-HUComObject $a }
                        }
                    } finally {
                        Remove-HUComObject $atts
                        $prev = $it; $it = $items.GetNext(); Remove-HUComObject $prev
                    }
                }
                if ($null -ne $it) { Remove-HUComObject $it }
            } finally { Remove-HUComObject @($items, $fld) }
            return $out.ToArray()
        })
}

# ---------------------------------------------------------------- Hineinziehen (auch Anhaenge direkt aus Outlook) -----------------------
# Aus dem Explorer kommt FileDrop (Pfade). Aus Outlook (Anhang oder ganze Mail) kommt kein FileDrop, sondern
# FileGroupDescriptorW (Namen und Groessen) und FileContents (Inhalt je Index). Der Inhalt wird in einen Ordner geschrieben.

<#
.SYNOPSIS
    FileGroupDescriptorW lesen: UINT Anzahl, dann je Datei ein FILEDESCRIPTORW mit 592 Byte
    (Name als UTF-16 ab Byte 72, 260 Zeichen; Groesse ab Byte 64/68, gueltig mit FD_FILESIZE 0x40).
    Rueckgabe je Datei: Index, Name (nur der Dateiname), Size (-1 = unbekannt).
#>
function Read-HUFileGroupDescriptor([byte[]]$Bytes) {
    if (-not $Bytes -or $Bytes.Length -lt 4) { return @() }
    $n = [BitConverter]::ToUInt32($Bytes, 0)
    $out = @()
    for ($i = 0; $i -lt $n -and $i -lt 100; $i++) {
        $o = 4 + 592 * $i
        if ($o + 592 -gt $Bytes.Length) { break }
        $flags = [BitConverter]::ToUInt32($Bytes, $o)
        $name = [Text.Encoding]::Unicode.GetString($Bytes, $o + 72, 520)
        $z = $name.IndexOf([char]0); if ($z -ge 0) { $name = $name.Substring(0, $z) }
        $name = (($name -split '[\\/]')[-1]).Trim()
        $size = [long]-1
        if ($flags -band 0x40) { $size = ([long][BitConverter]::ToUInt32($Bytes, $o + 64) -shl 32) + [long][BitConverter]::ToUInt32($Bytes, $o + 68) }
        $out += [pscustomobject]@{ Index = $i; Name = $name; Size = $size }
    }
    return $out
}

function Get-HUDropDescriptor($Data) {
    try {
        if (-not $Data.GetDataPresent('FileGroupDescriptorW')) { return @() }
        $s = $Data.GetData('FileGroupDescriptorW')
        if ($s -is [IO.MemoryStream]) { return @(Read-HUFileGroupDescriptor $s.ToArray()) }
        if ($s -is [byte[]]) { return @(Read-HUFileGroupDescriptor $s) }
    } catch { }
    return @()
}

function Initialize-HUDropNative {
    if ('HUDrop.Contents' -as [type]) { return }
    Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
namespace HUDrop {
    public static class Contents {
        [DllImport("kernel32.dll")] static extern IntPtr GlobalLock(IntPtr h);
        [DllImport("kernel32.dll")] static extern bool GlobalUnlock(IntPtr h);
        [DllImport("kernel32.dll")] static extern UIntPtr GlobalSize(IntPtr h);
        [DllImport("ole32.dll")] static extern void ReleaseStgMedium(ref STGMEDIUM m);
        // Inhalt der Datei Nr. index (FileContents) als Bytes; max = Obergrenze in Byte
        public static byte[] Get(object dataObject, short format, int index, long max) {
            System.Runtime.InteropServices.ComTypes.IDataObject com = dataObject as System.Runtime.InteropServices.ComTypes.IDataObject;
            if (com == null) throw new InvalidOperationException("Die gezogenen Daten lassen sich nicht lesen.");
            FORMATETC f = new FORMATETC();
            f.cfFormat = format; f.dwAspect = DVASPECT.DVASPECT_CONTENT; f.lindex = index; f.ptd = IntPtr.Zero;
            f.tymed = TYMED.TYMED_ISTREAM | TYMED.TYMED_HGLOBAL;
            STGMEDIUM m;
            com.GetData(ref f, out m);
            try {
                if (m.tymed == TYMED.TYMED_ISTREAM) {
                    IStream s = (IStream)Marshal.GetObjectForIUnknown(m.unionmember);
                    try {
                        MemoryStream ms = new MemoryStream();
                        byte[] buf = new byte[65536];
                        IntPtr pRead = Marshal.AllocCoTaskMem(4);
                        try {
                            while (true) {
                                s.Read(buf, buf.Length, pRead);
                                int n = Marshal.ReadInt32(pRead);
                                if (n <= 0) break;
                                ms.Write(buf, 0, n);
                                if (ms.Length > max) throw new InvalidOperationException("Die Datei ist zu groß.");
                            }
                        } finally { Marshal.FreeCoTaskMem(pRead); }
                        return ms.ToArray();
                    } finally { Marshal.ReleaseComObject(s); }
                }
                if (m.tymed == TYMED.TYMED_HGLOBAL) {
                    IntPtr p = GlobalLock(m.unionmember);
                    try {
                        long len = (long)GlobalSize(m.unionmember).ToUInt64();
                        if (len > max) throw new InvalidOperationException("Die Datei ist zu groß.");
                        byte[] b = new byte[len];
                        Marshal.Copy(p, b, 0, (int)len);
                        return b;
                    } finally { GlobalUnlock(m.unionmember); }
                }
                throw new InvalidOperationException("Unbekanntes Format der gezogenen Datei.");
            } finally { ReleaseStgMedium(ref m); }
        }
    }
}
'@
}

# Inhalt der gezogenen Datei Nr. Index (eigene Funktion, damit die Tests sie ersetzen koennen)
function Get-HUDropContent($Data, [int]$Index, [long]$MaxBytes = 104857600) {
    Initialize-HUDropNative
    $fmt = [int16][System.Windows.DataFormats]::GetDataFormat('FileContents').Id
    return , ([HUDrop.Contents]::Get($Data, $fmt, $Index, $MaxBytes))
}

<#
.SYNOPSIS
    Namen der gezogenen Dateien schon waehrend des Ziehens (DragOver): Pfade aus dem Explorer oder Namen der Outlook-Anhaenge.
#>
function Get-HUDropNames($Data) {
    try {
        if ($Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) { return @($Data.GetData([System.Windows.DataFormats]::FileDrop) | Where-Object { $_ }) }
    } catch { }
    return @(Get-HUDropDescriptor $Data | ForEach-Object { $_.Name } | Where-Object { $_ })
}

function Test-HUDropHasFiles($Data) {
    try { return ($Data.GetDataPresent([System.Windows.DataFormats]::FileDrop) -or $Data.GetDataPresent('FileGroupDescriptorW')) } catch { return $false }
}

<#
.SYNOPSIS
    Gezogene Dateien als Pfade (ab 1.0.4): aus dem Explorer die vorhandenen Dateien, aus Outlook die Anhaenge,
    geschrieben nach -TargetDir (eindeutige, bereinigte Namen). -Problems sammelt Texte zu Dateien, die nicht gingen.
    Ganze Mails aus Outlook kommen als .msg an; welche Endungen ein Modul nimmt, entscheidet das Modul.
#>
function Save-HUDroppedFiles {
    param([Parameter(Mandatory)]$Data, [Parameter(Mandatory)][string]$TargetDir, [long]$MaxBytes = 104857600, $Problems = $null)
    try {
        if ($Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
            return @($Data.GetData([System.Windows.DataFormats]::FileDrop) | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) })
        }
    } catch { }
    $desc = @(Get-HUDropDescriptor $Data)
    if (-not $desc.Count) { return @() }
    if (-not (Test-Path -LiteralPath $TargetDir)) { [void](New-Item -ItemType Directory -Path $TargetDir -Force) }
    $rootDir = [IO.Path]::GetFullPath($TargetDir).TrimEnd('\') + '\'
    $bad = @([IO.Path]::GetInvalidFileNameChars())
    $out = @()
    foreach ($d in $desc) {
        $safe = "$($d.Name)"
        foreach ($ch in $bad) { $safe = $safe.Replace([string]$ch, '_') }
        if ($safe.Length -gt 120) { $safe = $safe.Substring($safe.Length - 120) }
        $ext = [IO.Path]::GetExtension($safe)
        $base = [IO.Path]::GetFileNameWithoutExtension($safe); if (-not $base.Trim('.', ' ')) { $base = 'Anhang' }
        try {
            if ($d.Size -gt $MaxBytes) { throw 'Die Datei ist zu groß.' }
            $bytes = Get-HUDropContent -Data $Data -Index $d.Index -MaxBytes $MaxBytes
            if ($null -eq $bytes) { throw 'kein Inhalt' }
            # HGLOBAL kann groesser sein als die Datei: auf die Groesse laut Beschreibung kuerzen
            if ($d.Size -ge 0 -and $bytes.Length -gt $d.Size) { $cut = New-Object byte[] ([int]$d.Size); [Array]::Copy($bytes, $cut, [int]$d.Size); $bytes = $cut }
            $dst = Join-Path $TargetDir "$base$ext"; $i = 2
            while (Test-Path -LiteralPath $dst) { $dst = Join-Path $TargetDir "$($base)_$i$ext"; $i++ }
            if (-not ([IO.Path]::GetFullPath($dst)).StartsWith($rootDir, [StringComparison]::OrdinalIgnoreCase)) { throw 'unzulässiger Dateiname' }
            [IO.File]::WriteAllBytes($dst, $bytes)
            $out += $dst
        } catch {
            if ($null -ne $Problems) { [void]$Problems.Add("$($d.Name): $($_.Exception.Message)") }
        }
    }
    return $out
}

<#
.SYNOPSIS
    Mail im Outlook anzeigen, nie senden. Ohne Outlook (oder -Mailweg mailto): mailto: ohne Anhang.
    Rueckgabe: Way (outlook | mailto | fehler | testmodus), Message (Grund fuer den Rueckfall oder Fehler), From (verwendeter Absender).
    -From: Absender. -FromMode auto (Konto, sonst im Auftrag) | konto (nur ein Konto im Profil) | imAuftrag (gemeinsames Postfach).
    Nur im Fenster-Thread aufrufen (Display).
#>
function Show-HUMail {
    param([string[]]$To = @(), [string[]]$Cc = @(), [string]$Subject = '', [string]$Body = '', [string[]]$Attachments = @(),
        [ValidateSet('auto', 'mailto')][string]$Mailweg = 'auto', [int]$MaxLength = 1800, [string]$TruncatedNote = '(Text gekuerzt.)',
        [string]$From = '', [ValidateSet('auto', 'konto', 'imAuftrag')][string]$FromMode = 'auto')
    $toText = (@($To) | Where-Object { "$_".Trim() } | ForEach-Object { "$_".Trim() }) -join '; '
    $ccText = (@($Cc) | Where-Object { "$_".Trim() } | ForEach-Object { "$_".Trim() }) -join '; '
    if (Test-HUOfficeTestMode) { return [pscustomobject]@{ Way = 'testmodus'; Message = 'Testmodus: Es wurde keine Mail geöffnet.' } }
    $errText = ''
    if ($Mailweg -eq 'auto' -and (Test-HUOutlookClassic)) {
        $s = $null; $m = $null; $att = $null
        try {
            $s = Open-HUOutlook
            $m = $s.Application.CreateItem($script:HUOfItemType.Mail)
            if ($toText) { $m.To = $toText }
            if ($ccText) { $m.CC = $ccText }
            $m.Subject = $Subject
            $m.Body = $Body
            $usedFrom = ''
            if ($From) {
                # Konto im Profil: ueber dieses Konto senden. Sonst (gemeinsames Postfach) "im Auftrag von" - braucht das Recht im Exchange.
                $acc = $(if ($FromMode -ne 'imAuftrag') { Find-HUOutlookAccount $s $From } else { $null })
                if ($acc) {
                    try { Set-HUComProperty $m 'SendUsingAccount' $acc } finally { Remove-HUComObject $acc }
                    $usedFrom = "Konto $From"
                } elseif ($FromMode -eq 'konto') {
                    throw "Das Mailkonto $From gibt es im Outlook dieses PCs nicht."
                } else {
                    $m.SentOnBehalfOfName = $From
                    $usedFrom = "im Auftrag von $From"
                }
            }
            $att = $m.Attachments
            foreach ($a in @($Attachments)) {
                if (-not "$a") { continue }
                if (-not (Test-Path -LiteralPath $a)) { throw "Anhang fehlt: $a" }
                $x = $att.Add($a); Remove-HUComObject $x
            }
            $m.Display($false)
            return [pscustomobject]@{ Way = 'outlook'; Message = ''; From = $usedFrom }
        } catch { $errText = $_.Exception.Message }
        finally { Remove-HUComObject @($att, $m); Close-HUOutlook $s }
        # mit festem Absender nie still auf mailto ausweichen: dort laesst sich kein Absender setzen
        if ($From) { return [pscustomobject]@{ Way = 'fehler'; Message = "Outlook: $errText"; From = '' } }
    }
    try {
        $note = $TruncatedNote
        if (@($Attachments | Where-Object { "$_" }).Count -gt 0) { $Body = $Body + "`r`n`r`n(Anhang bitte selbst anfuegen.)" }
        $uri = New-HUMailtoUri -To $toText -Subject $Subject -Body $Body -MaxLength $MaxLength -Note $note
        Start-Process -FilePath $uri
        return [pscustomobject]@{ Way = 'mailto'; Message = $errText }
    } catch {
        return [pscustomobject]@{ Way = 'fehler'; Message = "$(if ($errText) { "Outlook: $errText; " })$($_.Exception.Message)" }
    }
}

# ---------------------------------------------------------------- .ics (RFC 5545) ----------------------------------------------------------------

# TEXT-Wert maskieren: Backslash, Semikolon, Komma, Zeilenwechsel (RFC 5545, 3.3.11)
function ConvertTo-HUIcsText([string]$Text) {
    $t = "$Text" -replace '\\', '\\' -replace ';', '\;' -replace ',', '\,'
    $t = $t -replace "`r`n", "`n" -replace "`r", "`n"
    return ($t -replace "`n", '\n')
}

# Zeile falten: hoechstens 75 Oktette (UTF-8), Fortsetzung mit CRLF + Leerzeichen, nie mitten in einem Zeichen
function Format-HUIcsLine([string]$Line) {
    $enc = New-Object System.Text.UTF8Encoding $false
    if ($enc.GetByteCount($Line) -le 75) { return $Line }
    $parts = New-Object System.Collections.Generic.List[string]
    $sb = New-Object System.Text.StringBuilder
    $limit = 75; $bytes = 0
    $i = 0
    while ($i -lt $Line.Length) {
        $len = $(if ([char]::IsHighSurrogate($Line[$i]) -and ($i + 1) -lt $Line.Length) { 2 } else { 1 })
        $ch = $Line.Substring($i, $len)
        $b = $enc.GetByteCount($ch)
        if ($bytes + $b -gt $limit) { $parts.Add($sb.ToString()); [void]$sb.Clear(); $limit = 74; $bytes = 0 }
        [void]$sb.Append($ch); $bytes += $b; $i += $len
    }
    $parts.Add($sb.ToString())
    return ($parts -join "`r`n ")
}

<#
.SYNOPSIS
    Inhalt einer .ics-Datei mit einem Termin (VEVENT, optional VALARM). Zeiten in UTC, ganztaegig als Datum (Ende exklusiv).
    -Uid bleibt fuer denselben Termin gleich, -Sequence zaehlt bei Aenderungen hoch.
#>
function New-HUIcsContent {
    param(
        [Parameter(Mandatory)][string]$Uid, [Parameter(Mandatory)][string]$Subject, [Parameter(Mandatory)][datetime]$Start,
        [Nullable[datetime]]$End = $null, [switch]$AllDay, [string]$Location = '', [string]$Description = '',
        [int]$AlarmMinutes = -1, [int]$Sequence = 0, [datetime]$Now = (Get-Date), [string]$ProdId = '-//HU-Amt//HU.Office//DE'
    )
    $utc = { param([datetime]$d) $d.ToUniversalTime().ToString("yyyyMMdd'T'HHmmss'Z'", [Globalization.CultureInfo]::InvariantCulture) }
    $day = { param([datetime]$d) $d.ToString('yyyyMMdd', [Globalization.CultureInfo]::InvariantCulture) }
    $l = New-Object System.Collections.Generic.List[string]
    $l.Add('BEGIN:VCALENDAR'); $l.Add('VERSION:2.0'); $l.Add("PRODID:$ProdId"); $l.Add('CALSCALE:GREGORIAN'); $l.Add('METHOD:PUBLISH')
    $l.Add('BEGIN:VEVENT')
    $l.Add("UID:$(ConvertTo-HUIcsText $Uid)")
    $l.Add("DTSTAMP:$(& $utc $Now)")
    $l.Add("SEQUENCE:$Sequence")
    if ($AllDay) {
        $last = $(if ($End) { ([datetime]$End).Date } else { $Start.Date })
        if ($last -lt $Start.Date) { throw 'Das Ende liegt vor dem Beginn.' }
        $l.Add("DTSTART;VALUE=DATE:$(& $day $Start.Date)"); $l.Add("DTEND;VALUE=DATE:$(& $day $last.AddDays(1))")
    } else {
        $e = $(if ($End) { ([datetime]$End) } else { $Start.AddHours(1) })
        if ($e -le $Start) { throw 'Das Ende liegt nicht nach dem Beginn.' }
        $l.Add("DTSTART:$(& $utc $Start)"); $l.Add("DTEND:$(& $utc $e)")
    }
    $l.Add("SUMMARY:$(ConvertTo-HUIcsText $Subject)")
    if ($Location) { $l.Add("LOCATION:$(ConvertTo-HUIcsText $Location)") }
    if ($Description) { $l.Add("DESCRIPTION:$(ConvertTo-HUIcsText $Description)") }
    if ($AlarmMinutes -ge 0) {
        $l.Add('BEGIN:VALARM'); $l.Add('ACTION:DISPLAY'); $l.Add("DESCRIPTION:$(ConvertTo-HUIcsText $Subject)"); $l.Add("TRIGGER:-PT$($AlarmMinutes)M"); $l.Add('END:VALARM')
    }
    $l.Add('END:VEVENT'); $l.Add('END:VCALENDAR')
    return ((@($l | ForEach-Object { Format-HUIcsLine $_ }) -join "`r`n") + "`r`n")
}

<#
.SYNOPSIS
    .ics-Datei schreiben (UTF-8 ohne BOM, CRLF). Parameter wie New-HUIcsContent. Rueckgabe: Pfad.
#>
function New-HUIcsFile {
    param(
        [Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Uid, [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][datetime]$Start, [Nullable[datetime]]$End = $null, [switch]$AllDay, [string]$Location = '',
        [string]$Description = '', [int]$AlarmMinutes = -1, [int]$Sequence = 0
    )
    $c = New-HUIcsContent -Uid $Uid -Subject $Subject -Start $Start -End $End -AllDay:$AllDay -Location $Location -Description $Description -AlarmMinutes $AlarmMinutes -Sequence $Sequence
    $full = [IO.Path]::GetFullPath($Path)
    $dir = Split-Path -Parent $full
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
    [IO.File]::WriteAllText($full, $c, (New-Object System.Text.UTF8Encoding $false))
    return $full
}

# ---------------------------------------------------------------- Word ----------------------------------------------------------------

# Text fuer Word: Zeilenwechsel werden manuelle Zeilenumbrueche (Absatzformat bleibt), Steuerzeichen fallen weg
function ConvertTo-HUWordText([string]$Text) {
    $t = "$Text" -replace "`r`n", "`n" -replace "`r", "`n"
    $t = $t -replace "`n", [string][char]11
    return ($t -replace '[\x00-\x08\x0C\x0E-\x1F]', '')
}

# Alle Fundstellen eines Platzhalters in einem Textbereich ersetzen (Fund wird direkt ueberschrieben, keine Laengengrenze)
function Set-HUWordPlaceholder($Story, [string]$Placeholder, [string]$Value) {
    $n = 0; $r = $null; $f = $null
    try {
        $r = $Story.Duplicate
        $f = $r.Find
        $f.ClearFormatting()
        $f.Text = $Placeholder; $f.Forward = $true; $f.Wrap = 0; $f.MatchCase = $true; $f.MatchWildcards = $false; $f.MatchWholeWord = $false
        while ($n -lt 1000 -and $f.Execute()) {
            $r.Text = $Value
            $r.Collapse(0)   # wdCollapseEnd
            $n++
        }
    } finally { Remove-HUComObject @($f, $r) }
    return $n
}

# Erste Fundstelle eines Platzhalters durch eine Tabelle ersetzen. $Rows: Zeilen als Arrays, die erste Zeile ist die Kopfzeile.
function Set-HUWordTable($Document, [string]$Placeholder, [object[]]$Rows) {
    $rows = @($Rows | Where-Object { $null -ne $_ })
    if ($rows.Count -eq 0) { return (Set-HUWordPlaceholder $Document.Content $Placeholder '') -gt 0 }
    $cols = ($rows | ForEach-Object { @($_).Count } | Measure-Object -Maximum).Maximum
    $r = $null; $f = $null; $tables = $null; $t = $null
    try {
        $r = $Document.Content
        $f = $r.Find
        $f.ClearFormatting()
        $f.Text = $Placeholder; $f.Forward = $true; $f.Wrap = 0; $f.MatchCase = $true; $f.MatchWildcards = $false
        if (-not $f.Execute()) { return $false }
        $r.Text = ''
        $tables = $Document.Tables
        $t = $tables.Add($r, $rows.Count, [int]$cols)
        $t.Borders.Enable = $true
        for ($i = 0; $i -lt $rows.Count; $i++) {
            $cells = @($rows[$i])
            for ($j = 0; $j -lt $cells.Count; $j++) {
                $c = $t.Cell($i + 1, $j + 1)
                try { $c.Range.Text = (ConvertTo-HUWordText "$($cells[$j])") } finally { Remove-HUComObject $c }
            }
        }
        $hdr = $t.Rows.Item(1)
        try { $hdr.Range.Font.Bold = $true; $hdr.HeadingFormat = $true } finally { Remove-HUComObject $hdr }
        return $true
    } finally { Remove-HUComObject @($t, $tables, $f, $r) }
}

<#
.SYNOPSIS
    Word-Vorlage (.docx) fuellen und als PDF (und/oder .docx) speichern. Die Vorlage selbst wird nie veraendert:
    Word oeffnet eine Kopie im Temp-Ordner. Word laeuft unsichtbar als eigene Instanz und wird immer beendet.
    -Values: @{ '{Mieter}' = 'Name'; ... } (in Text, Kopf- und Fusszeilen). -Tables: @{ '{Positionen}' = @(@('Pos','Preis'), @('Saal','270,00')) }.
    Rueckgabe: Pdf, Docx (Pfade oder leer), Missing (Platzhalter, die nicht vorkamen).
    -TimeoutSeconds: antwortet Word so lange nicht (z. B. unsichtbarer Dialog), wird genau diese Word-Instanz beendet und
    die Funktion wirft. Laeuft die eigene Instanz nach dem Beenden noch, wird sie ebenfalls beendet. Fremdes Word bleibt unberuehrt.
    Am besten in einem Hintergrund-Auftrag des Kerns aufrufen, damit das Fenster nicht stehen bleibt.
    Nur .docx und .dotx (ohne Makros). Makros bleiben auch sonst aus: AutomationSecurity = 3
    (msoAutomationSecurityForceDisable) vor dem Oeffnen; per Automation gestartetes Word hat sonst
    msoAutomationSecurityLow = alle Makros an (Microsoft Learn, MsoAutomationSecurity).
#>
function Invoke-HUWordTemplate {
    param(
        [Parameter(Mandatory)][string]$Template, [string]$PdfPath = '', [string]$DocxPath = '',
        [hashtable]$Values = @{}, [hashtable]$Tables = @{}, [ValidateRange(10, 3600)][int]$TimeoutSeconds = 120
    )
    Assert-HUWordTemplateExtension $Template
    Assert-HUOfficeLive 'Word'
    if (-not (Test-Path -LiteralPath $Template)) { throw "Vorlage nicht gefunden: $Template" }
    if (-not $PdfPath -and -not $DocxPath) { throw 'Kein Ziel angegeben (-PdfPath oder -DocxPath).' }
    if (-not (Test-HUWord)) { throw 'Word ist nicht installiert.' }
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtWord-' + [guid]::NewGuid().ToString('N') + [IO.Path]::GetExtension($Template))
    Copy-Item -LiteralPath $Template -Destination $tmp -Force
    (Get-Item -LiteralPath $tmp).IsReadOnly = $false
    $word = $null; $docs = $null; $doc = $null; $wordPid = 0; $guard = $null
    $missing = New-Object System.Collections.Generic.List[string]
    try {
        $before = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0   # wdAlertsNone
        $word.AutomationSecurity = 3   # msoAutomationSecurityForceDisable: keine Makros beim Oeffnen per Programm (KERN-B3)
        $docs = $word.Documents
        $doc = $docs.Open($tmp, $false, $false, $false)   # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles
        $wordPid = Get-HUWordProcessId -Document $doc -Before $before
        if ($wordPid) { $guard = Start-HUProcessGuard -ProcessId $wordPid -Seconds $TimeoutSeconds }
        foreach ($k in @($Tables.Keys)) { if (-not (Set-HUWordTable $doc "$k" @($Tables[$k]))) { $missing.Add("$k") } }
        foreach ($k in @($Values.Keys)) {
            $val = ConvertTo-HUWordText "$($Values[$k])"
            $hits = 0
            $stories = $doc.StoryRanges
            try {
                foreach ($st in $stories) {
                    $cur = $st
                    while ($null -ne $cur) {
                        $hits += (Set-HUWordPlaceholder $cur "$k" $val)
                        $next = $null
                        try { $next = $cur.NextStoryRange } catch { $next = $null }
                        Remove-HUComObject $cur
                        $cur = $next
                    }
                }
            } finally { Remove-HUComObject $stories }
            if ($hits -eq 0) { $missing.Add("$k") }
        }
        $pdf = ''; $docx = ''
        if ($PdfPath) {
            $pdf = [IO.Path]::GetFullPath($PdfPath)
            $d = Split-Path -Parent $pdf; if ($d -and -not (Test-Path -LiteralPath $d)) { [void](New-Item -ItemType Directory -Path $d -Force) }
            $doc.ExportAsFixedFormat($pdf, 17)   # wdExportFormatPDF
        }
        if ($DocxPath) {
            $docx = [IO.Path]::GetFullPath($DocxPath)
            $d = Split-Path -Parent $docx; if ($d -and -not (Test-Path -LiteralPath $d)) { [void](New-Item -ItemType Directory -Path $d -Force) }
            $doc.SaveAs2($docx, 16)   # wdFormatDocumentDefault
        }
        return [pscustomobject]@{ Pdf = $pdf; Docx = $docx; Missing = @($missing) }
    } catch {
        if ($guard -and $guard.Fired()) { throw "Word hat $TimeoutSeconds Sekunden nicht geantwortet und wurde beendet. Vorlage pruefen oder Word einmal von Hand oeffnen (Dialog?)." }
        throw
    } finally {
        if ($doc) { try { $doc.Close(0) } catch { } }       # wdDoNotSaveChanges
        # Quit ohne Argument: unter PowerShell 5.1 verlangt Quit(0) ein [ref] und wirft sonst (Word liefe weiter).
        # Das Dokument ist schon ohne Speichern geschlossen, darum gibt es nichts zu fragen.
        if ($word) { try { $word.Quit() } catch { } }
        Remove-HUComObject @($doc, $docs, $word)
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        if ($guard) { $guard.Stop() }
        if ($wordPid) { Stop-HUOwnProcess -ProcessId $wordPid -GraceSeconds 15 }
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    }
}

# Word-Vorlage nur als .docx oder .dotx (Formate ohne Makros), sonst eine klare Meldung (KERN-B3)
function Assert-HUWordTemplateExtension([string]$Template) {
    $ext = [IO.Path]::GetExtension("$Template").ToLowerInvariant()
    if (@('.docx', '.dotx') -notcontains $ext) {
        throw "Word-Vorlage muss eine .docx- oder .dotx-Datei sein (ohne Makros), nicht '$ext': $Template"
    }
}

# Prozess-Id der eigenen Word-Instanz: ueber das Dokumentfenster (Window.Hwnd, ab Word 2013), sonst der einzige neue WINWORD-Prozess
function Get-HUWordProcessId($Document, [int[]]$Before = @()) {
    $win = $null
    try {
        $win = $Document.ActiveWindow
        $h = [IntPtr][long]$win.Hwnd
        if ($h -ne [IntPtr]::Zero) {
            if (-not ('HUOffice.Win32' -as [type])) {
                Add-Type -Namespace HUOffice -Name Win32 -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);'
            }
            $p = [uint32]0
            [void][HUOffice.Win32]::GetWindowThreadProcessId($h, [ref]$p)
            if ($p -gt 0) { return [int]$p }
        }
    } catch { } finally { Remove-HUComObject $win }
    $new = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Where-Object { $Before -notcontains $_.Id })
    if ($new.Count -eq 1) { return [int]$new[0].Id }
    return 0   # nicht eindeutig: lieber keinen Prozess anfassen
}

# Aufpasser in eigenem Runspace: beendet den Prozess nach -Seconds, wenn Stop() nicht vorher kommt
function Start-HUProcessGuard([int]$ProcessId, [int]$Seconds) {
    $state = [hashtable]::Synchronized(@{ Fired = $false })
    $ps = [powershell]::Create()
    [void]$ps.AddScript({
            param($id, $sec, $st)
            Start-Sleep -Seconds $sec
            $st.Fired = $true
            Stop-Process -Id $id -Force -ErrorAction SilentlyContinue
        }).AddArgument($ProcessId).AddArgument($Seconds).AddArgument($state)
    $handle = $ps.BeginInvoke()
    $g = [pscustomobject]@{ PS = $ps; Handle = $handle; State = $state }
    $g | Add-Member ScriptMethod Fired { [bool]$this.State.Fired }
    $g | Add-Member ScriptMethod Stop { try { $this.PS.Stop() } catch { }; try { $this.PS.Dispose() } catch { } }
    return $g
}

# Eigenen Prozess im Hintergrund beenden, falls er nach -GraceSeconds noch laeuft (z. B. weil Quit nicht durchkam).
# Kehrt sofort zurueck; Word braucht nach Quit einige Sekunden. Der Runspace endet, sobald der Prozess weg ist.
function Stop-HUOwnProcess([int]$ProcessId, [int]$GraceSeconds = 15) {
    $ps = [powershell]::Create()
    [void]$ps.AddScript({
            param($id, $sec)
            for ($i = 0; $i -lt ($sec * 4); $i++) {
                if (-not (Get-Process -Id $id -ErrorAction SilentlyContinue)) { return }
                Start-Sleep -Milliseconds 250
            }
            Stop-Process -Id $id -Force -ErrorAction SilentlyContinue
        }).AddArgument($ProcessId).AddArgument($GraceSeconds)
    [void]$ps.BeginInvoke()
}

Export-ModuleMember -Function Test-HUOutlookClassic, Test-HUWord, Remove-HUComObject, Test-HUOfficeKey, Test-HUOfficeTestMode, Get-HUOutlookAccounts, Get-HUOutlookContacts,
    Open-HUOutlook, Close-HUOutlook, Invoke-HUOutlook, ConvertTo-HUOutlookFolderSpec, Get-HUOutlookFolder, Select-HUOutlookFolder,
    Find-HUOutlookItem, Get-HUOutlookItemKey, Set-HUOutlookAppointment, Set-HUOutlookTask, Complete-HUOutlookItem, Remove-HUOutlookItem,
    New-HUMailtoUri, Show-HUMail, Get-HUOutlookAttachments, Get-HUOutlookMailboxInbox, ConvertTo-HUIcsText,
    Read-HUFileGroupDescriptor, Get-HUDropDescriptor, Get-HUDropContent, Get-HUDropNames, Test-HUDropHasFiles, Save-HUDroppedFiles, Format-HUIcsLine, New-HUIcsContent, New-HUIcsFile,
    ConvertTo-HUWordText, Invoke-HUWordTemplate
