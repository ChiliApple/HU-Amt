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

<#
.SYNOPSIS
    Ordner im Outlook-Dialog waehlen (nur im Fenster-Thread). Rueckgabe: Ordnerangabe oder $null bei Abbruch.
    Passt der Ordnertyp nicht (z. B. Posteingang statt Kalender), wirft die Funktion mit einer verstaendlichen Meldung.
#>
function Select-HUOutlookFolder {
    param([ValidateSet('Calendar', 'Tasks')][string]$Type = 'Calendar', $Session = $null)
    return (Use-HUOutlookSession $Session {
            param($s)
            $f = $s.Namespace.PickFolder()
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
    Mail im Outlook anzeigen, nie senden. Ohne Outlook (oder -Mailweg mailto): mailto: ohne Anhang.
    Rueckgabe: Way (outlook | mailto | fehler), Message (Grund fuer den Rueckfall oder Fehler).
    Nur im Fenster-Thread aufrufen (Display).
#>
function Show-HUMail {
    param([string[]]$To = @(), [string[]]$Cc = @(), [string]$Subject = '', [string]$Body = '', [string[]]$Attachments = @(),
        [ValidateSet('auto', 'mailto')][string]$Mailweg = 'auto', [int]$MaxLength = 1800, [string]$TruncatedNote = '(Text gekuerzt.)')
    $toText = (@($To) | Where-Object { "$_".Trim() } | ForEach-Object { "$_".Trim() }) -join '; '
    $ccText = (@($Cc) | Where-Object { "$_".Trim() } | ForEach-Object { "$_".Trim() }) -join '; '
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
            $att = $m.Attachments
            foreach ($a in @($Attachments)) {
                if (-not "$a") { continue }
                if (-not (Test-Path -LiteralPath $a)) { throw "Anhang fehlt: $a" }
                $x = $att.Add($a); Remove-HUComObject $x
            }
            $m.Display($false)
            return [pscustomobject]@{ Way = 'outlook'; Message = '' }
        } catch { $errText = $_.Exception.Message }
        finally { Remove-HUComObject @($att, $m); Close-HUOutlook $s }
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
#>
function Invoke-HUWordTemplate {
    param(
        [Parameter(Mandatory)][string]$Template, [string]$PdfPath = '', [string]$DocxPath = '',
        [hashtable]$Values = @{}, [hashtable]$Tables = @{}, [ValidateRange(10, 3600)][int]$TimeoutSeconds = 120
    )
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

Export-ModuleMember -Function Test-HUOutlookClassic, Test-HUWord, Remove-HUComObject, Test-HUOfficeKey,
    Open-HUOutlook, Close-HUOutlook, Invoke-HUOutlook, ConvertTo-HUOutlookFolderSpec, Get-HUOutlookFolder, Select-HUOutlookFolder,
    Find-HUOutlookItem, Get-HUOutlookItemKey, Set-HUOutlookAppointment, Set-HUOutlookTask, Complete-HUOutlookItem, Remove-HUOutlookItem,
    New-HUMailtoUri, Show-HUMail, ConvertTo-HUIcsText, Format-HUIcsLine, New-HUIcsContent, New-HUIcsFile,
    ConvertTo-HUWordText, Invoke-HUWordTemplate
