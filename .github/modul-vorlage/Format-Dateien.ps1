#Requires -Version 5.1
<#
.SYNOPSIS
    Bringt die eigenen Textdateien des Moduls auf UTF-8 mit BOM und CRLF (.yml und .py ohne BOM, mit LF).
.NOTES
    Zielmaschine: nb001. Wird vor jedem Commit und von Deploy-Dev.ps1 (gleiche Regel) gebraucht.
#>
$ErrorActionPreference = 'Stop'
$src = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$bom = New-Object System.Text.UTF8Encoding $true
$nobom = New-Object System.Text.UTF8Encoding $false
foreach ($f in @(Get-ChildItem -LiteralPath $src -Recurse -File -Force | Where-Object { $_.FullName -notmatch '\\\.git\\' })) {
    $ext = $f.Extension.ToLowerInvariant()
    $lf = $ext -in '.yml', '.py'
    if (-not $lf -and $ext -notin '.ps1', '.psm1', '.xaml', '.json', '.md', '.html', '.txt') { if ($f.Name -notin 'LICENSE', '.gitignore', '.gitattributes') { continue } }
    $t = [IO.File]::ReadAllText($f.FullName, [Text.Encoding]::UTF8)
    $n = $t -replace "`r`n", "`n"
    if (-not $lf) { $n = $n -replace "`n", "`r`n" }
    $b = [IO.File]::ReadAllBytes($f.FullName)
    $hasBom = $b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF
    $wantBom = -not $lf -and $f.Name -notin '.gitignore', '.gitattributes'
    if ($n -ne $t -or $hasBom -ne $wantBom) {
        [IO.File]::WriteAllText($f.FullName, $n, $(if ($wantBom) { $bom } else { $nobom }))
        Write-Host "angepasst: $($f.FullName.Substring($src.Length + 1))"
    }
}
