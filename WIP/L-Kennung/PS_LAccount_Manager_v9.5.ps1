# FileName: PS_LAccount_Manager.ps1
# Version:  9.5
# Beschreibung: AD-Sync Tool — Sequenziell, StreamWriter-Export, Vollbild-Debug
# Author:   PS-Coding
#
# ARCHITEKTUR-ENTSCHEIDUNG v9.5:
#   Runspaces entfernt -> sequenzielle Verarbeitung.
#   Grund: AD-Cache wird einmalig geladen (der teure Teil), die eigentliche
#   Pro-Eintrag-Berechnung ist CPU-leicht -> Runspace-Overhead > Gewinn.
#   Vorteil: Kein PSDataCollection-Bug, kein Komma-Parse-Bug, kein RAM-Blowup.
#
# EXPORT-FIX:
#   Out-File mit vorgebautem String-Array (2091 x 200 Spalten) -> OOM / Hang.
#   Fix: System.IO.StreamWriter schreibt Zeile fuer Zeile direkt auf Disk.
#   RAM-Verbrauch: O(1) statt O(n*m).

[CmdletBinding()]
param(
    [string] $CsvPath,
    [string] $RequiredCsvPath,
    [switch] $SearchGlobal,
    [int]    $TestCount  = 0,
    [switch] $DebugMode
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

$Version   = "9.5"
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile   = Join-Path $ScriptDir "AD_Sync_$Timestamp.log"
$SW        = [System.Diagnostics.Stopwatch]::StartNew()

$AD_PROPS = @("DisplayName","Description","GivenName","Surname","l",
              "physicalDeliveryOfficeName","department","info","MemberOf")

# ══════════════════════════════════════════════════════════════════
#  UI & LOGGING
# ══════════════════════════════════════════════════════════════════
function Show-Banner {
    Clear-Host
    $ts = Get-Date -Format "dd.MM.yyyy  HH:mm:ss"
    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════════════════════╗" -ForegroundColor DarkCyan
    Write-Host "  ║  🛡️  AD COMPLIANCE & SYNC MANAGER  ·  v$Version              ║" -ForegroundColor Cyan
    Write-Host "  ║  📅  $ts                            ║" -ForegroundColor DarkCyan
    Write-Host "  ║  ⚙️  Modus: Sequenziell · StreamWriter-Export             ║" -ForegroundColor DarkCyan
    Write-Host "  ╚══════════════════════════════════════════════════════════╝" -ForegroundColor DarkCyan
    Write-Host ""
}

function Write-Log {
    param(
        [string] $Msg,
        [ValidateSet("INFO","OK","WARN","ERR","DBG","STEP","PERF")] $L = "INFO",
        [switch] $NoFile
    )
    if ($L -eq "DBG" -and -not $DebugMode) { return }

    $ts    = Get-Date -Format "HH:mm:ss.fff"
    $elapsed = "  +{0,7:F2}s" -f $SW.Elapsed.TotalSeconds

    $icon  = switch ($L) {
        "OK"   { "✅" } "WARN" { "⚠️ " } "ERR"  { "❌" }
        "DBG"  { "🔍" } "STEP" { "▶️ " } "PERF" { "⏱️ " }
        default{ "ℹ️ " }
    }
    $color = switch ($L) {
        "OK"   { "Green"   } "WARN" { "Yellow"  } "ERR"  { "Red"     }
        "DBG"  { "Magenta" } "STEP" { "Cyan"    } "PERF" { "DarkYellow" }
        default{ "Gray"    }
    }

    $line = "[$ts]$elapsed  $icon  $Msg"
    Write-Host $line -ForegroundColor $color
    if (-not $NoFile) { $line | Out-File $LogFile -Append }
}

function Write-Section {
    param([string]$Title)
    $bar = "─" * 60
    Write-Host ""
    Write-Host "  ┌$bar┐" -ForegroundColor DarkGray
    Write-Host "  │  $Title" -ForegroundColor White
    Write-Host "  └$bar┘" -ForegroundColor DarkGray
    Write-Log "=== $Title ===" "STEP"
}

function Write-Progress-Custom {
    param([int]$Done, [int]$Total, [string]$Label = "")
    $pct  = if ($Total -gt 0) { [math]::Round(($Done / $Total) * 100) } else { 0 }
    $barW = 40
    $fill = [math]::Round($barW * $pct / 100)
    $bar  = ("█" * $fill) + ("░" * ($barW - $fill))

    # ETA berechnen
    $elapsed = $SW.Elapsed.TotalSeconds
    $eta     = ""
    if ($Done -gt 0 -and $Done -lt $Total) {
        $secPerItem = $elapsed / $Done
        $remaining  = [math]::Round($secPerItem * ($Total - $Done))
        $eta        = "  ETA ~${remaining}s"
    }

    Write-Progress -Activity "⚙️  $Label" `
        -Status "  $bar  $Done / $Total  ($pct%)$eta" `
        -PercentComplete $pct
}

# ══════════════════════════════════════════════════════════════════
#  HILFSFUNKTIONEN
# ══════════════════════════════════════════════════════════════════
function Get-Str { param($v)
    if ($null -eq $v) { return "" }
    return $v.ToString().Trim()
}

function Test-Diff { param($a, $b)
    return ((Get-Str $a) -ine (Get-Str $b))
}

function Get-GroupName { param([string]$DN)
    return (($DN -split ',')[0] -replace 'CN=', '')
}

function Get-Verfahren { param($Row)
    if ($null -eq $Row) { return "[alle_Verfahren]" }
    $cols = @("Viva","Findus","MobiApps","AccVisio",
              "Verfahren5","Verfahren6","Verfahren7","Verfahren8","Verfahren9","Verfahren10")
    $hit  = foreach ($c in $cols) {
        if ($Row.$c -and $Row.$c.ToString().Trim().ToLower() -eq "x") { $c }
    }
    if ($hit) { return ($hit -join " - ") }
    return "[alle_Verfahren]"
}

# StreamWriter: Zeile fuer Zeile -> kein RAM-Aufbau
function Write-CsvLine {
    param([System.IO.StreamWriter]$Writer, [array]$Values)
    $escaped = foreach ($v in $Values) {
        $s = if ($null -ne $v) { $v.ToString() } else { "" }
        if ($s -match '[;\r\n"]') { '"' + $s.Replace('"','""') + '"' } else { $s }
    }
    $Writer.WriteLine($escaped -join ';')
}

# ══════════════════════════════════════════════════════════════════
#  LOGIC-FUNCTIONS
# ══════════════════════════════════════════════════════════════════

function Get-RequiredList {
    param([string]$Path)
    Write-Section "📋  BEDARFSLISTE laden"
    $list = New-Object 'System.Collections.Generic.HashSet[string]'
    if ($Path -and (Test-Path $Path)) {
        $lines = Get-Content $Path -Encoding UTF8
        Write-Log "Rohdaten: $($lines.Count) Zeilen" "DBG"
        foreach ($line in $lines) {
            if ($line -match "(L\d{6,8})") {
                [void]$list.Add($Matches[1].ToUpper())
            }
        }
        Write-Log "Bedarfsliste: $($list.Count) L-Kennungen eingelesen" "OK"
    } else {
        Write-Log "Keine Bedarfsliste angegeben oder Datei fehlt -> übersprungen" "WARN"
    }
    return $list
}

function Get-MasterCsv {
    param([string]$Path, [System.Diagnostics.Stopwatch]$SW)
    Write-Section "📊  MASTER-CSV laden"
    $t0 = $SW.Elapsed.TotalSeconds
    $dict = @{}
    $rawCsv = Import-Csv -Path $Path -Delimiter ';' -Encoding Default
    Write-Log "CSV eingelesen: $($rawCsv.Count) Rohdatensaetze" "DBG"
    foreach ($row in $rawCsv) {
        if ($row."L-Kennung") {
            $k = $row."L-Kennung".ToString().Trim().ToUpper()
            if ($k) { $dict[$k] = $row }
        }
    }
    Write-Log "Master-CSV: $($dict.Count) gueltige L-Kennungen  [+$([math]::Round($SW.Elapsed.TotalSeconds-$t0,2))s]" "OK"
    return $dict
}

function Get-ADDiscovery {
    param([switch]$SearchGlobal, [System.Diagnostics.Stopwatch]$SW)
    Write-Section "🔎  AD-DISCOVERY"
    $t0           = $SW.Elapsed.TotalSeconds
    $ADCache      = @{}
    $UniqueGroups = New-Object 'System.Collections.Generic.HashSet[string]'

    $TargetOUs = Get-ADOrganizationalUnit `
        -Filter "Name -eq '81' -or Name -eq '82'" `
        -ErrorAction SilentlyContinue

    Write-Log "Gefundene Ziel-OUs: $(@($TargetOUs).Count)" "DBG"

    foreach ($ou in $TargetOUs) {
        Write-Log "Scanne OU: $($ou.DistinguishedName)" "DBG"
        $t1    = $SW.Elapsed.TotalSeconds
        $users = Get-ADUser -Filter * `
                     -SearchBase $ou.DistinguishedName `
                     -Properties $AD_PROPS
        $cnt   = @($users).Count
        Write-Log "OU $($ou.Name): $cnt User geladen  [+$([math]::Round($SW.Elapsed.TotalSeconds-$t1,2))s]" "DBG"

        foreach ($u in $users) {
            $sam = $u.SamAccountName.ToUpper()
            $ADCache[$sam] = $u
            foreach ($gDN in $u.MemberOf) {
                $gn = Get-GroupName -DN $gDN
                [void]$UniqueGroups.Add($gn)
            }
        }
        Write-Log "OU $($ou.Name): verarbeitet → AD-Cache jetzt: $($ADCache.Count) Eintraege" "DBG"
    }

    if ($SearchGlobal) {
        Write-Log "Globaler AD-Scan (SamAccountName like 'L*') gestartet..." "STEP"
        $t1   = $SW.Elapsed.TotalSeconds
        $gAll = Get-ADUser -Filter "SamAccountName -like 'L*'" -Properties $AD_PROPS
        Write-Log "Globaler Scan: $(@($gAll).Count) Treffer  [+$([math]::Round($SW.Elapsed.TotalSeconds-$t1,2))s]" "DBG"
        $newCount = 0
        foreach ($gu in $gAll) {
            $sam = $gu.SamAccountName.ToUpper()
            if (-not $ADCache.ContainsKey($sam)) {
                $ADCache[$sam] = $gu
                $newCount++
                foreach ($gDN in $gu.MemberOf) {
                    $gn = Get-GroupName -DN $gDN
                    [void]$UniqueGroups.Add($gn)
                }
            }
        }
        Write-Log "Globaler Scan: $newCount neue Eintraege hinzugefuegt" "OK"
    }

    $SortedGroups = @($UniqueGroups | Sort-Object)
    Write-Log "AD-Discovery abgeschlossen: $($ADCache.Count) AD-Objekte  |  $($SortedGroups.Count) unique Gruppen  [+$([math]::Round($SW.Elapsed.TotalSeconds-$t0,2))s]" "PERF"

    return @{
        Cache = $ADCache
        Groups = $SortedGroups
    }
}

function Invoke-Processing {
    param(
        [array]$ProcessList,
        [hashtable]$MasterCsvData,
        [hashtable]$ADCache,
        [System.Collections.Generic.HashSet[string]]$RequiredList,
        [array]$SortedGroups,
        [System.Diagnostics.Stopwatch]$SW,
        [switch]$DebugMode
    )
    Write-Section "⚙️   VERARBEITUNG  ($($ProcessList.Count) Eintraege · sequenziell)"
    $t0      = $SW.Elapsed.TotalSeconds
    $Results = New-Object 'System.Collections.Generic.List[PSObject]'
    $Total   = $ProcessList.Count
    $idx     = 0
    $warnedLong = $false

    foreach ($LID in $ProcessList) {
        $idx++
        $tItem = $SW.Elapsed.TotalSeconds

        $CsvRow = $MasterCsvData[$LID]
        $ADObj  = $ADCache[$LID]

        # -- AD-Felder --
        $andereOU        = ""
        $statusGeloescht = ""
        $userGroups      = @()

        if ($ADObj) {
            $dn = $ADObj.DistinguishedName
            if ($dn -notmatch "OU=81" -and $dn -notmatch "OU=82") {
                $andereOU = $dn -replace '^CN=.*?,', ''
            }
            if ($ADObj.MemberOf) {
                $userGroups = @($ADObj.MemberOf | ForEach-Object {
                    Get-GroupName -DN $_
                })
            }
        } else {
            $statusGeloescht = "XXX"
        }

        # -- Berechnete Werte --
        $isRequired = if ($RequiredList.Contains($LID)) { "ZZZ" } else { "" }
        $isLafp     = ($LID -like "L110*") -or ($LID -like "L114*")
        $lafpStr    = if ($isLafp) { "LLL" } else { "" }
        $ouNum      = if ($isLafp) { "26" } else { "[OU]" }

        $targetOrt = ""
        if ($CsvRow -and $CsvRow.Standort -and ($CsvRow.Standort.ToString().Trim() -ne "")) {
            $targetOrt = $CsvRow.Standort.ToString().Trim()
        } elseif ($ADObj -and $ADObj.l) {
            $targetOrt = Get-Str $ADObj.l
        }

        $targetNachname = $targetOrt
        if ($isLafp -and ($targetOrt -notmatch "^LAFP\s-\s")) {
            $targetNachname = "LAFP - $targetOrt"
        }

        $buroPlatz = "[Platznummer]"
        if ($CsvRow -and $CsvRow.Raum -and ($CsvRow.Raum.ToString().Trim() -ne "")) {
            $buroPlatz = $CsvRow.Raum.ToString().Trim()
        }

        $fortbildung = if ($CsvRow -and $CsvRow.Fortbildungsbereich) { Get-Str $CsvRow.Fortbildungsbereich } else { "" }
        $anmerkung   = if ($CsvRow -and $CsvRow.Anmerkungen) { Get-Str $CsvRow.Anmerkungen } else { "" }
        $aenderInfo  = if ($fortbildung -ne "" -and $anmerkung -ne "") { "$fortbildung - $anmerkung" } else { "$fortbildung$anmerkung" }
        $verfStr     = Get-Verfahren -Row $CsvRow

        $adVorname  = if ($ADObj) { Get-Str $ADObj.GivenName                  } else { "" }
        $adNachname = if ($ADObj) { Get-Str $ADObj.Surname                    } else { "" }
        $adDisplay  = if ($ADObj) { Get-Str $ADObj.DisplayName                } else { "" }
        $adOrt      = if ($ADObj) { Get-Str $ADObj.l                          } else { "" }
        $adBuero    = if ($ADObj) { Get-Str $ADObj.physicalDeliveryOfficeName } else { "" }
        $adDez      = if ($ADObj) { Get-Str $ADObj.department                 } else { "" }
        $adDesc     = if ($ADObj) { Get-Str $ADObj.Description                } else { "" }
        $adInfo     = if ($ADObj) { Get-Str $ADObj.info                       } else { "" }

        $origStandort = if ($CsvRow) { Get-Str $CsvRow.Standort  } else { "" }
        $origRaum     = if ($CsvRow) { Get-Str $CsvRow.Raum      } else { "" }
        $origAbt      = if ($CsvRow) { Get-Str $CsvRow.Abteilung } else { "" }

        $aendernDez  = if ($CsvRow -and $ADObj -and $adDez -ne "") { "$origAbt - $adDez" } else { $origAbt }
        $displayName = "$LID $targetNachname"
        $aendernDesc = "$ouNum - $verfStr - $buroPlatz - [Verantwortlicher] - [TEL] | $aenderInfo"

        # -- Objekt aufbauen --
        $Props = [ordered]@{
            "L-Kennung"                    = $LID
            "andere_OU"                    = $andereOU
            "GELOESCHT"                    = $statusGeloescht
            "Benoetigt"                    = $isRequired
            "Geaendert"                    = ""
            "LAFP_LZPD_LKA"               = $lafpStr
            "NICHT_KONFORM"               = ""
            "AD_Vorname"                  = $adVorname
            "AENDERN_Vorname"             = $LID
            "AD_Nachname"                 = $adNachname
            "AENDERN_Nachname"            = $targetNachname
            "AD_DisplayName"              = $adDisplay
            "AENDERN_DisplayName"         = $displayName
            "ORIGINAL_Standort"           = $origStandort
            "ORIGINAL_Raum_Schulungskreis"= $origRaum
            "AD_Ort"                      = $adOrt
            "AENDERN_Ort"                 = $targetOrt
            "AD_Buero"                    = $adBuero
            "AENDERN_Buero"               = $buroPlatz
            "ORIGINAL_Abteilung"          = $origAbt
            "AD_Dez"                      = $adDez
            "AENDERN_Dez"                 = $aendernDez
            "AD_Description"              = $adDesc
            "AENDERN_Description"         = $aendernDesc
            "AENDERN_OU"                  = $ouNum
            "AENDERN_Info"                = $aenderInfo
            "LOESCHEN"                    = ""
        }

        # -- Konformitaet --
        if ($ADObj) {
            $codes = New-Object 'System.Collections.Generic.List[string]'
            if (Test-Diff $adVorname  $LID)           { $codes.Add("VN")   }
            if (Test-Diff $adNachname $targetNachname) { $codes.Add("NN")   }
            if (Test-Diff $adDisplay  $displayName)   { $codes.Add("DN")   }
            if (Test-Diff $adDesc     $aendernDesc)   { $codes.Add("DESC") }
            if (Test-Diff $adOrt      $targetOrt)     { $codes.Add("ORT")  }
            if (Test-Diff $adBuero    $buroPlatz)     { $codes.Add("GEB")  }
            if (Test-Diff $adDez      $aendernDez)    { $codes.Add("DEZ")  }
            if (Test-Diff $adInfo     $aenderInfo)    { $codes.Add("INFO") }
            if ($codes.Count -gt 0) {
                $cs = $codes -join ","
                $Props."Geaendert"    = $cs
                $Props."NICHT_KONFORM" = $cs
            }
        }

        # -- Gruppen --
        foreach ($gn in $SortedGroups) {
            $Props["GRP_$gn"] = if ($userGroups -contains $gn) { "X" } else { "" }
        }

        [void]$Results.Add([PSCustomObject]$Props)

        # -- Live-Output --
        $itemMs = [math]::Round(($SW.Elapsed.TotalSeconds - $tItem) * 1000, 1)
        if ($DebugMode) {
            $src = if ($ADObj -and $CsvRow) { "AD+CSV" } elseif ($ADObj) { "AD   " } else { "CSV  " }
            Write-Host ("  [{0,4}/{1}]  {2,-14}  src:{3}  nk:{4,-20}  {5,5}ms" -f `
                $idx, $Total, $LID, $src, $Props."NICHT_KONFORM", $itemMs) `
                -ForegroundColor DarkGray
        }

        # Progress alle 50 Eintraege oder am Ende
        if (($idx % 50 -eq 0) -or ($idx -eq $Total)) {
            $elapsed   = $SW.Elapsed.TotalSeconds
            $secPer    = $elapsed / $idx
            $remaining = [math]::Round($secPer * ($Total - $idx))
            $pct       = [math]::Round(($idx / $Total) * 100)
            $barFill   = [math]::Round(40 * $pct / 100)
            $bar       = ("█" * $barFill) + ("░" * (40 - $barFill))

            $etaStr    = if ($idx -lt $Total) { "  ⏳ ETA ~${remaining}s" } else { "  ✅ fertig" }
            Write-Host "`r  [$bar]  $idx/$Total  ($pct%)$etaStr    " `
                -ForegroundColor Cyan -NoNewline

            Write-Progress -Activity "⚙️  Verarbeitung" `
                -Status "$idx / $Total  ($pct%)  ~${remaining}s verbleibend" `
                -PercentComplete $pct
        }

        # Hang-Warnung: wenn ein einzelner Eintrag > 5s braucht
        if ($itemMs -gt 5000 -and -not $warnedLong) {
            $warnedLong = $true
            Write-Log "⚠️  HANG-WARNUNG: Eintrag $LID brauchte ${itemMs}ms!" "WARN"
        }
    }

    Write-Host ""  # Newline nach \r Progress
    Write-Progress -Activity "⚙️  Verarbeitung" -Completed
    $procTime = [math]::Round($SW.Elapsed.TotalSeconds - $t0, 2)
    Write-Log "Verarbeitung abgeschlossen: $($Results.Count) Datensaetze in ${procTime}s  (~$([math]::Round($procTime/$Total*1000,1))ms/Eintrag)" "PERF"

    return $Results
}

function Export-Results {
    param(
        [System.Collections.Generic.List[PSObject]]$Results,
        [array]$SortedGroups,
        [System.Diagnostics.Stopwatch]$SW,
        [string]$ScriptDir,
        [string]$Version
    )
    Write-Section "💾  EXPORT  ($($Results.Count) Zeilen · $($SortedGroups.Count) Gruppenspalten)"

    if ($Results.Count -eq 0) {
        Write-Log "Keine Ergebnisse — Export uebersprungen" "ERR"; return
    }

    # --- Tabelle 1: Gruppen-Uebersicht ---
    $t0    = $SW.Elapsed.TotalSeconds
    $Path1 = Join-Path $ScriptDir "L-Kennungen_Full_Analysis_v$Version.csv"
    Write-Log "Schreibe Tab1: $Path1" "DBG"

    $cols1   = @("L-Kennung","andere_OU","GELOESCHT","Benoetigt","Geaendert",
                 "LOESCHEN","AD_Vorname","AD_Nachname")
    $grpCols = @($SortedGroups | ForEach-Object { "GRP_$_" })
    $allCols1 = $cols1 + $grpCols

    $sw1 = New-Object System.IO.StreamWriter($Path1, $false, [System.Text.Encoding]::UTF8)
    try {
        Write-CsvLine -Writer $sw1 -Values $allCols1
        $wIdx = 0
        foreach ($r in $Results) {
            $wIdx++
            $vals = foreach ($c in $allCols1) { $r.$c }
            Write-CsvLine -Writer $sw1 -Values $vals
            if ($wIdx % 200 -eq 0) {
                Write-Host "`r  💾 Tab1: $wIdx/$($Results.Count) Zeilen...    " -NoNewline -ForegroundColor DarkGray
            }
        }
    } finally { $sw1.Close() }
    Write-Host ""
    Write-Log "Tab1 fertig: $Path1  [+$([math]::Round($SW.Elapsed.TotalSeconds-$t0,2))s]" "OK"

    # --- Tabelle 2: Alle Properties ---
    $t0    = $SW.Elapsed.TotalSeconds
    $Path2 = Join-Path $ScriptDir "L-Kennungen_Properties_v$Version.csv"
    Write-Log "Schreibe Tab2: $Path2" "DBG"

    $allCols2 = @($Results[0].psobject.Properties.Name | Where-Object { $_ -ne "LOESCHEN" })
    Write-Log "Tab2: $($allCols2.Count) Spalten × $($Results.Count) Zeilen = $($allCols2.Count * $Results.Count) Zellen" "DBG"

    $sw2 = New-Object System.IO.StreamWriter($Path2, $false, [System.Text.Encoding]::UTF8)
    try {
        Write-CsvLine -Writer $sw2 -Values $allCols2
        $wIdx = 0
        foreach ($r in $Results) {
            $wIdx++
            $vals = foreach ($c in $allCols2) { $r.$c }
            Write-CsvLine -Writer $sw2 -Values $vals
            if ($wIdx % 200 -eq 0) {
                Write-Host "`r  💾 Tab2: $wIdx/$($Results.Count) Zeilen...    " -NoNewline -ForegroundColor DarkGray
            }
        }
    } finally { $sw2.Close() }
    Write-Host ""
    Write-Log "Tab2 fertig: $Path2  [+$([math]::Round($SW.Elapsed.TotalSeconds-$t0,2))s]" "OK"

    return @{ Path1 = $Path1; Path2 = $Path2 }
}


# ══════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════
function Invoke-Main {
    Show-Banner
    Import-Module ActiveDirectory -Verbose:$false

    Write-Log "PowerShell v$($PSVersionTable.PSVersion)  |  PID $PID  |  Host: $env:COMPUTERNAME" "DBG"
    Write-Log "Logfile: $LogFile" "INFO"

    # ──────────────────────────────────────────────────────────────
    # INPUT
    # ──────────────────────────────────────────────────────────────
    if (-not $CsvPath) { $CsvPath = Read-Host "📂  Pfad zur Master-CSV" }
    $CsvPath = $CsvPath.Trim().Trim('"')
    if (-not (Test-Path $CsvPath)) {
        Write-Log "Master-CSV nicht gefunden: '$CsvPath'" "ERR"; return
    }
    Write-Log "Master-CSV: $CsvPath" "DBG"

    $RequiredList = Get-RequiredList -Path $RequiredCsvPath

    $MasterCsvData = Get-MasterCsv -Path $CsvPath -SW $SW

    $ADResult = Get-ADDiscovery -SearchGlobal:$SearchGlobal -SW $SW
    $ADCache = $ADResult.Cache
    $SortedGroups = $ADResult.Groups

    # Arbeitsliste aufbauen
    $AllSAMs     = @(@($ADCache.Keys) + @($MasterCsvData.Keys) | Select-Object -Unique | Sort-Object)
    Write-Log "Gesamt unique SAMs (AD + CSV): $($AllSAMs.Count)" "DBG"

    $ProcessList = if ($TestCount -gt 0) {
        Write-Log "⚠️  TESTMODUS: nur erste $TestCount Eintraege" "WARN"
        @($AllSAMs | Select-Object -First $TestCount)
    } else { $AllSAMs }

    Write-Log "Zu verarbeitende Eintraege: $($ProcessList.Count)" "OK"

    $Results = Invoke-Processing -ProcessList $ProcessList -MasterCsvData $MasterCsvData -ADCache $ADCache -RequiredList $RequiredList -SortedGroups $SortedGroups -SW $SW -DebugMode:$DebugMode

    $ExportPaths = Export-Results -Results $Results -SortedGroups $SortedGroups -SW $SW -ScriptDir $ScriptDir -Version $Version
    $Path1 = $ExportPaths.Path1
    $Path2 = $ExportPaths.Path2

    # ──────────────────────────────────────────────────────────────
    # ABSCHLUSS
    # ──────────────────────────────────────────────────────────────
    $SW.Stop()
    $totalMB = [math]::Round([System.GC]::GetTotalMemory($false) / 1MB, 1)

    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "  ║  ✅  ABGESCHLOSSEN                                        ║" -ForegroundColor Green
    Write-Host ("  ║  📦  {0,-52}║" -f "Datensaetze  : $($Results.Count)") -ForegroundColor Green
    Write-Host ("  ║  🗂️   {0,-51}║" -f "Gruppen      : $($SortedGroups.Count)") -ForegroundColor Green
    Write-Host ("  ║  ⏱️   {0,-52}║" -f "Gesamtdauer  : $($SW.Elapsed.ToString('mm\:ss\.ff')) min") -ForegroundColor Green
    Write-Host ("  ║  🧠  {0,-52}║" -f "RAM (managed): ${totalMB} MB") -ForegroundColor Green
    Write-Host "  ║                                                          ║" -ForegroundColor DarkGreen
    Write-Host ("  ║  📄  Tab1 : {0,-47}║" -f (Split-Path $Path1 -Leaf)) -ForegroundColor DarkGreen
    Write-Host ("  ║  📄  Tab2 : {0,-47}║" -f (Split-Path $Path2 -Leaf)) -ForegroundColor DarkGreen
    Write-Host ("  ║  📋  Log  : {0,-47}║" -f (Split-Path $LogFile -Leaf)) -ForegroundColor DarkGreen
    Write-Host "  ╚══════════════════════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host ""

    Write-Log "Fertig. Dauer: $($SW.Elapsed.ToString('mm\:ss\.ff'))  RAM: ${totalMB}MB  Zeilen: $($Results.Count)" "PERF"
}

Invoke-Main
