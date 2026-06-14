# FileName: PS_LAccount_Manager.ps1
# Version:  9.6
# Author:   PS-Coding
#
# PARAMETER:
#   -CsvPath          Pfad zur Master-CSV
#   -RequiredCsvPath  Pfad zur Bedarfsliste
#   -SearchGlobal     Globalen AD-Scan aktivieren
#   -Parallel         Verarbeitung parallel (Runspace-Pool, empfohlen ab 500+ Eintraegen)
#   -MaxThreads       Anzahl paralleler Threads (nur bei -Parallel, Default: 8)
#   -GroupMode        'Columns' = eine Spalte pro Gruppe (Standard)
#                     'Single'  = alle Gruppen in einer Zelle (semikolonsepariert)
#   -TestCount        Nur erste N Eintraege verarbeiten (Test)
#   -DebugMode        Ausfuehrliches Logging pro Eintrag
#
# BEISPIELE:
#   .\PS_LAccount_Manager.ps1 -CsvPath .\master.csv -Parallel -MaxThreads 12
#   .\PS_LAccount_Manager.ps1 -CsvPath .\master.csv -GroupMode Single

[CmdletBinding()]
param(
    [string] $CsvPath,
    [string] $RequiredCsvPath,
    [switch] $SearchGlobal,
    [switch] $Parallel,
    [int]    $MaxThreads  = 8,
    [ValidateSet("Columns","Single")]
    [string] $GroupMode   = "Columns",
    [int]    $TestCount   = 0,
    [switch] $DebugMode
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

$Version   = "9.6"
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile   = Join-Path $ScriptDir "AD_Sync_$Timestamp.log"
$SW        = [System.Diagnostics.Stopwatch]::StartNew()

$AD_PROPS  = @("DisplayName","Description","GivenName","Surname","l",
               "physicalDeliveryOfficeName","department","info","MemberOf")

# ══════════════════════════════════════════════════════════════════
#  UI & LOGGING  (kein Emoji in -f Format-Strings → Alignment-Fix)
# ══════════════════════════════════════════════════════════════════
function Show-Banner {
    Clear-Host
    $mode  = if ($Parallel) { "Parallel ($MaxThreads Threads)" } else { "Sequenziell" }
    $grpM  = if ($GroupMode -eq "Single") { "Gruppen: 1 Spalte (alle)" } else { "Gruppen: N Spalten" }
    $ts    = Get-Date -Format "dd.MM.yyyy  HH:mm:ss"
    Write-Host ""
    Write-Host "  +----------------------------------------------------------+" -ForegroundColor DarkCyan
    Write-Host "  |  AD COMPLIANCE & SYNC MANAGER  v$Version                      |" -ForegroundColor Cyan
    Write-Host "  |  $ts                               |" -ForegroundColor DarkCyan
    Write-Host "  |  Modus  : $($mode.PadRight(48))|" -ForegroundColor DarkCyan
    Write-Host "  |  Export : $($grpM.PadRight(48))|" -ForegroundColor DarkCyan
    Write-Host "  +----------------------------------------------------------+" -ForegroundColor DarkCyan
    Write-Host ""
}

function Write-Log {
    param(
        [string] $Msg,
        [ValidateSet("INFO","OK","WARN","ERR","DBG","STEP","PERF")] $L = "INFO",
        [switch] $NoFile
    )
    if ($L -eq "DBG" -and -not $DebugMode) { return }
    $ts      = Get-Date -Format "HH:mm:ss.fff"
    $elapsed = "+{0,8:F2}s" -f $SW.Elapsed.TotalSeconds
    $icon    = switch ($L) {
        "OK"   { "[OK]  " } "WARN" { "[WARN]" } "ERR"  { "[ERR] " }
        "DBG"  { "[DBG] " } "STEP" { "[----] " } "PERF" { "[TIME]" }
        default{ "[INFO]" }
    }
    $color = switch ($L) {
        "OK"   { "Green"   } "WARN" { "Yellow" } "ERR"  { "Red"     }
        "DBG"  { "Magenta" } "STEP" { "Cyan"   } "PERF" { "DarkYellow" }
        default{ "Gray"    }
    }
    $line = "[$ts]  $elapsed  $icon  $Msg"
    Write-Host $line -ForegroundColor $color
    if (-not $NoFile) { $line | Out-File $LogFile -Append }
}

function Write-Section {
    param([string]$Title)
    $bar = "-" * 62
    Write-Host ""
    Write-Host "  +$bar+" -ForegroundColor DarkGray
    Write-Host "  |  $($Title.PadRight(60))  |" -ForegroundColor White
    Write-Host "  +$bar+" -ForegroundColor DarkGray
    Write-Log "=== $Title ===" "STEP"
}

function Write-Bar {
    param([int]$Done, [int]$Total, [string]$Label = "")
    if ($Total -le 0) { return }
    $pct     = [math]::Round(($Done / $Total) * 100)
    $fill    = [math]::Round(40 * $pct / 100)
    $bar     = ("#" * $fill) + ("." * (40 - $fill))
    $elapsed = $SW.Elapsed.TotalSeconds
    $etaStr  = ""
    if ($Done -gt 0 -and $Done -lt $Total) {
        $rem    = [math]::Round(($elapsed / $Done) * ($Total - $Done))
        $etaStr = "  ETA ~${rem}s"
    }
    Write-Progress -Activity $Label `
        -Status "[$bar] $Done / $Total  ($pct%)$etaStr" `
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
    return (($DN -split ',')[0] -replace 'CN=','')
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

# Schneller CSV-Export: direkt alle Property-Values in einem Zug holen
# KEIN foreach-über-Spaltennamen → O(Props) statt O(Props) mit Name-Lookup
function Export-FastCsv {
    param(
        [string] $Path,
        [array]  $ColNames,           # geordnete Spaltenliste
        [System.Collections.Generic.List[PSObject]] $Data,
        [string] $Label = "Export"
    )
    $t0     = $SW.Elapsed.TotalSeconds
    $total  = $Data.Count
    $writer = New-Object System.IO.StreamWriter($Path, $false, [System.Text.Encoding]::UTF8)
    try {
        # Header
        $writer.WriteLine([string]::Join(';', $ColNames))

        $idx = 0
        foreach ($r in $Data) {
            $idx++
            # Alle Values auf einmal holen – deutlich schneller als $r.$colName in einer Schleife
            $props = $r.psobject.Properties
            $vals  = foreach ($c in $ColNames) {
                $v = $props[$c].Value
                $s = if ($null -ne $v) { $v.ToString() } else { "" }
                if ($s -match '[;\r\n"]') { '"' + $s.Replace('"','""') + '"' } else { $s }
            }
            $writer.WriteLine([string]::Join(';', $vals))

            if ($idx % 100 -eq 0 -or $idx -eq $total) {
                $pct  = [math]::Round(($idx / $total) * 100)
                $fill = [math]::Round(30 * $pct / 100)
                $bar  = ("#" * $fill) + ("." * (30 - $fill))
                Write-Host ("`r  [{0}] {1,4}/{2}  ({3,3}%)    " -f $bar, $idx, $total, $pct) `
                    -NoNewline -ForegroundColor DarkGray
                Write-Progress -Activity $Label -Status "$idx / $total  ($pct%)" -PercentComplete $pct
            }
        }
    } finally {
        $writer.Close()
    }
    Write-Host ""
    Write-Progress -Activity $Label -Completed
    Write-Log ("$Label fertig: {0}  [{1:F2}s]" -f (Split-Path $Path -Leaf), ($SW.Elapsed.TotalSeconds - $t0)) "OK"
}

# ══════════════════════════════════════════════════════════════════
#  RECORD-BERECHNUNG  (als Funktion – fuer seq. UND parallel nutzbar)
# ══════════════════════════════════════════════════════════════════
# Hinweis: In Runspaces wird diese Funktion als ScriptBlock uebergeben.
# Alle Logikaeste werden VOR dem Hashtable in Variablen berechnet.
$CalcScriptBlock = {
    param($LID, $CsvRow, $ADObj, $RequiredList, $SortedGroups, $GroupMode)

    function Get-StrI { param($v)
        if ($null -eq $v) { return "" }; return $v.ToString().Trim()
    }
    function Test-DiffI { param($a,$b)
        return ((Get-StrI $a) -ine (Get-StrI $b))
    }
    function Get-GrpNameI { param([string]$DN)
        return (($DN -split ',')[0] -replace 'CN=','')
    }
    function Get-VerfahrenI { param($Row)
        if ($null -eq $Row) { return "[alle_Verfahren]" }
        $cols = @("Viva","Findus","MobiApps","AccVisio","Verfahren5","Verfahren6","Verfahren7","Verfahren8","Verfahren9","Verfahren10")
        $hit  = foreach ($c in $cols) {
            if ($Row.$c -and $Row.$c.ToString().Trim().ToLower() -eq "x") { $c }
        }
        if ($hit) { return ($hit -join " - ") }
        return "[alle_Verfahren]"
    }

    # AD-Basiswerte
    $andereOU        = ""
    $statusGeloescht = ""
    $userGroups      = @()

    if ($ADObj) {
        $dn = $ADObj.DistinguishedName
        if ($dn -notmatch "OU=81" -and $dn -notmatch "OU=82") {
            $andereOU = $dn -replace '^CN=.*?,',''
        }
        if ($ADObj.MemberOf) {
            $userGroups = @($ADObj.MemberOf | ForEach-Object { Get-GrpNameI -DN $_ })
        }
    } else {
        $statusGeloescht = "XXX"
    }

    # Berechnete Werte
    $isRequired = ""
    if ($RequiredList.Contains($LID.ToUpper())) { $isRequired = "ZZZ" }

    $isLafp  = ($LID -like "L110*") -or ($LID -like "L114*")
    $lafpStr = ""
    $ouNum   = "[OU]"
    if ($isLafp) { $lafpStr = "LLL"; $ouNum = "26" }

    $targetOrt = ""
    if ($CsvRow -and $CsvRow.Standort -and $CsvRow.Standort.ToString().Trim() -ne "") {
        $targetOrt = $CsvRow.Standort.ToString().Trim()
    } elseif ($ADObj -and $ADObj.l) {
        $targetOrt = Get-StrI $ADObj.l
    }

    $targetNachname = $targetOrt
    if ($isLafp -and $targetOrt -notmatch "^LAFP\s-\s") { $targetNachname = "LAFP - $targetOrt" }

    $buroPlatz = "[Platznummer]"
    if ($CsvRow -and $CsvRow.Raum -and $CsvRow.Raum.ToString().Trim() -ne "") {
        $buroPlatz = $CsvRow.Raum.ToString().Trim()
    }

    $fort  = if ($CsvRow -and $CsvRow.Fortbildungsbereich) { Get-StrI $CsvRow.Fortbildungsbereich } else { "" }
    $anm   = if ($CsvRow -and $CsvRow.Anmerkungen)         { Get-StrI $CsvRow.Anmerkungen }         else { "" }
    $aInfo = if ($fort -ne "" -and $anm -ne "") { "$fort - $anm" } else { "$fort$anm" }
    $verf  = Get-VerfahrenI -Row $CsvRow

    $adVn   = if ($ADObj) { Get-StrI $ADObj.GivenName }                  else { "" }
    $adNn   = if ($ADObj) { Get-StrI $ADObj.Surname }                    else { "" }
    $adDn   = if ($ADObj) { Get-StrI $ADObj.DisplayName }                else { "" }
    $adOrt  = if ($ADObj) { Get-StrI $ADObj.l }                          else { "" }
    $adBue  = if ($ADObj) { Get-StrI $ADObj.physicalDeliveryOfficeName } else { "" }
    $adDez  = if ($ADObj) { Get-StrI $ADObj.department }                 else { "" }
    $adDesc = if ($ADObj) { Get-StrI $ADObj.Description }                else { "" }
    $adInf  = if ($ADObj) { Get-StrI $ADObj.info }                       else { "" }

    $oStand = if ($CsvRow) { Get-StrI $CsvRow.Standort  } else { "" }
    $oRaum  = if ($CsvRow) { Get-StrI $CsvRow.Raum      } else { "" }
    $oAbt   = if ($CsvRow) { Get-StrI $CsvRow.Abteilung } else { "" }

    $aeDez  = if ($CsvRow -and $ADObj -and $adDez -ne "") { "$oAbt - $adDez" } else { $oAbt }
    $aeDn   = "$LID $targetNachname"
    $aeDesc = "$ouNum - $verf - $buroPlatz - [Verantwortlicher] - [TEL] | $aInfo"

    $Props = [ordered]@{
        "L-Kennung"                    = $LID
        "andere_OU"                    = $andereOU
        "GELOESCHT"                    = $statusGeloescht
        "Benoetigt"                    = $isRequired
        "Geaendert"                    = ""
        "LAFP_LZPD_LKA"               = $lafpStr
        "NICHT_KONFORM"               = ""
        "AD_Vorname"                  = $adVn
        "AENDERN_Vorname"             = $LID
        "AD_Nachname"                 = $adNn
        "AENDERN_Nachname"            = $targetNachname
        "AD_DisplayName"              = $adDn
        "AENDERN_DisplayName"         = $aeDn
        "ORIGINAL_Standort"           = $oStand
        "ORIGINAL_Raum_Schulungskreis"= $oRaum
        "AD_Ort"                      = $adOrt
        "AENDERN_Ort"                 = $targetOrt
        "AD_Buero"                    = $adBue
        "AENDERN_Buero"               = $buroPlatz
        "ORIGINAL_Abteilung"          = $oAbt
        "AD_Dez"                      = $adDez
        "AENDERN_Dez"                 = $aeDez
        "AD_Description"              = $adDesc
        "AENDERN_Description"         = $aeDesc
        "AENDERN_OU"                  = $ouNum
        "AENDERN_Info"                = $aInfo
        "LOESCHEN"                    = ""
    }

    # Konformitaet
    if ($ADObj) {
        $codes = New-Object 'System.Collections.Generic.List[string]'
        if (Test-DiffI $adVn   $LID)            { $codes.Add("VN")   }
        if (Test-DiffI $adNn   $targetNachname) { $codes.Add("NN")   }
        if (Test-DiffI $adDn   $aeDn)           { $codes.Add("DN")   }
        if (Test-DiffI $adDesc $aeDesc)         { $codes.Add("DESC") }
        if (Test-DiffI $adOrt  $targetOrt)      { $codes.Add("ORT")  }
        if (Test-DiffI $adBue  $buroPlatz)      { $codes.Add("GEB")  }
        if (Test-DiffI $adDez  $aeDez)          { $codes.Add("DEZ")  }
        if (Test-DiffI $adInf  $aInfo)          { $codes.Add("INFO") }
        if ($codes.Count -gt 0) {
            $cs = $codes -join ","
            $Props."Geaendert"    = $cs
            $Props."NICHT_KONFORM" = $cs
        }
    }

    # Gruppen
    if ($GroupMode -eq "Single") {
        # Alle Gruppen des Users in einer Zelle (semikolonsepariert)
        $Props["Gruppen"] = $userGroups -join ";"
    } else {
        # Eine Spalte pro Gruppe
        foreach ($gn in $SortedGroups) {
            $inGrp = ($userGroups -contains $gn)
            $Props["GRP_$gn"] = if ($inGrp) { "X" } else { "" }
        }
    }

    return [PSCustomObject]$Props
}

# ══════════════════════════════════════════════════════════════════
#  EXTRAHIERTE FUNKTIONEN (MAIN-HELPER)
# ══════════════════════════════════════════════════════════════════
function Get-RequiredList {
    param([string]$RequiredCsvPath)
    Write-Section "BEDARFSLISTE laden"
    $RequiredList = New-Object 'System.Collections.Generic.HashSet[string]'
    if ($RequiredCsvPath -and (Test-Path $RequiredCsvPath)) {
        foreach ($line in (Get-Content $RequiredCsvPath -Encoding UTF8)) {
            if ($line -match "(L\d{6,8})") {
                $v = $Matches[1].ToUpper()
                [void]$RequiredList.Add($v)
            }
        }
        Write-Log "Bedarfsliste: $($RequiredList.Count) Eintraege" "OK"
    } else {
        Write-Log "Keine Bedarfsliste -> uebersprungen" "WARN"
    }
    return $RequiredList
}

function Get-MasterCsvData {
    param(
        [string]$CsvPath,
        [System.Diagnostics.Stopwatch]$SW
    )
    Write-Section "MASTER-CSV laden"
    $t0 = $SW.Elapsed.TotalSeconds
    $MasterCsvData = @{}
    foreach ($row in (Import-Csv -Path $CsvPath -Delimiter ';' -Encoding Default)) {
        if ($row."L-Kennung") {
            $k = $row."L-Kennung".ToString().Trim().ToUpper()
            if ($k) { $MasterCsvData[$k] = $row }
        }
    }
    Write-Log "Master-CSV: $($MasterCsvData.Count) Zeilen  [{0:F2}s]" -f ($SW.Elapsed.TotalSeconds-$t0) | Out-Null
    Write-Log ("Master-CSV: {0} Zeilen  [{1:F2}s]" -f $MasterCsvData.Count, ($SW.Elapsed.TotalSeconds-$t0)) "OK"
    return $MasterCsvData
}

function Get-ADData {
    param(
        [switch]$SearchGlobal,
        [System.Diagnostics.Stopwatch]$SW,
        [array]$AD_PROPS
    )
    Write-Section "AD-DISCOVERY"
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
        $users = Get-ADUser -Filter * -SearchBase $ou.DistinguishedName -Properties $AD_PROPS
        Write-Log ("OU {0}: {1} User  [{2:F2}s]" -f $ou.Name, @($users).Count, ($SW.Elapsed.TotalSeconds-$t1)) "DBG"

        foreach ($u in $users) {
            $sam = $u.SamAccountName.ToUpper()
            $ADCache[$sam] = $u
            foreach ($gDN in $u.MemberOf) {
                $gn = Get-GroupName -DN $gDN
                [void]$UniqueGroups.Add($gn)
            }
        }
        Write-Log "OU $($ou.Name): AD-Cache jetzt $($ADCache.Count) Eintraege" "DBG"
    }

    if ($SearchGlobal) {
        Write-Log "Globaler Scan (L*)..." "STEP"
        $t1   = $SW.Elapsed.TotalSeconds
        $gAll = Get-ADUser -Filter "SamAccountName -like 'L*'" -Properties $AD_PROPS
        Write-Log ("Globaler Scan: {0} Treffer  [{1:F2}s]" -f @($gAll).Count, ($SW.Elapsed.TotalSeconds-$t1)) "DBG"
        $newN = 0
        foreach ($gu in $gAll) {
            $sam = $gu.SamAccountName.ToUpper()
            if (-not $ADCache.ContainsKey($sam)) {
                $ADCache[$sam] = $gu
                $newN++
                foreach ($gDN in $gu.MemberOf) {
                    $gn = Get-GroupName -DN $gDN
                    [void]$UniqueGroups.Add($gn)
                }
            }
        }
        Write-Log "$newN neue Eintraege aus globalem Scan" "OK"
    }

    $SortedGroups = @($UniqueGroups | Sort-Object)
    Write-Log ("Discovery fertig: {0} AD-Objekte  {1} Gruppen  [{2:F2}s]" -f `
        $ADCache.Count, $SortedGroups.Count, ($SW.Elapsed.TotalSeconds-$t0)) "PERF"

    return [PSCustomObject]@{
        ADCache      = $ADCache
        SortedGroups = $SortedGroups
    }
}

function Invoke-Processing {
    param(
        [array]$ProcessList,
        [hashtable]$MasterCsvData,
        [hashtable]$ADCache,
        [System.Collections.Generic.HashSet[string]]$RequiredList,
        [array]$SortedGroups,
        [string]$GroupMode,
        [switch]$Parallel,
        [int]$MaxThreads,
        [int]$TestCount,
        [switch]$DebugMode,
        [System.Diagnostics.Stopwatch]$SW,
        [scriptblock]$CalcScriptBlock,
        [string]$LogFile
    )

    $modeLabel = if ($Parallel) { "PARALLEL ($MaxThreads Threads)" } else { "SEQUENZIELL" }
    Write-Section "VERARBEITUNG  ($($ProcessList.Count) Eintraege  ·  $modeLabel)"
    $t0 = $SW.Elapsed.TotalSeconds

    $Results = New-Object 'System.Collections.Generic.List[PSObject]'
    $Total   = $ProcessList.Count

    if ($Parallel) {
        # ── PARALLEL (Runspace-Pool) ─────────────────────────────
        Write-Log "Starte Runspace-Pool mit $MaxThreads Threads..." "STEP"

        $Pool = [runspacefactory]::CreateRunspacePool(
            1, $MaxThreads,
            [system.management.automation.runspaces.initialsessionstate]::CreateDefault(),
            $Host)
        $Pool.Open()

        # Jobs aufbauen
        $Jobs = New-Object 'System.Collections.Generic.List[PSObject]'
        foreach ($LID in $ProcessList) {
            $psi = [powershell]::Create()
            [void]$psi.AddScript($CalcScriptBlock)
            [void]$psi.AddArgument($LID)
            [void]$psi.AddArgument($MasterCsvData[$LID])
            [void]$psi.AddArgument($ADCache[$LID])
            [void]$psi.AddArgument($RequiredList)
            [void]$psi.AddArgument($SortedGroups)
            [void]$psi.AddArgument($GroupMode)
            $psi.RunspacePool = $Pool
            $entry = New-Object PSObject -Property @{
                SAM         = $LID
                Instance    = $psi
                AsyncResult = $psi.BeginInvoke()
            }
            [void]$Jobs.Add($entry)
        }
        Write-Log "$($Jobs.Count) Jobs gestartet" "DBG"

        $done     = 0
        $errCount = 0
        while ($Jobs.Count -gt 0) {
            $finished = @($Jobs | Where-Object { $_.AsyncResult.IsCompleted })
            $remove   = New-Object 'System.Collections.Generic.List[PSObject]'

            foreach ($job in $finished) {
                $done++

                # Fehler-Stream
                foreach ($e in $job.Instance.Streams.Error) {
                    $errCount++
                    Write-Log "Runspace $($job.SAM): $e" "ERR"
                }

                # Ergebnis einsammeln
                $raw = $job.Instance.EndInvoke($job.AsyncResult)
                foreach ($item in $raw) {
                    if ($null -ne $item -and $null -ne $item.PSObject) {
                        $tn = $item.GetType().FullName
                        if ($tn -notlike "*Collection*") {
                            [void]$Results.Add($item)
                        }
                    }
                }

                if ($DebugMode) {
                    Write-Host ("  [{0,4}/{1}]  {2}" -f $done, $Total, $job.SAM) `
                        -ForegroundColor DarkGray
                }

                $job.Instance.Dispose()
                [void]$remove.Add($job)
            }
            foreach ($j in $remove) { [void]$Jobs.Remove($j) }

            if ($done % 50 -eq 0 -or ($Jobs.Count -eq 0 -and $done -eq $Total)) {
                Write-Bar -Done $done -Total $Total -Label "Parallel-Verarbeitung"
                $elapsed = $SW.Elapsed.TotalSeconds
                $rem     = if ($done -gt 0) { [math]::Round(($elapsed/$done)*($Total-$done)) } else { 0 }
                Write-Host ("`r  [{0}]  {1,4}/{2}  ETA ~{3}s    " -f `
                    ("#" * [math]::Round(30*$done/$Total)) + ("." * (30 - [math]::Round(30*$done/$Total))), `
                    $done, $Total, $rem) -NoNewline -ForegroundColor Cyan
            }

            if ($Jobs.Count -gt 0) { Start-Sleep -Milliseconds 30 }
        }
        Write-Host ""
        Write-Progress -Activity "Parallel-Verarbeitung" -Completed
        $Pool.Close(); $Pool.Dispose()

        if ($errCount -gt 0) {
            Write-Log "$errCount Runspace-Fehler. Details im Log: $LogFile" "WARN"
        }

    } else {
        # ── SEQUENZIELL ──────────────────────────────────────────
        $idx = 0
        foreach ($LID in $ProcessList) {
            $idx++
            $tItem = $SW.Elapsed.TotalSeconds

            # Direkt den ScriptBlock als Skript aufrufen (kein Runspace-Overhead)
            $item = & $CalcScriptBlock `
                $LID `
                $MasterCsvData[$LID] `
                $ADCache[$LID] `
                $RequiredList `
                $SortedGroups `
                $GroupMode

            [void]$Results.Add($item)

            if ($DebugMode) {
                $ms  = [math]::Round(($SW.Elapsed.TotalSeconds - $tItem) * 1000, 1)
                $src = if ($ADCache[$LID] -and $MasterCsvData[$LID]) { "AD+CSV" } `
                       elseif ($ADCache[$LID]) { "AD   " } else { "CSV  " }
                Write-Host ("  [{0,4}/{1}]  {2,-14}  src:{3}  nk:{4,-20}  {5,5}ms" -f `
                    $idx, $Total, $LID, $src, $item."NICHT_KONFORM", $ms) -ForegroundColor DarkGray
            }

            if ($idx % 50 -eq 0 -or $idx -eq $Total) {
                $elapsed = $SW.Elapsed.TotalSeconds
                $rem     = if ($idx -gt 0) { [math]::Round(($elapsed/$idx)*($Total-$idx)) } else { 0 }
                $pct     = [math]::Round(($idx/$Total)*100)
                $fill    = [math]::Round(30*$pct/100)
                $bar     = ("#" * $fill) + ("." * (30-$fill))
                $etaStr  = if ($idx -lt $Total) { "ETA ~${rem}s" } else { "fertig!   " }
                Write-Host ("`r  [{0}]  {1,4}/{2}  ({3,3}%)  {4}    " -f $bar, $idx, $Total, $pct, $etaStr) `
                    -NoNewline -ForegroundColor Cyan
                Write-Progress -Activity "Sequenzielle Verarbeitung" `
                    -Status "$idx / $Total  ($pct%)  ~${rem}s" `
                    -PercentComplete $pct
            }
        }
        Write-Host ""
        Write-Progress -Activity "Sequenzielle Verarbeitung" -Completed
    }

    $procSec = [math]::Round($SW.Elapsed.TotalSeconds - $t0, 2)
    $msPerEntry = if ($Results.Count -gt 0) { [math]::Round($procSec/$Results.Count*1000,1) } else { 0 }
    Write-Log ("Verarbeitung: {0} Eintraege in {1}s  ({2}ms/Eintrag)" -f `
        $Results.Count, $procSec, $msPerEntry) "PERF"

    return $Results
}

function Export-Results {
    param(
        [System.Collections.Generic.List[PSObject]]$Results,
        [string]$GroupMode,
        [array]$SortedGroups,
        [string]$ScriptDir,
        [string]$Version
    )
    Write-Section "EXPORT"

    if ($Results.Count -eq 0) { Write-Log "Keine Ergebnisse -> Abbruch" "ERR"; return }

    # Spaltenlisten
    $cols1Base = @("L-Kennung","andere_OU","GELOESCHT","Benoetigt","Geaendert",
                   "LOESCHEN","AD_Vorname","AD_Nachname")
    if ($GroupMode -eq "Single") {
        $cols1 = $cols1Base + @("Gruppen")
    } else {
        $grpCols = @($SortedGroups | ForEach-Object { "GRP_$_" })
        $cols1   = $cols1Base + $grpCols
    }
    $cols2 = @($Results[0].psobject.Properties.Name | Where-Object { $_ -ne "LOESCHEN" })

    Write-Log "Tab1: $($cols1.Count) Spalten  Tab2: $($cols2.Count) Spalten" "DBG"

    $Path1 = Join-Path $ScriptDir "L-Kennungen_Full_Analysis_v$Version.csv"
    $Path2 = Join-Path $ScriptDir "L-Kennungen_Properties_v$Version.csv"

    Export-FastCsv -Path $Path1 -ColNames $cols1 -Data $Results -Label "Tab1 (Gruppen-Uebersicht)"
    Export-FastCsv -Path $Path2 -ColNames $cols2 -Data $Results -Label "Tab2 (Properties)"

    return [PSCustomObject]@{
        Path1 = $Path1
        Path2 = $Path2
    }
}

# ══════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════
function Invoke-Main {
    Show-Banner

    Import-Module ActiveDirectory -Verbose:$false

    Write-Log "PS v$($PSVersionTable.PSVersion)  PID=$PID  Host=$env:COMPUTERNAME" "DBG"
    Write-Log "Modus: $(if($Parallel){"Parallel (Threads=$MaxThreads)"}else{"Sequenziell"})  GroupMode=$GroupMode" "INFO"

    if (-not $CsvPath) { $CsvPath = (Read-Host "Pfad zur Master-CSV").Trim().Trim('"') }
    $CsvPath = $CsvPath.Trim().Trim('"')
    if (-not (Test-Path $CsvPath)) { Write-Log "CSV nicht gefunden: $CsvPath" "ERR"; return }

    # ── 1. BEDARFSLISTE ──────────────────────────────────────────
    $RequiredList = Get-RequiredList -RequiredCsvPath $RequiredCsvPath

    # ── 2. MASTER-CSV ────────────────────────────────────────────
    $MasterCsvData = Get-MasterCsvData -CsvPath $CsvPath -SW $SW

    # ── 3. AD-DISCOVERY ──────────────────────────────────────────
    $adResult = Get-ADData -SearchGlobal:$SearchGlobal -SW $SW -AD_PROPS $AD_PROPS
    $ADCache      = $adResult.ADCache
    $SortedGroups = $adResult.SortedGroups

    if ($GroupMode -eq "Columns") {
        Write-Log "GroupMode=Columns: $($SortedGroups.Count) Gruppenspalten → Tab2 wird breit" "WARN"
    } else {
        Write-Log "GroupMode=Single: alle Gruppen in Spalte 'Gruppen' (semikolonsepariert)" "INFO"
    }

    $AllSAMs     = @(@($ADCache.Keys) + @($MasterCsvData.Keys) | Select-Object -Unique | Sort-Object)
    $ProcessList = if ($TestCount -gt 0) {
        Write-Log "TESTMODUS: erste $TestCount Eintraege" "WARN"
        @($AllSAMs | Select-Object -First $TestCount)
    } else { $AllSAMs }
    Write-Log "Zu verarbeitende SAMs: $($ProcessList.Count)" "OK"

    # ── 4. VERARBEITUNG ──────────────────────────────────────────
    $Results = Invoke-Processing `
        -ProcessList $ProcessList `
        -MasterCsvData $MasterCsvData `
        -ADCache $ADCache `
        -RequiredList $RequiredList `
        -SortedGroups $SortedGroups `
        -GroupMode $GroupMode `
        -Parallel:$Parallel `
        -MaxThreads $MaxThreads `
        -TestCount $TestCount `
        -DebugMode:$DebugMode `
        -SW $SW `
        -CalcScriptBlock $CalcScriptBlock `
        -LogFile $LogFile

    # ── 5. EXPORT ────────────────────────────────────────────────
    if ($Results -and $Results.Count -gt 0) {
        $exportResult = Export-Results `
            -Results $Results `
            -GroupMode $GroupMode `
            -SortedGroups $SortedGroups `
            -ScriptDir $ScriptDir `
            -Version $Version
        $Path1 = $exportResult.Path1
        $Path2 = $exportResult.Path2
    } else {
        $Path1 = ""
        $Path2 = ""
    }

    # ── ABSCHLUSS ────────────────────────────────────────────────
    $SW.Stop()
    $mb = [math]::Round([System.GC]::GetTotalMemory($false) / 1MB, 1)
    $dur = $SW.Elapsed.ToString('mm\:ss\.ff')

    Write-Host ""
    Write-Host "  +----------------------------------------------------------+" -ForegroundColor Green
    Write-Host "  |  ABGESCHLOSSEN                                           |" -ForegroundColor Green
    Write-Host ("  |  Datensaetze  : {0,-42}|" -f $Results.Count)              -ForegroundColor Green
    Write-Host ("  |  Gruppen      : {0,-42}|" -f "$($SortedGroups.Count)  (Mode: $GroupMode)") -ForegroundColor Green
    Write-Host ("  |  Gesamtdauer  : {0,-42}|" -f "$dur min")                  -ForegroundColor Green
    Write-Host ("  |  RAM (managed): {0,-42}|" -f "${mb} MB")                  -ForegroundColor Green
    Write-Host "  |                                                          |" -ForegroundColor DarkGreen
    Write-Host ("  |  Tab1 : {0,-50}|" -f (Split-Path $Path1 -Leaf))           -ForegroundColor DarkGreen
    Write-Host ("  |  Tab2 : {0,-50}|" -f (Split-Path $Path2 -Leaf))           -ForegroundColor DarkGreen
    Write-Host ("  |  Log  : {0,-50}|" -f (Split-Path $LogFile -Leaf))         -ForegroundColor DarkGreen
    Write-Host "  +----------------------------------------------------------+" -ForegroundColor Green
    Write-Host ""

    Write-Log ("Fertig. Dauer={0}  RAM={1}MB  Zeilen={2}  Gruppen={3}  Modus={4}" -f `
        $dur, $mb, $Results.Count, $SortedGroups.Count, $(if($Parallel){"Parallel"}else{"Seq"})) "PERF"
}

Invoke-Main
