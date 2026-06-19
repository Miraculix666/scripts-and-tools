# FileName: PS_LAccount_Manager.ps1
# Version:  10.0
# Author:   PS-Coding
#
# AENDERUNGEN v10.0 gegenueber v9.6:
#   [1] Parallel immer aktiv (kein -Parallel Switch mehr)
#   [2] GroupMode Single ist Standard (alle Gruppen in einer Zelle)
#   [3] Keine GRP_-Prefix auf Gruppenspaltennamen
#   [4] Spaltennamen 100% kompatibel mit PS_LAccount_Apply v2.0
#   [5] Suchbereich: SearchScope Subtree auf OU=81 UND OU=82
#       -> alle Sub-OUs werden immer erfasst, auch unbekannte
#   [6] Neue Spalte OU: zeigt den unmittelbaren AD-Container des Kontos
#   [7] -GroupFilter: erzeugt Markerspalte FILTER_<Wert> mit X wenn
#       Gruppenname den Teilstring enthaelt (case-insensitiv)
#       -> dient als Selektionshilfe in Excel / Apply
#   [8] -OUFilter: beschleunigt Discovery, nur User aus passenden OUs
#   [9] -NameFilter: beschleunigt Verarbeitung, nur passende SAMAccountNames
#   [10] -LafpOnly: beschleunigt Verarbeitung, nur LAFP-Kennungen (L110/L114)
#
# PARAMETER:
#   -CsvPath          Pfad zur Master-CSV (Pflicht)
#   -RequiredCsvPath  Pfad zur Bedarfsliste (optional)
#   -MaxThreads       Anzahl paralleler Threads (Default: 8)
#   -GroupMode        'Single' = alle Gruppen in einer Zelle (Standard)
#                     'Columns' = eine Spalte pro Gruppe
#   -GroupFilter      Teilstring (case-insensitiv): Gruppen die passen erhalten
#                     X in der Spalte FILTER_<GroupFilter>.
#                     z.B. -GroupFilter "test" -> Spalte FILTER_test
#   -OUFilter         Nur User aus OUs deren DN diesen Teilstring enthaelt.
#                     Wird WAEHREND Discovery angewendet -> beschleunigt Scan.
#                     z.B. -OUFilter "OU=81" oder -OUFilter "Benutzer"
#   -NameFilter       Nur SAMAccountNames die diesem Muster entsprechen.
#                     Wildcards moeglich. Wird VOR Verarbeitung angewendet.
#                     z.B. -NameFilter "L110*"  oder -NameFilter "*1234*"
#   -LafpOnly         Nur LAFP-Kennungen (L110* und L114*) verarbeiten.
#                     Kombinierbar mit -NameFilter.
#   -TestCount        Nur erste N Eintraege (Testmodus)
#   -DebugMode        Ausfuehrliches Logging pro Eintrag
#
# BEISPIELE:
#   # Standardlauf:
#   .\PS_LAccount_Manager.ps1 -CsvPath .\master.csv
#
#   # Nur OU 81, alle Gruppen markieren die "Schulung" enthalten:
#   .\PS_LAccount_Manager.ps1 -CsvPath .\master.csv -OUFilter "OU=81" -GroupFilter "Schulung"
#
#   # Nur LAFP-Konten, Gruppen als Spalten:
#   .\PS_LAccount_Manager.ps1 -CsvPath .\master.csv -LafpOnly -GroupMode Columns
#
#   # Schneller Test mit 20 Eintraegen und Debug:
#   .\PS_LAccount_Manager.ps1 -CsvPath .\master.csv -TestCount 20 -DebugMode

[CmdletBinding()]
param(
    [string] $CsvPath,
    [string] $RequiredCsvPath,

    # Verarbeitung
    [int]    $MaxThreads   = 8,

    # Gruppen-Export
    [ValidateSet("Single","Columns")]
    [string] $GroupMode    = "Single",
    [string] $GroupFilter  = "",        # Teilstring-Filter -> Markerspalte FILTER_<x>

    # Prozess-Filter (beschleunigen den Lauf)
    [string] $OUFilter     = "",        # Teilstring im OU-DN -> eingeschraenkter AD-Scan
    [string] $NameFilter   = "",        # SAMAccountName-Wildcard -> eingeschraenkte Verarbeitung
    [switch] $LafpOnly,                 # nur L110* und L114*

    [int]    $TestCount    = 0,
    [switch] $DebugMode
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

$Version   = "10.0"
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile   = Join-Path $ScriptDir "AD_Sync_$Timestamp.log"
$SW        = [System.Diagnostics.Stopwatch]::StartNew()

# AD-Properties die geladen werden muessen
$AD_PROPS = @(
    "DisplayName","Description","GivenName","Surname","l",
    "physicalDeliveryOfficeName","department","info","MemberOf"
)

# ══════════════════════════════════════════════════════════════════
#  UI & LOGGING
# ══════════════════════════════════════════════════════════════════
function Show-Banner {
    Clear-Host
    $filters = @()
    if ($OUFilter    -ne "") { $filters += "OU=$OUFilter"       }
    if ($NameFilter  -ne "") { $filters += "Name=$NameFilter"   }
    if ($LafpOnly)           { $filters += "LafpOnly"           }
    if ($GroupFilter -ne "") { $filters += "GrpFilter=$GroupFilter" }
    $filterStr = if ($filters.Count -gt 0) { $filters -join "  |  " } else { "keine" }

    Write-Host ""
    Write-Host "  +----------------------------------------------------------+" -ForegroundColor DarkCyan
    Write-Host "  |  AD COMPLIANCE & SYNC MANAGER  v$Version                     |" -ForegroundColor Cyan
    Write-Host "  |  $(Get-Date -Format 'dd.MM.yyyy  HH:mm:ss')                               |" -ForegroundColor DarkCyan
    Write-Host ("  |  Threads  : {0,-46}|" -f $MaxThreads)                        -ForegroundColor DarkCyan
    Write-Host ("  |  GroupMode: {0,-46}|" -f $GroupMode)                         -ForegroundColor DarkCyan
    Write-Host ("  |  Filter   : {0,-46}|" -f $filterStr)                         -ForegroundColor DarkYellow
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
    $tag     = switch ($L) {
        "OK"   { "[OK]   " } "WARN" { "[WARN] " } "ERR"  { "[ERR]  " }
        "DBG"  { "[DBG]  " } "STEP" { "[----] " } "PERF" { "[TIME] " }
        default{ "[INFO] " }
    }
    $color = switch ($L) {
        "OK"   { "Green"      } "WARN" { "Yellow" } "ERR"  { "Red"         }
        "DBG"  { "Magenta"    } "STEP" { "Cyan"   } "PERF" { "DarkYellow"  }
        default{ "Gray"       }
    }
    $line = "[$ts]  $elapsed  $tag  $Msg"
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

# Schneller StreamWriter-Export (O(1) RAM)
function Export-FastCsv {
    param(
        [string] $Path,
        [array]  $ColNames,
        [System.Collections.Generic.List[PSObject]] $Data,
        [string] $Label = "Export"
    )
    $t0     = $SW.Elapsed.TotalSeconds
    $total  = $Data.Count
    $writer = New-Object System.IO.StreamWriter($Path, $false, [System.Text.Encoding]::UTF8)
    try {
        $writer.WriteLine([string]::Join(';', $ColNames))
        $idx = 0
        foreach ($r in $Data) {
            $idx++
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
                Write-Host ("`r  [{0}] {1,4}/{2}  ({3,3}%)  " -f $bar, $idx, $total, $pct) `
                    -NoNewline -ForegroundColor DarkGray
                Write-Progress -Activity $Label -Status "$idx / $total  ($pct%)" -PercentComplete $pct
            }
        }
    } finally { $writer.Close() }
    Write-Host ""
    Write-Progress -Activity $Label -Completed
    Write-Log ("{0} fertig: {1}  [{2:F2}s]" -f $Label, (Split-Path $Path -Leaf), ($SW.Elapsed.TotalSeconds-$t0)) "OK"
}

function Get-GroupName { param([string]$DN)
    return (($DN -split ',')[0] -replace 'CN=','')
}

# OU des Benutzers aus DN extrahieren (unmittelbarer Container)
function Get-UserOU { param([string]$DN)
    # DN: CN=L1234567,OU=Benutzer,OU=82,...
    # Entfernt CN= Teil, gibt Rest zurueck als OU-Pfad
    $withoutCN = $DN -replace '^CN=[^,]+,',''
    # Nur die erste OU-Ebene (unmittelbarer Container)
    $firstPart = ($withoutCN -split ',')[0] -replace 'OU=',''
    return $firstPart
}

# Vollstaendiger OU-Pfad (fuer Spalte OU_Pfad)
function Get-UserOUPath { param([string]$DN)
    return ($DN -replace '^CN=[^,]+,','')
}

# ══════════════════════════════════════════════════════════════════
#  RECORD-BERECHNUNG (ScriptBlock fuer Runspaces)
#  Alle bedingten Werte VOR dem Hashtable vorberechnen.
#  Kein inline-if im Hashtable. Kein return-if. Kein .Add(Ausdruck).
# ══════════════════════════════════════════════════════════════════
$CalcScriptBlock = {
    param(
        [string] $LID,
        $CsvRow,
        $ADObj,
        $RequiredList,
        $SortedGroups,
        [string] $GroupMode,
        [string] $GroupFilter
    )

    function Get-StrI { param($v)
        if ($null -eq $v) { return "" }
        return $v.ToString().Trim()
    }
    function Test-DiffI { param($a, $b)
        return ((Get-StrI $a) -ine (Get-StrI $b))
    }
    function Get-GrpNameI { param([string]$DN)
        return (($DN -split ',')[0] -replace 'CN=','')
    }
    function Get-VerfahrenI { param($Row)
        if ($null -eq $Row) { return "[alle_Verfahren]" }
        $cols = @("Viva","Findus","MobiApps","AccVisio",
                  "Verfahren5","Verfahren6","Verfahren7","Verfahren8","Verfahren9","Verfahren10")
        $hit  = foreach ($c in $cols) {
            if ($Row.$c -and $Row.$c.ToString().Trim().ToLower() -eq "x") { $c }
        }
        if ($hit) { return ($hit -join " - ") }
        return "[alle_Verfahren]"
    }

    # ── AD-Basiswerte ────────────────────────────────────────────
    $andereOU        = ""
    $ouName          = ""
    $ouPfad          = ""
    $statusGeloescht = ""
    $userGroups      = @()

    if ($ADObj) {
        $dn     = $ADObj.DistinguishedName
        $ouPfad = $dn -replace '^CN=[^,]+,',''
        $ouName = ($ouPfad -split ',')[0] -replace 'OU=',''

        if ($dn -notmatch "OU=81" -and $dn -notmatch "OU=82") {
            $andereOU = $ouPfad
        }
        if ($ADObj.MemberOf) {
            $userGroups = @($ADObj.MemberOf | ForEach-Object { Get-GrpNameI -DN $_ })
        }
    } else {
        $statusGeloescht = "XXX"
    }

    # ── Berechnete Werte ──────────────────────────────────────────
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
    if ($isLafp -and $targetOrt -notmatch "^LAFP\s-\s") {
        $targetNachname = "LAFP - $targetOrt"
    }

    $buroPlatz = "[Platznummer]"
    if ($CsvRow -and $CsvRow.Raum -and $CsvRow.Raum.ToString().Trim() -ne "") {
        $buroPlatz = $CsvRow.Raum.ToString().Trim()
    }

    $fort   = if ($CsvRow -and $CsvRow.Fortbildungsbereich) { Get-StrI $CsvRow.Fortbildungsbereich } else { "" }
    $anm    = if ($CsvRow -and $CsvRow.Anmerkungen)         { Get-StrI $CsvRow.Anmerkungen }         else { "" }
    $aInfo  = if ($fort -ne "" -and $anm -ne "") { "$fort - $anm" } else { "$fort$anm" }
    $verfStr = Get-VerfahrenI -Row $CsvRow

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

    $aeDez   = if ($CsvRow -and $ADObj -and $adDez -ne "") { "$oAbt - $adDez" } else { $oAbt }
    $aeDn    = "$LID $targetNachname"
    $aeDesc  = "$ouNum - $verfStr - $buroPlatz - [Verantwortlicher] - [TEL] | $aInfo"

    # ── GroupFilter-Markerspalte ─────────────────────────────────
    # Prueft ob einer der Gruppen des Users den Teilstring enthaelt.
    # Case-insensitiv. Ergebnis: "X" oder "".
    $filterMark = ""
    if ($GroupFilter -ne "") {
        foreach ($ug in $userGroups) {
            if ($ug -match [regex]::Escape($GroupFilter)) {
                $filterMark = "X"
                break
            }
        }
    }

    # ── Hashtable: KEINE inline-if, KEINE Ausdruecke ─────────────
    $Props = [ordered]@{
        "L-Kennung"                    = $LID
        "OU"                           = $ouName
        "OU_Pfad"                      = $ouPfad
        "andere_OU"                    = $andereOU
        "GELOESCHT"                    = $statusGeloescht
        "Benoetigt"                    = $isRequired
        "Geaendert"                    = ""
        "LAFP_LZPD_LKA"               = $lafpStr
        "NICHT_KONFORM"                = ""
        "AD_Vorname"                   = $adVn
        "AENDERN_Vorname"              = $LID
        "AD_Nachname"                  = $adNn
        "AENDERN_Nachname"             = $targetNachname
        "AD_DisplayName"               = $adDn
        "AENDERN_DisplayName"          = $aeDn
        "ORIGINAL_Standort"            = $oStand
        "ORIGINAL_Raum_Schulungskreis" = $oRaum
        "AD_Ort"                       = $adOrt
        "AENDERN_Ort"                  = $targetOrt
        "AD_Buero"                     = $adBue
        "AENDERN_Buero"                = $buroPlatz
        "ORIGINAL_Abteilung"           = $oAbt
        "AD_Dez"                       = $adDez
        "AENDERN_Dez"                  = $aeDez
        "AD_Description"               = $adDesc
        "AENDERN_Description"          = $aeDesc
        "AENDERN_OU"                   = $ouNum
        "AENDERN_Info"                 = $aInfo
        "LOESCHEN"                     = ""
    }

    # GroupFilter-Spalte nur hinzufuegen wenn Parameter gesetzt
    if ($GroupFilter -ne "") {
        $filterColName = "FILTER_$GroupFilter"
        $Props[$filterColName] = $filterMark
    }

    # ── Konformitaetspruefung ────────────────────────────────────
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
            $cs                   = $codes -join ","
            $Props."Geaendert"    = $cs
            $Props."NICHT_KONFORM" = $cs
        }
    }

    # ── Gruppen-Export ───────────────────────────────────────────
    if ($GroupMode -eq "Single") {
        # Alle Gruppen in einer Zelle, semikolonsepariert
        $Props["Gruppen"] = $userGroups -join ";"
    } else {
        # Eine Spalte pro Gruppe, Spaltenname = Gruppenname (kein GRP_-Prefix)
        foreach ($gn in $SortedGroups) {
            $inGrp = ($userGroups -contains $gn)
            $Props[$gn] = if ($inGrp) { "X" } else { "" }
        }
    }

    return [PSCustomObject]$Props
}

# ══════════════════════════════════════════════════════════════════
#  HELPER FUNCTIONS (DATA LOADING)
# ══════════════════════════════════════════════════════════════════
function Get-RequiredList {
    param([string]$Path)
    Write-Section "BEDARFSLISTE"
    $list = New-Object 'System.Collections.Generic.HashSet[string]'
    if ($Path -and (Test-Path $Path)) {
        foreach ($line in (Get-Content $Path -Encoding UTF8)) {
            if ($line -match "(L\d{6,8})") {
                $v = $Matches[1].ToUpper()
                [void]$list.Add($v)
            }
        }
        Write-Log "Bedarfsliste: $($list.Count) Eintraege" "OK"
    } else {
        Write-Log "Keine Bedarfsliste angegeben -> uebersprungen" "WARN"
    }
    return $list
}

function Get-MasterCsvData {
    param([string]$Path)
    Write-Section "MASTER-CSV laden"
    $t0 = $SW.Elapsed.TotalSeconds
    $data = @{}
    foreach ($row in (Import-Csv -Path $Path -Delimiter ';' -Encoding Default)) {
        if ($row."L-Kennung") {
            $k = $row."L-Kennung".ToString().Trim().ToUpper()
            if ($k) { $data[$k] = $row }
        }
    }
    Write-Log ("Master-CSV: {0} Zeilen  [{1:F2}s]" -f $data.Count, ($SW.Elapsed.TotalSeconds-$t0)) "OK"
    return $data
}

function Get-ADDiscovery {
    param(
        [string]$OUFilter,
        [string]$GroupFilter,
        [string[]]$ADProps
    )
    Write-Section "AD-DISCOVERY  (Subtree 81+82, alle Sub-OUs)"
    $t0           = $SW.Elapsed.TotalSeconds
    $ADCache      = @{}
    $UniqueGroups = New-Object 'System.Collections.Generic.HashSet[string]'

    # Beide Ziel-OUs suchen (81 UND 82), SearchScope Subtree erfasst alle Sub-OUs
    $TargetOUs = @(Get-ADOrganizationalUnit `
        -Filter "Name -eq '81' -or Name -eq '82'" `
        -ErrorAction SilentlyContinue)

    if ($TargetOUs.Count -eq 0) {
        Write-Log "WARNUNG: Keine OUs mit Name 81 oder 82 gefunden!" "WARN"
    }
    Write-Log "Ziel-OUs gefunden: $($TargetOUs.Count)  ($($TargetOUs.Name -join ', '))" "DBG"

    foreach ($ou in $TargetOUs) {
        Write-Log "Scanne OU (Subtree): $($ou.DistinguishedName)" "DBG"
        $t1 = $SW.Elapsed.TotalSeconds

        # SearchScope Subtree -> erfasst ALLE Sub-OUs, auch unbekannte
        $users = Get-ADUser `
            -Filter * `
            -SearchBase  $ou.DistinguishedName `
            -SearchScope Subtree `
            -Properties  $ADProps

        $userArr  = @($users)
        $rawCount = $userArr.Count
        Write-Log ("OU {0}: {1} User gefunden  [{2:F2}s]" -f $ou.Name, $rawCount, ($SW.Elapsed.TotalSeconds-$t1)) "DBG"

        foreach ($u in $userArr) {
            $sam = $u.SamAccountName.ToUpper()
            $dn  = $u.DistinguishedName

            # OUFilter: wenn gesetzt, nur User aus passenden OUs einlesen
            # -> reduziert ADCache und spart Verarbeitungszeit
            if ($OUFilter -ne "") {
                if ($dn -notlike "*$OUFilter*") {
                    Write-Log "OUFilter: $sam uebersprungen ($dn)" "DBG"
                    continue
                }
            }

            $ADCache[$sam] = $u

            foreach ($gDN in $u.MemberOf) {
                $gn = Get-GroupName -DN $gDN
                [void]$UniqueGroups.Add($gn)
            }
        }
        Write-Log ("OU {0}: {1} User in Cache aufgenommen (OUFilter={2})" -f `
            $ou.Name, $ADCache.Count, $(if($OUFilter -ne ""){"'$OUFilter'"}else{"aus"})) "DBG"
    }

    $SortedGroups = @($UniqueGroups | Sort-Object)
    Write-Log ("Discovery: {0} AD-Objekte  |  {1} unique Gruppen  [{2:F2}s]" -f `
        $ADCache.Count, $SortedGroups.Count, ($SW.Elapsed.TotalSeconds-$t0)) "PERF"

    $grpFilterCount = 0
    if ($GroupFilter -ne "") {
        $grpFilterCount = @($SortedGroups | Where-Object { $_ -like "*$GroupFilter*" }).Count
        Write-Log ("GroupFilter='$GroupFilter': $grpFilterCount Gruppen treffen zu -> Spalte FILTER_$GroupFilter") "INFO"
    }

    return @{
        ADCache        = $ADCache
        SortedGroups   = $SortedGroups
        GrpFilterCount = $grpFilterCount
    }
}

function Get-ProcessList {
    param(
        $ADCacheKeys,
        $MasterCsvKeys,
        [switch]$LafpOnly,
        [string]$NameFilter,
        [int]$TestCount
    )
    Write-Section "FILTER & ARBEITSLISTE"
    $t0 = $SW.Elapsed.TotalSeconds

    # Alle SAMs aus AD-Cache UND Master-CSV zusammenfuehren
    $AllSAMs = @(@($ADCacheKeys) + @($MasterCsvKeys) | Select-Object -Unique | Sort-Object)
    Write-Log "Gesamt unique SAMs (AD+CSV): $($AllSAMs.Count)" "DBG"

    # Filter anwenden (BEVOR parallele Verarbeitung -> spart Laufzeit)
    $ProcessList = $AllSAMs

    # Filter 1: LafpOnly -> nur L110* und L114*
    if ($LafpOnly) {
        $before = $ProcessList.Count
        $ProcessList = @($ProcessList | Where-Object { $_ -like "L110*" -or $_ -like "L114*" })
        Write-Log ("LafpOnly: {0} -> {1} SAMs  (-{2})" -f $before, $ProcessList.Count, ($before-$ProcessList.Count)) "WARN"
    }

    # Filter 2: NameFilter -> SAMAccountName-Wildcard
    if ($NameFilter -ne "") {
        $before = $ProcessList.Count
        $ProcessList = @($ProcessList | Where-Object { $_ -like $NameFilter })
        Write-Log ("NameFilter '{0}': {1} -> {2} SAMs  (-{3})" -f `
            $NameFilter, $before, $ProcessList.Count, ($before-$ProcessList.Count)) "WARN"
    }

    # Filter 3: TestCount
    if ($TestCount -gt 0) {
        Write-Log "TESTMODUS: erste $TestCount Eintraege" "WARN"
        $ProcessList = @($ProcessList | Select-Object -First $TestCount)
    }

    Write-Log ("Zu verarbeitende SAMs: {0}  [{1:F2}s]" -f $ProcessList.Count, ($SW.Elapsed.TotalSeconds-$t0)) "OK"

    return $ProcessList
}

function Invoke-ParallelProcessing {
    param(
        $ProcessList,
        [int]$MaxThreads,
        $CalcScriptBlock,
        $MasterCsvData,
        $ADCache,
        $RequiredList,
        $SortedGroups,
        [string]$GroupMode,
        [string]$GroupFilter,
        [switch]$DebugMode
    )
    Write-Section "VERARBEITUNG  ($($ProcessList.Count) Eintraege  |  $MaxThreads Threads)"
    $t0 = $SW.Elapsed.TotalSeconds

    $Pool = [runspacefactory]::CreateRunspacePool(
        1,
        $MaxThreads,
        [system.management.automation.runspaces.initialsessionstate]::CreateDefault(),
        $Host)
    $Pool.Open()
    Write-Log "Runspace-Pool geoeffnet ($MaxThreads Threads)" "DBG"

    # Alle Jobs aufbauen
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
        [void]$psi.AddArgument($GroupFilter)
        $psi.RunspacePool = $Pool
        [void]$Jobs.Add([PSCustomObject]@{
            SAM         = $LID
            Instance    = $psi
            AsyncResult = $psi.BeginInvoke()
        })
    }
    Write-Log "$($Jobs.Count) Jobs gestartet" "DBG"

    # Ergebnisse einsammeln
    $Results  = New-Object 'System.Collections.Generic.List[PSObject]'
    $done     = 0
    $errCount = 0
    $Total    = $Jobs.Count

    while ($Jobs.Count -gt 0) {
        $finished = @($Jobs | Where-Object { $_.AsyncResult.IsCompleted })
        $toRemove = New-Object 'System.Collections.Generic.List[PSObject]'

        foreach ($job in $finished) {
            $done++

            # Fehler-Stream loggen
            foreach ($e in $job.Instance.Streams.Error) {
                $errCount++
                Write-Log "Runspace $($job.SAM): $e" "ERR"
            }

            # Ergebnis einsammeln (PSDataCollection-sicher)
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
                $nk = if ($item) { $item."NICHT_KONFORM" } else { "?" }
                Write-Host ("  [{0,4}/{1}]  {2,-14}  nk:{3}" -f $done, $Total, $job.SAM, $nk) `
                    -ForegroundColor DarkGray
            }

            $job.Instance.Dispose()
            [void]$toRemove.Add($job)
        }
        foreach ($j in $toRemove) { [void]$Jobs.Remove($j) }

        # Progress alle 50 oder am Ende
        if ($done % 50 -eq 0 -or $Jobs.Count -eq 0) {
            $elapsed = $SW.Elapsed.TotalSeconds
            $rem     = if ($done -gt 0) { [math]::Round(($elapsed/$done)*($Total-$done)) } else { 0 }
            $pct     = [math]::Round(($done/$Total)*100)
            $fill    = [math]::Round(30*$pct/100)
            $bar     = ("#" * $fill) + ("." * (30-$fill))
            $eta     = if ($done -lt $Total) { "ETA ~${rem}s" } else { "fertig!   " }
            Write-Host ("`r  [{0}]  {1,4}/{2}  ({3,3}%)  {4}    " -f $bar,$done,$Total,$pct,$eta) `
                -NoNewline -ForegroundColor Cyan
            Write-Progress -Activity "Parallel-Verarbeitung" `
                -Status "$done / $Total  ($pct%)  ~${rem}s" `
                -PercentComplete $pct
        }

        if ($Jobs.Count -gt 0) { Start-Sleep -Milliseconds 30 }
    }

    Write-Host ""
    Write-Progress -Activity "Parallel-Verarbeitung" -Completed
    $Pool.Close(); $Pool.Dispose()

    $procSec    = [math]::Round($SW.Elapsed.TotalSeconds - $t0, 2)
    $msPerEntry = if ($Results.Count -gt 0) { [math]::Round($procSec/$Results.Count*1000,1) } else { 0 }

    if ($errCount -gt 0) {
        Write-Log "$errCount Runspace-Fehler. Details: $LogFile" "WARN"
    }
    Write-Log ("Verarbeitung: {0} Eintraege  {1}s  ({2}ms/Eintrag)  {3} Fehler" -f `
        $Results.Count, $procSec, $msPerEntry, $errCount) "PERF"

    return @{
        Results  = $Results
        ErrCount = $errCount
    }
}

function Export-ProcessingResults {
    param(
        $Results,
        [string]$GroupMode,
        [string]$GroupFilter,
        $SortedGroups,
        [string]$Version,
        [string]$ScriptDir
    )
    Write-Section "EXPORT"
    if ($Results.Count -eq 0) { Write-Log "Keine Ergebnisse -> Abbruch" "ERR"; return }

    # Alle verfuegbaren Spalten aus erstem Ergebnis-Objekt holen
    $allResultCols = @($Results[0].psobject.Properties.Name)

    # Tabelle 1: Gruppen-Uebersicht (kompakt)
    # Basis + OU + GroupFilter-Marker + Gruppen-Spalten
    $cols1Base = @("L-Kennung","OU","GELOESCHT","Benoetigt","Geaendert","LOESCHEN","AD_Vorname","AD_Nachname")

    # GroupFilter-Markerspalte hinzufuegen wenn gesetzt
    $filterColName = ""
    if ($GroupFilter -ne "") {
        $filterColName = "FILTER_$GroupFilter"
        $cols1Base += $filterColName
    }

    # Gruppen-Spalten (ohne GRP_-Prefix, korrekte Spaltennamen)
    if ($GroupMode -eq "Single") {
        $cols1 = $cols1Base + @("Gruppen")
    } else {
        $grpCols = @($SortedGroups)   # kein GRP_-Prefix mehr
        $cols1   = $cols1Base + $grpCols
    }

    # Tabelle 2: Alle Properties (fuer Apply-Script) ohne LOESCHEN
    $cols2 = @($allResultCols | Where-Object { $_ -ne "LOESCHEN" })

    Write-Log ("Tab1: {0} Spalten  |  Tab2: {1} Spalten" -f $cols1.Count, $cols2.Count) "DBG"
    if ($GroupMode -eq "Columns") {
        Write-Log "GroupMode=Columns: $($SortedGroups.Count) Gruppen = breite Tabelle" "WARN"
    }

    $Path1 = Join-Path $ScriptDir "L-Kennungen_Full_Analysis_v$Version.csv"
    $Path2 = Join-Path $ScriptDir "L-Kennungen_Properties_v$Version.csv"

    Export-FastCsv -Path $Path1 -ColNames $cols1 -Data $Results -Label "Tab1 Gruppen-Uebersicht"
    Export-FastCsv -Path $Path2 -ColNames $cols2 -Data $Results -Label "Tab2 Properties (Apply-Eingabe)"

    return @{ Path1 = $Path1; Path2 = $Path2 }
}

# ══════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════
function Invoke-Main {
    Show-Banner
    Import-Module ActiveDirectory -Verbose:$false

    Write-Log ("PS v{0}  PID={1}  Host={2}  Threads={3}  GroupMode={4}" -f `
        $PSVersionTable.PSVersion, $PID, $env:COMPUTERNAME, $MaxThreads, $GroupMode) "DBG"

    if (-not $CsvPath) {
        $CsvPath = (Read-Host "Pfad zur Master-CSV").Trim().Trim('"')
    }
    $CsvPath = $CsvPath.Trim().Trim('"')
    if (-not (Test-Path $CsvPath)) { Write-Log "CSV nicht gefunden: $CsvPath" "ERR"; return }

    # ── 1. BEDARFSLISTE ──────────────────────────────────────────
    $RequiredList = Get-RequiredList -Path $RequiredCsvPath

    # ── 2. MASTER-CSV ────────────────────────────────────────────
    $MasterCsvData = Get-MasterCsvData -Path $CsvPath

    # ── 3. AD-DISCOVERY (Subtree OU=81 und OU=82) ────────────────
    $discResult = Get-ADDiscovery -OUFilter $OUFilter -GroupFilter $GroupFilter -ADProps $AD_PROPS
    $ADCache        = $discResult.ADCache
    $SortedGroups   = $discResult.SortedGroups
    $grpFilterCount = $discResult.GrpFilterCount

    # ── 4. ARBEITSLISTE AUFBAUEN + FILTER ────────────────────────
    $ProcessList = Get-ProcessList `
        -ADCacheKeys   $ADCache.Keys `
        -MasterCsvKeys $MasterCsvData.Keys `
        -LafpOnly:$LafpOnly `
        -NameFilter    $NameFilter `
        -TestCount     $TestCount

    if ($ProcessList.Count -eq 0) {
        Write-Log "Keine SAMs nach Filterung -> Abbruch" "ERR"; return
    }

    # ── 5. PARALLELE VERARBEITUNG (immer aktiv) ──────────────────
    $procResult = Invoke-ParallelProcessing `
        -ProcessList      $ProcessList `
        -MaxThreads       $MaxThreads `
        -CalcScriptBlock  $CalcScriptBlock `
        -MasterCsvData    $MasterCsvData `
        -ADCache          $ADCache `
        -RequiredList     $RequiredList `
        -SortedGroups     $SortedGroups `
        -GroupMode        $GroupMode `
        -GroupFilter      $GroupFilter `
        -DebugMode:$DebugMode

    $Results  = $procResult.Results
    $errCount = $procResult.ErrCount

    # ── 6. EXPORT ────────────────────────────────────────────────
    $exportResult = Export-ProcessingResults `
        -Results      $Results `
        -GroupMode    $GroupMode `
        -GroupFilter  $GroupFilter `
        -SortedGroups $SortedGroups `
        -Version      $Version `
        -ScriptDir    $ScriptDir

    $Path1 = $exportResult.Path1
    $Path2 = $exportResult.Path2

    if (-not $Path1 -or -not $Path2) {
        return
    }

    # ── ABSCHLUSS ────────────────────────────────────────────────
    Show-Summary -ResultsCount $Results.Count -SortedGroupsCount $SortedGroups.Count `
        -GroupMode $GroupMode -GroupFilter $GroupFilter -GrpFilterCount $grpFilterCount `
        -OUFilter $OUFilter -NameFilter $NameFilter -LafpOnly:$LafpOnly `
        -ErrCount $errCount -Path1 $Path1 -Path2 $Path2
}

function Show-Summary {
    param(
        [int]$ResultsCount,
        [int]$SortedGroupsCount,
        [string]$GroupMode,
        [string]$GroupFilter,
        [int]$GrpFilterCount,
        [string]$OUFilter,
        [string]$NameFilter,
        [switch]$LafpOnly,
        [int]$ErrCount,
        [string]$Path1,
        [string]$Path2
    )
    $SW.Stop()
    $dur = $SW.Elapsed.ToString('mm\:ss\.ff')
    $mb  = [math]::Round([System.GC]::GetTotalMemory($false) / 1MB, 1)

    $grpFilterInfo = if ($GroupFilter -ne "") { "$($GrpFilterCount)x  -> Spalte FILTER_$GroupFilter" } else { "kein" }
    $filterInfo    = @()
    if ($OUFilter  -ne "") { $filterInfo += "OU=$OUFilter"   }
    if ($NameFilter -ne "") { $filterInfo += "Name=$NameFilter" }
    if ($LafpOnly)          { $filterInfo += "LafpOnly" }
    $filterSummary = if ($filterInfo.Count -gt 0) { $filterInfo -join "  |  " } else { "keine" }

    Write-Host ""
    Write-Host "  +----------------------------------------------------------+" -ForegroundColor Green
    Write-Host "  |  ABGESCHLOSSEN                                           |" -ForegroundColor Green
    Write-Host ("  |  Datensaetze : {0,-43}|" -f $ResultsCount)                  -ForegroundColor Green
    Write-Host ("  |  Gruppen     : {0,-43}|" -f "$($SortedGroupsCount)  (Mode: $GroupMode)") -ForegroundColor Green
    Write-Host ("  |  GrpFilter   : {0,-43}|" -f $grpFilterInfo)                 -ForegroundColor Green
    Write-Host ("  |  Filter      : {0,-43}|" -f $filterSummary)                 -ForegroundColor Green
    Write-Host ("  |  Fehler      : {0,-43}|" -f $ErrCount)                      -ForegroundColor $(if($ErrCount -gt 0){"Red"}else{"Green"})
    Write-Host ("  |  Dauer       : {0,-43}|" -f $dur)                           -ForegroundColor Green
    Write-Host ("  |  RAM         : {0,-43}|" -f "${mb} MB")                     -ForegroundColor Green
    Write-Host "  |                                                          |" -ForegroundColor DarkGreen
    Write-Host ("  |  Tab1 : {0,-50}|" -f (Split-Path $Path1 -Leaf))             -ForegroundColor DarkGreen
    Write-Host ("  |  Tab2 : {0,-50}|" -f (Split-Path $Path2 -Leaf))             -ForegroundColor DarkGreen
    Write-Host ("  |  Log  : {0,-50}|" -f (Split-Path $LogFile -Leaf))           -ForegroundColor DarkGreen
    Write-Host "  +----------------------------------------------------------+" -ForegroundColor Green
    Write-Host ""

    Write-Log ("Fertig. Dauer={0}  RAM={1}MB  Zeilen={2}  Gruppen={3}  Fehler={4}" -f `
        $dur, $mb, $ResultsCount, $SortedGroupsCount, $ErrCount) "PERF"
}

Invoke-Main
