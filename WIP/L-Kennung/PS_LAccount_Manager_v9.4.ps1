# FileName: PS_LAccount_Manager.ps1
# Version:  9.4
# Beschreibung: AD-Sync Tool. Vollstaendig gefixt und getestet.
# Author:   PS-Coding
#
# ALLE FIXES:
#   [FIX-1] 'Start-Process' -> 'Invoke-Main' (Namenskonflikt mit eingebautem Cmdlet)
#   [FIX-2] PSDataCollection-Unrolling: IsOpen/Count-Bug in CSV beseitigt
#   [FIX-3] Doppelte $Anmerkung-Zuweisung beseitigt
#   [FIX-4] Collection-Modifikation waehrend foreach-Iteration gesichert
#   [FIX-5] 'Geaendert'-Spalte wird jetzt korrekt befuellt
#   [FIX-6] Bounds-Check fuer $Results[0]
#   [FIX-7] Keine inline-if-Ausdruecke in Hashtable-Werten im ScriptBlock
#   [FIX-8] Komma in '-replace x,y' innerhalb .Add() als 2. Argument geparst
#           -> ALLE .Add()-Aufrufe nutzen vorberechnete Variablen, KEIN Inline-Ausdruck
#   [FIX-9] Mehrzeilige -Properties ohne Backtick koennen als neue Anweisung geparst
#           werden -> alle -Properties auf eine Zeile oder mit explizitem Backtick

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)] [string]$CsvPath,
    [Parameter(Mandatory=$false)] [string]$RequiredCsvPath,
    [Parameter(Mandatory=$false)] [switch]$SearchGlobal,
    [Parameter(Mandatory=$false)] [int]$TestCount  = 0,
    [Parameter(Mandatory=$false)] [int]$MaxThreads = 8,
    [Parameter(Mandatory=$false)] [switch]$DebugMode
)

Set-StrictMode -Off
$Version               = "9.4"
$ErrorActionPreference = 'Stop'
$ScriptDir             = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD }
$Timestamp             = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile               = Join-Path $ScriptDir "PS_LAccountManager_$Timestamp.log"
$Stopwatch             = [System.Diagnostics.Stopwatch]::StartNew()

# Alle AD-Properties als Konstante - einmal definieren, ueberall verwenden
$AD_PROPS = @("DisplayName","Description","GivenName","Surname","l","physicalDeliveryOfficeName","department","info","MemberOf")

function Show-Header {
    Clear-Host
    Write-Host @"
##############################################################
#    AD COMPLIANCE & SYNC MANAGER  v$Version  (Multi-Thread)    #
#    Status: Live-Reporting Aktiv | Threads: $MaxThreads             #
##############################################################
"@ -ForegroundColor Cyan
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","SUCCESS","DEBUG")] $Level = "INFO"
    )
    if ($Level -eq "DEBUG" -and -not $DebugMode) { return }
    $Time = Get-Date -Format "HH:mm:ss"
    $Color = switch ($Level) {
        "SUCCESS" { "Green"   }
        "ERROR"   { "Red"     }
        "WARN"    { "Yellow"  }
        "DEBUG"   { "Magenta" }
        default   { "Gray"    }
    }
    $Line = "[$Time] [$Level] $Message"
    Write-Host $Line -ForegroundColor $Color
    $Line | Out-File -FilePath $LogFile -Append
}

# Hilfsfunktion: Gruppenname aus DN extrahieren - OHNE Komma-Problem in Methodenaufruf
function Get-GroupName {
    param([string]$DN)
    $part = ($DN -split ',')[0]
    $name = $part -replace 'CN=', ''
    return $name
}

function Get-RequiredList {
    param([string]$RequiredCsvPath)
    # ----------------------------------------------------------------
    # 1. BEDARFSLISTE laden
    # ----------------------------------------------------------------
    $RequiredList = New-Object 'System.Collections.Generic.HashSet[string]'
    if ($RequiredCsvPath -and (Test-Path $RequiredCsvPath)) {
        Write-Log "Lade Bedarfsliste..." "DEBUG"
        foreach ($line in (Get-Content $RequiredCsvPath -Encoding UTF8)) {
            if ($line -match "(L\d{6,8})") {
                $val = $Matches[1].ToUpper()
                [void]$RequiredList.Add($val)          # FIX-8: nur einfache Variable
            }
        }
        Write-Log "Bedarfsliste: $($RequiredList.Count) Eintraege." "SUCCESS"
    }

    return $RequiredList
}
function Get-MasterCsvData {
    param([string]$CsvPath)
    # ----------------------------------------------------------------
    # 2. MASTER-CSV laden
    # ----------------------------------------------------------------
    Write-Log "Lese Master-CSV..." "DEBUG"
    $MasterCsvData = @{}
    foreach ($row in (Import-Csv -Path $CsvPath -Delimiter ';' -Encoding Default)) {
        if ($row."L-Kennung") {
            $k = $row."L-Kennung".ToString().Trim().ToUpper()
            if ($k) { $MasterCsvData[$k] = $row }
        }
    }
    Write-Log "Master-CSV: $($MasterCsvData.Count) Zeilen." "SUCCESS"

    return $MasterCsvData
}
function Invoke-ADDiscovery {
    param([hashtable]$MasterCsvData, [switch]$SearchGlobal, [int]$TestCount)
    # ----------------------------------------------------------------
    # 3. AD-DISCOVERY
    # ----------------------------------------------------------------
    Write-Log "Schritt 1: AD-Discovery (OUs 81/82)..." "INFO"
    $ADCache      = @{}
    $UniqueGroups = New-Object 'System.Collections.Generic.HashSet[string]'

    # FIX-9: -Properties alle auf einer Zeile, kein mehrzeiliger Umbruch ohne Backtick
    $TargetOUs = Get-ADOrganizationalUnit -Filter "Name -eq '81' -or Name -eq '82'" -ErrorAction SilentlyContinue

    foreach ($ou in $TargetOUs) {
        Write-Log "Scanne OU: $($ou.DistinguishedName)" "DEBUG"

        $users = Get-ADUser -Filter * -SearchBase $ou.DistinguishedName -Properties $AD_PROPS

        foreach ($u in $users) {
            $sam = $u.SamAccountName.ToUpper()
            $ADCache[$sam] = $u

            foreach ($groupDN in $u.MemberOf) {
                # FIX-8: Get-GroupName extrahiert den Namen sauber,
                # kein Komma-Problem mehr im .Add()-Aufruf
                $groupName = Get-GroupName -DN $groupDN
                [void]$UniqueGroups.Add($groupName)
            }
        }
    }

    if ($SearchGlobal) {
        Write-Log "Schritt 2: Globaler AD-Check..." "INFO"

        $GlobalUsers = Get-ADUser -Filter "SamAccountName -like 'L*'" -Properties $AD_PROPS

        foreach ($gu in $GlobalUsers) {
            $sam = $gu.SamAccountName.ToUpper()
            if (-not $ADCache.ContainsKey($sam)) {
                $ADCache[$sam] = $gu
                foreach ($groupDN in $gu.MemberOf) {
                    $groupName = Get-GroupName -DN $groupDN
                    [void]$UniqueGroups.Add($groupName)
                }
            }
        }
    }

    $SortedGroups  = @($UniqueGroups | Sort-Object)
    $AllSAMs       = @($ADCache.Keys) + @($MasterCsvData.Keys) | Select-Object -Unique | Sort-Object
    $ProcessList   = if ($TestCount -gt 0) {
        Write-Log "TESTMODUS: Begrenze auf $TestCount Eintraege." "WARN"
        @($AllSAMs | Select-Object -First $TestCount)
    } else {
        @($AllSAMs)
    }

    Write-Log ("Discovery: {0}s. Verarbeite {1} Eintraege..." -f `
        $Stopwatch.Elapsed.TotalSeconds.ToString('F2'), $ProcessList.Count) "SUCCESS"

    return [PSCustomObject]@{
        ADCache = $ADCache
        SortedGroups = $SortedGroups
        ProcessList = $ProcessList
    }
}
function Invoke-ProcessingJobs {
    param([array]$ProcessList, [hashtable]$MasterCsvData, [hashtable]$ADCache, [System.Collections.Generic.HashSet[string]]$RequiredList, [array]$SortedGroups, [int]$MaxThreads)
    # ----------------------------------------------------------------
    # 4. RUNSPACE-POOL + SCRIPTBLOCK
    # ----------------------------------------------------------------
    $Pool = [runspacefactory]::CreateRunspacePool(
        1,
        $MaxThreads,
        [system.management.automation.runspaces.initialsessionstate]::CreateDefault(),
        $Host
    )
    $Pool.Open()

    $ScriptBlock = {
        param(
            [string]$LID,
            $CsvRow,
            $ADObj,
            $RequiredList,
            $SortedGroups
        )

        # --- Hilfsfunktionen ---
        function Get-StrVal {
            param($v)
            if ($null -eq $v) { return "" }
            return $v.ToString().Trim()
        }

        function Test-Diff {
            param($a, $b)
            return ((Get-StrVal $a) -ine (Get-StrVal $b))
        }

        function Get-Verfahren {
            param($Row)
            if ($null -eq $Row) { return "[alle_Verfahren]" }
            $cols   = @("Viva","Findus","MobiApps","AccVisio","Verfahren5","Verfahren6","Verfahren7","Verfahren8","Verfahren9","Verfahren10")
            $active = New-Object 'System.Collections.Generic.List[string]'
            foreach ($c in $cols) {
                if ($Row.$c -and $Row.$c.ToString().Trim().ToLower() -eq "x") {
                    $active.Add($c)
                }
            }
            if ($active.Count -gt 0) { return ($active -join " - ") }
            return "[alle_Verfahren]"
        }

        # --- Basiswerte aus AD ---
        $andereOU        = ""
        $statusGeloescht = ""
        $userGroups      = @()

        if ($ADObj) {
            $dn = $ADObj.DistinguishedName
            if ($dn -notmatch "OU=81" -and $dn -notmatch "OU=82") {
                $andereOU = $dn -replace '^CN=.*?,', ''
            }
            if ($ADObj.MemberOf) {
                $userGroups = @(
                    $ADObj.MemberOf | ForEach-Object {
                        $part = ($_ -split ',')[0]
                        $part -replace 'CN=', ''
                    }
                )
            }
        } else {
            $statusGeloescht = "XXX"
        }

        # --- Alle Werte vorberechnen (FIX-7: kein inline-if im Hashtable) ---

        $isRequired = ""
        if ($RequiredList.Contains($LID.ToUpper())) { $isRequired = "ZZZ" }

        $isLafp  = ($LID -like "L110*") -or ($LID -like "L114*")
        $lafpStr = ""
        $ouNum   = "[OU]"
        if ($isLafp) { $lafpStr = "LLL"; $ouNum = "26" }

        # Zielort
        $targetOrt = ""
        if ($CsvRow -and $CsvRow.Standort -and ($CsvRow.Standort.ToString().Trim() -ne "")) {
            $targetOrt = $CsvRow.Standort.ToString().Trim()
        } elseif ($ADObj -and $ADObj.l) {
            $targetOrt = Get-StrVal $ADObj.l
        }

        # Nachname
        $targetNachname = $targetOrt
        if ($isLafp -and ($targetOrt -notmatch "^LAFP\s-\s")) {
            $targetNachname = "LAFP - $targetOrt"
        }

        # Bueroraum
        $buroPlatz = "[Platznummer]"
        if ($CsvRow -and $CsvRow.Raum -and ($CsvRow.Raum.ToString().Trim() -ne "")) {
            $buroPlatz = $CsvRow.Raum.ToString().Trim()
        }

        # Fortbildung + Anmerkung
        $fortbildung = ""
        if ($CsvRow -and $CsvRow.Fortbildungsbereich) {
            $fortbildung = Get-StrVal $CsvRow.Fortbildungsbereich
        }
        $anmerkung = ""
        if ($CsvRow -and $CsvRow.Anmerkungen) {
            $anmerkung = Get-StrVal $CsvRow.Anmerkungen
        }
        $aenderInfo = "$fortbildung$anmerkung"
        if ($fortbildung -ne "" -and $anmerkung -ne "") {
            $aenderInfo = "$fortbildung - $anmerkung"
        }

        # Verfahren
        $verfStr = Get-Verfahren -Row $CsvRow

        # AD-Felder
        $adVorname  = ""
        $adNachname = ""
        $adDisplay  = ""
        $adOrt      = ""
        $adBuero    = ""
        $adDez      = ""
        $adDesc     = ""
        $adInfo     = ""
        if ($ADObj) {
            $adVorname  = Get-StrVal $ADObj.GivenName
            $adNachname = Get-StrVal $ADObj.Surname
            $adDisplay  = Get-StrVal $ADObj.DisplayName
            $adOrt      = Get-StrVal $ADObj.l
            $adBuero    = Get-StrVal $ADObj.physicalDeliveryOfficeName
            $adDez      = Get-StrVal $ADObj.department
            $adDesc     = Get-StrVal $ADObj.Description
            $adInfo     = Get-StrVal $ADObj.info
        }

        # CSV-Felder
        $origStandort = ""
        $origRaum     = ""
        $origAbt      = ""
        if ($CsvRow) {
            $origStandort = Get-StrVal $CsvRow.Standort
            $origRaum     = Get-StrVal $CsvRow.Raum
            $origAbt      = Get-StrVal $CsvRow.Abteilung
        }

        # Berechnete Zielwerte
        $aendernDez  = $origAbt
        if ($CsvRow -and $ADObj -and $adDez -ne "") {
            $aendernDez = "$origAbt - $adDez"
        }
        $displayName = "$LID $targetNachname"
        $aendernDesc = "$ouNum - $verfStr - $buroPlatz - [Verantwortlicher] - [TEL] | $aenderInfo"

        # --- Hashtable: nur Variablen, KEIN inline-if (FIX-7) ---
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

        # --- Konformitaetspruefung ---
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
                $codeStr                 = $codes -join ","
                $Props."Geaendert"       = $codeStr
                $Props."NICHT_KONFORM"   = $codeStr
            }
        }

        # --- Gruppenspalten ---
        foreach ($gn in $SortedGroups) {
            $inGroup = ($userGroups -contains $gn)
            $Props["GRP_$gn"] = if ($inGroup) { "X" } else { "" }
        }

        return [PSCustomObject]$Props
    }

    # --- Jobs starten ---
    $Jobs = New-Object 'System.Collections.Generic.List[PSObject]'
    foreach ($LID in $ProcessList) {
        $psi = [powershell]::Create()
        [void]$psi.AddScript($ScriptBlock)
        [void]$psi.AddArgument($LID)
        [void]$psi.AddArgument($MasterCsvData[$LID])
        [void]$psi.AddArgument($ADCache[$LID])
        [void]$psi.AddArgument($RequiredList)
        [void]$psi.AddArgument($SortedGroups)
        $psi.RunspacePool = $Pool

        $entry = New-Object PSObject -Property @{
            SAM         = $LID
            Instance    = $psi
            AsyncResult = $psi.BeginInvoke()
        }
        [void]$Jobs.Add($entry)
    }

    # --- Jobs einsammeln ---
    $Results  = New-Object 'System.Collections.Generic.List[PSObject]'
    $Count    = 0
    $ErrCount = 0
    $Total    = $Jobs.Count

    while ($Jobs.Count -gt 0) {
        $Finished = @($Jobs | Where-Object { $_.AsyncResult.IsCompleted })
        $ToRemove = New-Object 'System.Collections.Generic.List[PSObject]'

        foreach ($Job in $Finished) {
            $Count++

            # Fehler-Stream loggen (OPT-1)
            foreach ($e in $Job.Instance.Streams.Error) {
                $ErrCount++
                $em = "[$( Get-Date -Format 'HH:mm:ss')] [ERROR] Runspace $($Job.SAM): $e"
                Write-Host $em -ForegroundColor Red
                $em | Out-File $LogFile -Append
            }

            # FIX-2: Robustes Unrolling der PSDataCollection
            $rawOut = $Job.Instance.EndInvoke($Job.AsyncResult)
            foreach ($item in $rawOut) {
                if ($null -ne $item -and $null -ne $item.PSObject) {
                    # Nur echte Datenobjekte, keine Collection-Metadaten
                    $typeName = $item.GetType().FullName
                    if ($typeName -notlike "*PSDataCollection*" -and `
                        $typeName -notlike "*Collection*") {
                        [void]$Results.Add($item)
                    }
                }
            }

            Write-Host "[$Count/$Total] OK: $($Job.SAM)" -ForegroundColor Gray
            Write-Progress -Activity "L-Kennungen verarbeiten" `
                -Status "$Count / $Total" `
                -PercentComplete ([math]::Round(($Count / $Total) * 100))

            $Job.Instance.Dispose()
            [void]$ToRemove.Add($Job)
        }

        # FIX-4: Entfernen erst nach dem Loop
        foreach ($j in $ToRemove) { [void]$Jobs.Remove($j) }
        if ($Jobs.Count -gt 0) { Start-Sleep -Milliseconds 50 }
    }

    Write-Progress -Activity "L-Kennungen verarbeiten" -Completed
    $Pool.Close()
    $Pool.Dispose()

    return [PSCustomObject]@{
        Results = $Results
        ErrCount = $ErrCount
    }
}
function Export-Results {
    param([System.Collections.Generic.List[PSObject]]$Results, [int]$ErrCount, [array]$SortedGroups)
    # ----------------------------------------------------------------
    # 5. EXPORT
    # ----------------------------------------------------------------
    # FIX-6: Bounds-Check vor $Results[0]
    if ($Results.Count -eq 0) {
        Write-Log "FEHLER: Keine Ergebnisse. AD-Verbindung und CSV pruefen." "ERROR"
        return
    }
    if ($ErrCount -gt 0) {
        Write-Log "$ErrCount Runspace-Fehler aufgetreten. Details: $LogFile" "WARN"
    }

    Write-Log "Erzeuge Ausgabedateien ($($Results.Count) Datensaetze)..." "INFO"

    # Tabelle 1: Gruppen-Uebersicht
    $Path1    = Join-Path $ScriptDir "L-Kennungen_Full_Analysis_v$Version.csv"
    $cols1    = @("L-Kennung","andere_OU","GELOESCHT","Benoetigt","Geaendert","LOESCHEN","AD_Vorname","AD_Nachname")
    $grpCols  = @($SortedGroups | ForEach-Object { "GRP_$_" })
    $Results | Select-Object ($cols1 + $grpCols) |
        Export-Csv -Path $Path1 -Delimiter ';' -NoTypeInformation -Encoding UTF8

    # Tabelle 2: Properties (ohne LOESCHEN-Spalte, mit korrektem CSV-Escaping)
    $Path2   = Join-Path $ScriptDir "L-Kennungen_Properties_v$Version.csv"
    $allCols = @($Results[0].psobject.Properties.Name | Where-Object { $_ -ne "LOESCHEN" })
    $lines   = New-Object 'System.Collections.Generic.List[string]'
    [void]$lines.Add($allCols -join ';')

    foreach ($r in $Results) {
        $rowArr = foreach ($col in $allCols) {
            $val = $r.$col
            $str = if ($null -ne $val) { $val.ToString() } else { "" }
            # RFC-4180 Escaping: Semikolon oder Anfuehrungszeichen -> quoten
            if ($str -match '[;\r\n"]') {
                '"' + $str.Replace('"', '""') + '"'
            } else {
                $str
            }
        }
        [void]$lines.Add($rowArr -join ';')
    }

    $lines | Out-File -FilePath $Path2 -Encoding UTF8 -Force

    $Stopwatch.Stop()

    Write-Host ""
    Write-Host "====================================================" -ForegroundColor Cyan
    Write-Host " FERTIG" -ForegroundColor Yellow
    Write-Host " Version : $Version"
    Write-Host " Datum   : $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')"
    Write-Host " Zeilen  : $($Results.Count)"
    Write-Host " Fehler  : $ErrCount"
    Write-Host " Dauer   : $($Stopwatch.Elapsed.TotalSeconds.ToString('F2'))s"
    Write-Host " Tab 1   : $Path1"
    Write-Host " Tab 2   : $Path2"
    Write-Host " Log     : $LogFile"
    Write-Host "====================================================" -ForegroundColor Cyan
}
function Invoke-Main {
    Show-Header
    Import-Module ActiveDirectory

    if (-not $CsvPath) { $CsvPath = Read-Host "Pfad zur Master-CSV eingeben" }
    $CsvPath = $CsvPath.Trim().Trim('"')
    if (-not (Test-Path $CsvPath)) {
        Write-Log "Master-CSV nicht gefunden: '$CsvPath'" "ERROR"
        return
    }

    $RequiredList = Get-RequiredList -RequiredCsvPath $RequiredCsvPath
    $MasterCsvData = Get-MasterCsvData -CsvPath $CsvPath

    $DiscoveryRes = Invoke-ADDiscovery -MasterCsvData $MasterCsvData -SearchGlobal $SearchGlobal -TestCount $TestCount
    $ADCache = $DiscoveryRes.ADCache
    $SortedGroups = $DiscoveryRes.SortedGroups
    $ProcessList = $DiscoveryRes.ProcessList

    $JobRes = Invoke-ProcessingJobs -ProcessList $ProcessList -MasterCsvData $MasterCsvData -ADCache $ADCache -RequiredList $RequiredList -SortedGroups $SortedGroups -MaxThreads $MaxThreads
    $Results = $JobRes.Results
    $ErrCount = $JobRes.ErrCount

    Export-Results -Results $Results -ErrCount $ErrCount -SortedGroups $SortedGroups
}

Invoke-Main
