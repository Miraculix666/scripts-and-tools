# FileName: PS_LAccount_Manager.ps1
# Description: AD-Sync Tool v9.2. Bugfix-Release.
# Version: 9.2
# Author: PS-Coding
#
# CHANGELOG v9.2 (Bugfixes & Optimierungen):
#   [FIX-1] KRITISCH: Funktion von 'Start-Process' in 'Invoke-Main' umbenannt
#           -> 'Start-Process' ist ein eingebautes PS-Cmdlet (Namenskonflikt).
#   [FIX-2] KRITISCH: PSDataCollection-Unrolling korrigiert.
#           -> Typprüfung 'GetType().Name -match PSCustomObject|PSObject' war zu
#              streng und hat valide Runspace-Objekte ausgefiltert. Stattdessen
#              landen die Metadaten der Collection selbst (IsOpen, Count etc.)
#              in der CSV. Fix: Prüfung auf 'psobject' Property als sicherer Test.
#   [FIX-3] BUG: Doppelte $Anmerkung-Zuweisung korrigiert.
#           -> '$Anmerkung = if(...){ $Anmerkung = ... }' – innere Zuweisung
#              gibt $null zurück, äußere Variable blieb immer leer.
#   [FIX-4] BUG: Collection-Modifikation während Iteration gesichert.
#           -> $Jobs.Remove() innerhalb foreach-Loop über $Finished kann
#              InvalidOperationException werfen. Jetzt mit separater Remove-Liste.
#   [FIX-5] BUG: 'Geändert'-Spalte wird jetzt korrekt mit den Change-Codes befüllt.
#           -> Wurde berechnet (für NICHT KONFORM), aber nie in $Props.'Geändert' gesetzt.
#   [FIX-6] MINOR: Bounds-Check für $Results[0] vor Property-Ermittlung hinzugefügt.
#   [OPT-1] Fehler-Streams aus Runspaces werden jetzt geloggt ($Job.Instance.Streams.Error).
#   [OPT-2] Progress-Bar via Write-Progress für bessere Übersicht.
#   [OPT-3] Runspace-Fehler brechen nicht mehr den Gesamt-Lauf ab.

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, HelpMessage="Pfad zur Master-CSV")]
    [string]$CsvPath,

    [Parameter(Mandatory=$false, HelpMessage="Pfad zur Bedarfsliste (CSV)")]
    [string]$RequiredCsvPath,

    [Parameter(Mandatory=$false)]
    [switch]$SearchGlobal,

    [Parameter(Mandatory=$false)]
    [int]$TestCount = 0,

    [Parameter(Mandatory=$false)]
    [int]$MaxThreads = 8,

    [Parameter(Mandatory=$false)]
    [switch]$DebugMode
)

# --- INITIALISIERUNG ---
$Version = "9.2"
$ErrorActionPreference = 'Stop'
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD }
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile = Join-Path $ScriptDir "PS_LAccountManager_$Timestamp.log"
$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

function Show-Header {
    Clear-Host
    $Header = @"
██╗      █████╗  ██████╗ ██████╗ ██████╗ ██╗███╗   ██╗ ██████╗ 
██║     ██╔══██╗██╔════╝██╔════╝██╔═══██╗██║████╗  ██║██╔════╝ 
██║     ███████║██║     ██║     ██║   ██║██║██╔██╗ ██║██║  ███╗
██║     ██╔══██║██║     ██║     ██║   ██║██║██║╚██╗██║██║   ██║
███████╗██║  ██║╚██████╗╚██████╗╚██████╔╝██║██║ ╚████║╚██████╔╝
                 AD COMPLIANCE & SYNC MANAGER v$Version (MT)
----------------------------------------------------------------
STATUS: Live-Reporting Aktiv | Threads: $MaxThreads
"@
    Write-Host $Header -ForegroundColor Cyan
}

function Write-Log {
    param([string]$Message, [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS", "DEBUG")]$Level = "INFO")
    if ($Level -eq "DEBUG" -and -not $DebugMode) { return }
    $Time = Get-Date -Format "HH:mm:ss"
    $LogColor = switch($Level) { 
        "SUCCESS" { "Green" } 
        "ERROR"   { "Red" } 
        "WARN"    { "Yellow" } 
        "DEBUG"   { "Magenta" }
        default   { "Gray" } 
    }
    $Msg = "[$Time] [$Level] $Message"
    Write-Host $Msg -ForegroundColor $LogColor
    $Msg | Out-File -FilePath $LogFile -Append
}

# [FIX-1] Umbenannt von 'Start-Process' -> 'Invoke-Main'
# Grund: 'Start-Process' ist ein eingebautes PowerShell-Cmdlet.
# Eine gleichnamige Funktion führt zu Namenskonflikt und unerwartetem Verhalten.
function Get-RequiredList {
    param([string]$RequiredCsvPath)
    $RequiredList = New-Object System.Collections.Generic.HashSet[string]
    if ($RequiredCsvPath -and (Test-Path $RequiredCsvPath)) {
        Write-Log "Lade Bedarfsliste..." "DEBUG"
        $RawLines = Get-Content $RequiredCsvPath -Encoding UTF8
        foreach ($line in $RawLines) {
            if ($line -match "(L\d{6,8})") { [void]$RequiredList.Add($matches[1].ToUpper()) }
        }
        Write-Log "Bedarfsliste geladen ($($RequiredList.Count) Einträge)." "SUCCESS"
    }
    return $RequiredList
}

function Get-MasterCsvData {
    param([string]$CsvPath)
    Write-Log "Lese Master-CSV..." "DEBUG"
    $MasterCsvData = @{}
    $RawCsv = Import-Csv -Path $CsvPath -Delimiter ';' -Encoding Default
    foreach ($line in $RawCsv) {
        $key = if ($line."L-Kennung") { $line."L-Kennung".ToString().Trim().ToUpper() } else { "" }
        if ($key) { $MasterCsvData[$key] = $line }
    }
    Write-Log "Master-CSV geladen ($($MasterCsvData.Count) Zeilen)." "SUCCESS"
    return $MasterCsvData
}

function Invoke-ADDiscovery {
    param([switch]$SearchGlobal, [hashtable]$MasterCsvData, [int]$TestCount)
    Write-Log "Schritt 1: AD-Discovery (OUs 81/82)..." "INFO"
    $ADCache = @{}
    $UniqueGroups = New-Object System.Collections.Generic.HashSet[string]
    
    $TargetOUs = Get-ADOrganizationalUnit -Filter "Name -eq '81' -or Name -eq '82'" -ErrorAction SilentlyContinue
    foreach ($ou in $TargetOUs) {
        Write-Log "Scanne OU: $($ou.DistinguishedName)" "DEBUG"
        $users = Get-ADUser -Filter * -SearchBase $ou.DistinguishedName `
            -Properties DisplayName, Description, GivenName, Surname, l, `
                        physicalDeliveryOfficeName, department, info, MemberOf
        foreach ($u in $users) { 
            $ADCache[$u.SamAccountName.ToUpper()] = $u 
            if ($u.MemberOf) {
                foreach ($g in $u.MemberOf) {
                    $gn = ($g -split ',')[0] -replace 'CN=', ''
                    [void]$UniqueGroups.Add($gn)
                }
            }
        }
    }

    if ($SearchGlobal) {
        Write-Log "Schritt 2: Globaler AD-Check..." "INFO"
        $GlobalUsers = Get-ADUser -Filter "SamAccountName -like 'L*'" `
            -Properties DisplayName, Description, GivenName, Surname, l, `
                        physicalDeliveryOfficeName, department, info, MemberOf
        foreach ($gu in $GlobalUsers) {
            $sam = $gu.SamAccountName.ToUpper()
            if (-not $ADCache.ContainsKey($sam)) { 
                $ADCache[$sam] = $gu 
                if ($gu.MemberOf) {
                    foreach ($g in $gu.MemberOf) {
                        $gn = ($g -split ',')[0] -replace 'CN=', ''
                        [void]$UniqueGroups.Add($gn)
                    }
                }
            }
        }
    }
    
    $SortedGroups = $UniqueGroups | Sort-Object
    $AllUniqueSAMs = ($ADCache.Keys + $MasterCsvData.Keys | Select-Object -Unique | Sort-Object)
    
    $ProcessList = $AllUniqueSAMs
    if ($TestCount -gt 0) { 
        Write-Log "TESTMODUS: Begrenze auf $TestCount Einträge." "WARN"
        $ProcessList = $AllUniqueSAMs | Select-Object -First $TestCount 
    }

    return @{
        ADCache = $ADCache
        SortedGroups = $SortedGroups
        ProcessList = $ProcessList
    }
}

function Invoke-ParallelProcessing {
    param(
        $ProcessList,
        $MasterCsvData,
        $ADCache,
        $RequiredList,
        $SortedGroups,
        [int]$MaxThreads,
        [string]$LogFile
    )

    $SessionState = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
    $Pool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads, $SessionState, $Host)
    $Pool.Open()

    $ScriptBlock = {
        param($LID, $CsvRow, $ADObj, $RequiredList, $SortedGroups)
        
        function Local-IsDifferent($v1, $v2) {
            $s1 = if ($null -ne $v1) { $v1.ToString().Trim() } else { "" }
            $s2 = if ($null -ne $v2) { $v2.ToString().Trim() } else { "" }
            return ($s1 -ine $s2)
        }

        function Local-GetVerfahren($Row) {
            if ($null -eq $Row) { return "[alle_Verfahren]" }
            $Active = New-Object System.Collections.Generic.List[string]
            $Cols = @("Viva", "Findus", "MobiApps", "AccVisio", "Verfahren5",
                      "Verfahren6", "Verfahren7", "Verfahren8", "Verfahren9", "Verfahren10")
            foreach ($c in $Cols) {
                if ($Row.$c -and $Row.$c.ToString().Trim().ToLower() -eq "x") { $Active.Add($c) }
            }
            # [FIX-7] 'return if (...)' ist kein gültiger PS-Ausdruck.
            # 'if' ist ein Statement, kein Wert -> Zuweisung zu Variable, dann return.
            $Result = if ($Active.Count -gt 0) { $Active -join " - " } else { "[alle_Verfahren]" }
            return $Result
        }

        $AndereOU = ""; $StatusGeloescht = ""; $UserGroups = @()
        if ($ADObj) {
            $dn = $ADObj.DistinguishedName
            if ($dn -notmatch "OU=81" -and $dn -notmatch "OU=82") {
                $AndereOU = ($dn -replace '^CN=.*?,', '')
            }
            if ($ADObj.MemberOf) {
                $UserGroups = $ADObj.MemberOf | ForEach-Object { ($_ -split ',')[0] -replace 'CN=', '' }
            }
        } else {
            $StatusGeloescht = "XXX"
        }

        $IsRequired = if ($RequiredList.Contains($LID.ToUpper())) { "ZZZ" } else { "" }
        $IsLafp = ($LID -like "L110*") -or ($LID -like "L114*")

        $TargetOrt = if ($CsvRow -and $CsvRow.Standort) {
            $CsvRow.Standort.ToString().Trim()
        } elseif ($ADObj) {
            $ADObj.l
        } else { "" }

        $TargetNachname = if ($IsLafp -and $TargetOrt -notmatch "^LAFP\s-\s") {
            "LAFP - $TargetOrt"
        } else { $TargetOrt }

        $BuroPlatz = if ($CsvRow -and $CsvRow.Raum) {
            $CsvRow.Raum.ToString().Trim()
        } else { "[Platznummer]" }
        
        $Fortbildung = if ($CsvRow -and $CsvRow.Fortbildungsbereich) {
            $CsvRow.Fortbildungsbereich.ToString().Trim()
        } else { "" }

        # [FIX-3] Doppelte Zuweisung korrigiert:
        # ALT: $Anmerkung = if (...) { $Anmerkung = $CsvRow.Anmerkungen... }
        #      -> innere Zuweisung gibt $null zurück, äußere blieb leer
        # NEU: Direkt den Ausdruck zurückgeben
        $Anmerkung = if ($CsvRow -and $CsvRow.Anmerkungen) {
            $CsvRow.Anmerkungen.ToString().Trim()
        } else { "" }

        $ÄnderInfo = if ($Fortbildung -and $Anmerkung) {
            "$Fortbildung - $Anmerkung"
        } else { "$Fortbildung$Anmerkung" }

        $VerfStr = Local-GetVerfahren -Row $CsvRow

        $Props = [ordered]@{
            "L-Kennung"              = $LID
            "andere OU"              = $AndereOU
            "GELÖSCHT"               = $StatusGeloescht
            "Benötigt"               = $IsRequired
            # [FIX-5] Geändert wird weiter unten befüllt (nach Codes-Berechnung)
            "Geändert"               = ""
            "LAFP_LZPD_LKA"          = if ($IsLafp) { "LLL" } else { "" }
            "NICHT KONFORM"          = ""
            "AD_Vorname"             = if ($ADObj) { $ADObj.GivenName } else { "" }
            "ÄNDERN_Vorname"         = $LID
            "AD_Nachname"            = if ($ADObj) { $ADObj.Surname } else { "" }
            "ÄNDERN_Nachname"        = $TargetNachname
            "AD_DisplayName"         = if ($ADObj) { $ADObj.DisplayName } else { "" }
            "ÄNDERN_DisplayName"     = "$LID $TargetNachname"
            "ORIGINAL_Standort"      = if ($CsvRow) { $CsvRow.Standort } else { "" }
            "ORIGINAL_Raum / Schulungskreis" = if ($CsvRow) { $CsvRow.Raum } else { "" }
            "AD_Ort"                 = if ($ADObj) { $ADObj.l } else { "" }
            "ÄNDERN_Ort"             = $TargetOrt
            "AD_Büro"                = if ($ADObj) { $ADObj.physicalDeliveryOfficeName } else { "" }
            "ÄNDERN_Büro"            = $BuroPlatz
            "ORIGINAL_Abteilung"     = if ($CsvRow) { $CsvRow.Abteilung } else { "" }
            "AD_Dez"                 = if ($ADObj) { $ADObj.department } else { "" }
            "ÄNDERN_Dez"             = if ($CsvRow -and $ADObj) {
                                           "$($CsvRow.Abteilung) - $($ADObj.department)"
                                       } else { "$($CsvRow.Abteilung)" }
            "AD_Description"         = if ($ADObj) { $ADObj.Description } else { "" }
            "ÄNDERN_Description"     = "$(if($IsLafp){'26'}else{'[OU]'}) - $VerfStr - $BuroPlatz - [Verantwortlicher] - [TEL] | $ÄnderInfo"
            "ÄNDERN_OU"              = if ($IsLafp) { "26" } else { "[OU]" }
            "ÄNDERN_Info"            = $ÄnderInfo
            "LÖSCHEN"                = ""
        }

        if ($ADObj) {
            $Codes = @()
            if (Local-IsDifferent $ADObj.GivenName              $Props."ÄNDERN_Vorname")      { $Codes += "VN"   }
            if (Local-IsDifferent $ADObj.Surname                $Props."ÄNDERN_Nachname")     { $Codes += "NN"   }
            if (Local-IsDifferent $ADObj.DisplayName            $Props."ÄNDERN_DisplayName")  { $Codes += "DN"   }
            if (Local-IsDifferent $ADObj.Description            $Props."ÄNDERN_Description")  { $Codes += "DESC" }
            if (Local-IsDifferent $ADObj.l                      $Props."ÄNDERN_Ort")          { $Codes += "ORT"  }
            if (Local-IsDifferent $ADObj.physicalDeliveryOfficeName $Props."ÄNDERN_Büro")     { $Codes += "GEB"  }
            if (Local-IsDifferent $ADObj.department             $Props."ÄNDERN_Dez")          { $Codes += "DEZ"  }
            if (Local-IsDifferent $ADObj.info                   $Props."ÄNDERN_Info")         { $Codes += "INFO" }

            if ($Codes.Count -gt 0) {
                $CodeStr = $Codes -join ","
                # [FIX-5] Beide Spalten mit denselben Codes befüllen
                $Props."NICHT KONFORM" = $CodeStr
                $Props."Geändert"      = $CodeStr
            }
        }

        foreach ($gn in $SortedGroups) {
            $Props["GRP_$gn"] = if ($UserGroups -contains $gn) { "X" } else { "" }
        }

        return [PSCustomObject]$Props
    }

    $Jobs = New-Object System.Collections.Generic.List[PSObject]
    foreach ($LID in $ProcessList) {
        $PSInstance = [powershell]::Create()
        [void]$PSInstance.AddScript($ScriptBlock)
        [void]$PSInstance.AddArgument($LID)
        [void]$PSInstance.AddArgument($MasterCsvData[$LID])
        [void]$PSInstance.AddArgument($ADCache[$LID])
        [void]$PSInstance.AddArgument($RequiredList)
        [void]$PSInstance.AddArgument($SortedGroups)
        $PSInstance.RunspacePool = $Pool
        [void]$Jobs.Add([PSCustomObject]@{
            SAM         = $LID
            Instance    = $PSInstance
            AsyncResult = $PSInstance.BeginInvoke()
        })
    }

    $Results   = New-Object System.Collections.Generic.List[PSObject]
    $Count     = 0
    $Total     = $Jobs.Count
    $ErrCount  = 0

    while ($Jobs.Count -gt 0) {
        $Finished = @($Jobs | Where-Object { $_.AsyncResult.IsCompleted })

        # [FIX-4] Remove-Liste getrennt sammeln, NICHT während foreach entfernen
        $ToRemove = New-Object System.Collections.Generic.List[PSObject]

        foreach ($Job in $Finished) {
            $Count++

            # [OPT-1] Fehler-Stream aus Runspace loggen
            if ($Job.Instance.Streams.Error.Count -gt 0) {
                $ErrCount++
                foreach ($e in $Job.Instance.Streams.Error) {
                    Write-Host "[$Count/$Total] FEHLER in $($Job.SAM): $e" -ForegroundColor Red
                    "[$( Get-Date -Format 'HH:mm:ss')] [ERROR] Runspace $($Job.SAM): $e" | Out-File $LogFile -Append
                }
            }

            # [FIX-2] PSDataCollection-Unrolling:
            # ALT: $item.GetType().Name -match "PSCustomObject|PSObject"
            #      -> zu streng; Runspace-Objekte sind PSObject-Wrapper, nicht "PSCustomObject"
            #      -> bei fehlgeschlagenem Filter wird die PSDataCollection selbst
            #         als Objekt behandelt (Spalten: IsOpen, Count, SyncRoot usw.)
            # NEU: Prüfung auf 'psobject' (existiert bei allen PS-Objekten)
            #      und expliziter Ausschluss von Sammlungs-Typen
            $RawOutput = $Job.Instance.EndInvoke($Job.AsyncResult)
            foreach ($item in $RawOutput) {
                if ($null -ne $item -and $item.PSObject -ne $null -and
                    $item -isnot [System.Management.Automation.PSDataCollection[System.Management.Automation.PSObject]]) {
                    [void]$Results.Add($item)
                }
            }

            Write-Host "[$Count/$Total] OK: $($Job.SAM)" -ForegroundColor Gray
            # [OPT-2] Progress-Bar
            Write-Progress -Activity "Verarbeite L-Kennungen" `
                           -Status "$Count von $Total abgeschlossen" `
                           -PercentComplete (($Count / $Total) * 100)

            $Job.Instance.Dispose()
            [void]$ToRemove.Add($Job)
        }

        # [FIX-4] Erst nach dem Loop entfernen
        foreach ($j in $ToRemove) { [void]$Jobs.Remove($j) }

        if ($Jobs.Count -gt 0) { Start-Sleep -Milliseconds 50 }
    }

    Write-Progress -Activity "Verarbeite L-Kennungen" -Completed
    $Pool.Close()
    $Pool.Dispose()

    return @{
        Results = $Results
        ErrCount = $ErrCount
    }
}

function Export-Results {
    param($Results, $SortedGroups, $ScriptDir, $Version, $LogFile)

    # 5. EXPORTE
    # [FIX-6] Bounds-Check vor $Results[0]-Zugriff
    if ($Results.Count -eq 0) {
        Write-Log "FEHLER: Keine Ergebnisse generiert. Bitte Eingabedaten und AD-Verbindung prüfen." "ERROR"
        return $null
    }

    Write-Log "Erzeuge Ausgabedateien ($($Results.Count) Datensätze)..." "INFO"

    # Tabelle 1: Gruppen-Übersicht
    $Path1    = Join-Path $ScriptDir "L-Kennungen_Full_Analysis_v$Version.csv"
    $BaseCols = @("L-Kennung", "andere OU", "GELÖSCHT", "Benötigt", "Geändert", "LÖSCHEN", "AD_Vorname", "AD_Nachname")
    $GrpCols  = $SortedGroups | ForEach-Object { "GRP_$_" }
    $Results | Select-Object ($BaseCols + $GrpCols) |
        Export-Csv -Path $Path1 -Delimiter ';' -NoTypeInformation -Encoding UTF8

    # Tabelle 2: Properties-Übersicht (alle Spalten außer LÖSCHEN)
    $Path2    = Join-Path $ScriptDir "L-Kennungen_Properties_v$Version.csv"
    $AllProps = $Results[0].psobject.Properties.Name | Where-Object { $_ -ne "LÖSCHEN" }

    # Robusteres CSV-Schreiben mit korrektem Escaping
    $CsvLines = New-Object System.Collections.Generic.List[string]
    $CsvLines.Add(($AllProps -join ';'))
    foreach ($r in $Results) {
        $row = $AllProps | ForEach-Object {
            $val = $r.$_
            $str = if ($null -ne $val) { $val.ToString() } else { "" }
            # Semikolon und Zeilenumbrüche im Wert escapen
            if ($str -match '[;\r\n"]') { '"' + $str.Replace('"', '""') + '"' } else { $str }
        }
        $CsvLines.Add(($row -join ';'))
    }
    $CsvLines | Out-File -FilePath $Path2 -Encoding UTF8 -Force

    return @{ Path1 = $Path1; Path2 = $Path2 }
}

function Show-Summary {
    param($Version, $ResultsCount, $ErrCount, $Stopwatch, $Path1, $Path2, $LogFile)
    # FINALER OUTPUT
    Write-Host "`n====================================================" -ForegroundColor Cyan
    Write-Host "ERGEBNISSE:" -ForegroundColor Yellow
    Write-Host "VERSION:`t$Version"
    Write-Host "DATUM:`t`t$(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')"
    Write-Host "ZEILEN:`t`t$ResultsCount"
    Write-Host "FEHLER:`t`t$ErrCount"
    Write-Host "DAUER:`t`t$($Stopwatch.Elapsed.TotalSeconds.ToString('F2'))s"
    Write-Host "TABELLE 1:`t$Path1"
    Write-Host "TABELLE 2:`t$Path2"
    Write-Host "LOG:`t`t$LogFile"
    Write-Host "====================================================`n" -ForegroundColor Cyan
}

function Invoke-Main {
    Show-Header
    Import-Module ActiveDirectory

    if (-not $CsvPath) { $CsvPath = Read-Host "Pfad zur Master-CSV eingeben" }
    $CsvPath = $CsvPath.Trim().Trim('"')
    if (-not (Test-Path $CsvPath)) { Write-Log "Master-CSV nicht gefunden: '$CsvPath'" "ERROR"; return }

    # 1. BEDARFSLISTE
    $RequiredList = Get-RequiredList -RequiredCsvPath $RequiredCsvPath

    # 2. MASTER-CSV
    $MasterCsvData = Get-MasterCsvData -CsvPath $CsvPath

    # 3. AD-DISCOVERY
    $Discovery = Invoke-ADDiscovery -SearchGlobal:$SearchGlobal -MasterCsvData $MasterCsvData -TestCount $TestCount
    $ADCache = $Discovery.ADCache
    $SortedGroups = $Discovery.SortedGroups
    $ProcessList = $Discovery.ProcessList

    Write-Log "Discovery beendet ($($Stopwatch.Elapsed.TotalSeconds.ToString('F2'))s). Starte Verarbeitung von $($ProcessList.Count) Einträgen..." "SUCCESS"

    # 4. PARALLELE VERARBEITUNG
    $Processing = Invoke-ParallelProcessing -ProcessList $ProcessList -MasterCsvData $MasterCsvData -ADCache $ADCache -RequiredList $RequiredList -SortedGroups $SortedGroups -MaxThreads $MaxThreads -LogFile $LogFile
    $Results = $Processing.Results
    $ErrCount = $Processing.ErrCount

    if ($ErrCount -gt 0) {
        Write-Log "WARNUNG: $ErrCount Runspace-Fehler aufgetreten. Details im Log: $LogFile" "WARN"
    }

    # 5. EXPORTE
    $Exports = Export-Results -Results $Results -SortedGroups $SortedGroups -ScriptDir $ScriptDir -Version $Version -LogFile $LogFile

    $Stopwatch.Stop()

    # FINALER OUTPUT
    if ($Exports) {
        Show-Summary -Version $Version -ResultsCount $Results.Count -ErrCount $ErrCount -Stopwatch $Stopwatch -Path1 $Exports.Path1 -Path2 $Exports.Path2 -LogFile $LogFile
    }
}
# [FIX-1] Aufruf der umbenannten Funktion
Invoke-Main
