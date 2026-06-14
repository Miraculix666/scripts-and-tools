<#
#https://gemini.google.com/gem/a0a0ee9c06a3/71a6525ab796d42f
.SYNOPSIS
    RDP User Manager & Mail Generator (v5.1)

.DESCRIPTION
    Professionelles Admin-Tool zur Verwaltung von lokalen RDP-Rechten und Verteilung von Zugangsdaten.
    
    SICHERHEITS-LOGIK (Admin vs. User):
    - Modus 1 & 2 (Rechte setzen): MUSS als Administrator laufen.
    - Modus 3 (Outlook): DARF NICHT als Administrator laufen (verhindert COM-Fehler).
    - Modus 3 (SMTP): Kann als Admin laufen.

.PARAMETER SetRDPRights
    Modus 1: Fügt Benutzer zur lokalen Gruppe "Remotedesktopbenutzer" hinzu.

.PARAMETER RemoveRDPRights
    Modus 2: Entfernt Benutzer aus der lokalen Gruppe.

.PARAMETER GenerateRDPFiles
    Modus 3: Erstellt nur RDP-Dateien und E-Mails (ohne Rechteänderung).
    Alias: generatemail

.PARAMETER GenerateFromLog
    Modus 4: Wiederholt den Mail-Versand basierend auf einem erfolgreichen Log-File.
    Alias: genfromlog

.NOTES
    Version:    5.1 (Strict Admin Checks, Improved Mail Syntax, Robust Reporting)
    Autor:      PS-Coding
    Datum:      18.11.2025
#>

[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Interactive')]
param(
    #--- MODI ---
    [Parameter(Mandatory = $false, ParameterSetName = 'Set-RDPRights')][Switch]$SetRDPRights,
    [Parameter(Mandatory = $false, ParameterSetName = 'Remove-RDPRights')][Switch]$RemoveRDPRights,
    [Parameter(Mandatory = $false, ParameterSetName = 'Generate-RDPFiles')][Alias('generatemail')][Switch]$GenerateRDPFiles,
    [Parameter(Mandatory = $false, ParameterSetName = 'GenerateFromLog')][Alias('genfromlog')][Switch]$GenerateFromLog,
    
    #--- QUELLEN ---
    [Parameter(Mandatory = $true, ParameterSetName = 'Set-RDPRights')]
    [Parameter(Mandatory = $true, ParameterSetName = 'Remove-RDPRights')]
    [Parameter(Mandatory = $true, ParameterSetName = 'Generate-RDPFiles')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })][string]$UserListPath,

    [Parameter(Mandatory = $true, ParameterSetName = 'Set-RDPRights')]
    [Parameter(Mandatory = $true, ParameterSetName = 'Remove-RDPRights')]
    [Parameter(Mandatory = $true, ParameterSetName = 'Generate-RDPFiles')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })][string]$ClientListPath,

    [Parameter(Mandatory = $true, ParameterSetName = 'GenerateFromLog')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })][string]$InputLogPath,

    #--- E-MAIL OPTIONEN ---
    [Parameter(Mandatory = $false)][Switch]$SendEmail, 
    [Parameter(Mandatory = $false)][Switch]$SaveAsMsgOnly, 
    [Parameter(Mandatory = $false)][string]$SmtpServer, 
    [Parameter(Mandatory = $false)][System.Management.Automation.PSCredential]$Credential = $null, 

    #--- SONSTIGES ---
    [Parameter(Mandatory = $false)][string]$OutputPath, 
    [Parameter(Mandatory = $false)][string]$UserColumn = 'sAMAccountName',
    [Parameter(Mandatory = $false)][string]$ClientColumn = 'ComputerName'
)

#==============================================================================
# GLOBALE INIT
#==============================================================================
$Version = "5.1"
$GlobalErrorLog = [System.Collections.ArrayList]::new()
$GlobalTimestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$GlobalLogDir = Join-Path -Path $PSScriptRoot -ChildPath "Logs"

#==============================================================================
# HELPER FUNCTIONS
#==============================================================================

Function Test-IsAdmin {
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [System.Security.Principal.WindowsPrincipal]$identity
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

Function Get-DynamicADDomain {
    try {
        $adDomain = Get-ADDomain -ErrorAction Stop
        return [PSCustomObject]@{ NetBIOS = $adDomain.NetBIOSName; FQDN = $adDomain.DNSRoot }
    }
    catch {
        $msg = "WARNUNG: AD-Domäne nicht ermittelbar. Fallback auf 'WORKGROUP'."
        Write-Warning $msg
        $GlobalErrorLog.Add($msg) | Out-Null
        return [PSCustomObject]@{ NetBIOS = "WORKGROUP"; FQDN = $null }
    }
}

Function Invoke-RemoteGroupMembership {
    param($ComputerName, $UserName, $Domain, $Action, $LocalGroupName = "Remotedesktopbenutzer")

    Write-Verbose ("Verbinde zu {0}..." -f $ComputerName)
    $status = 'Failed' 

    try {
        $group = [ADSI]"WinNT://$ComputerName/$LocalGroupName,group"
        $userPath = "WinNT://$Domain/$UserName,user"
        
        if ($Action -eq 'Add') {
            $group.Add($userPath) | Out-Null
            $group.RefreshCache()
            # Echte Prüfung
            $members = @($group.Invoke("Members")) | ForEach-Object { $_.GetType().InvokeMember("Name", "GetProperty", $null, $_, $null) }
            if ($members -contains $UserName) { $status = 'Success' } else { $status = 'VerificationFailed' }
        }
        elseif ($Action -eq 'Remove') {
            $group.Remove($userPath) | Out-Null
            $group.RefreshCache()
            $members = @($group.Invoke("Members")) | ForEach-Object { $_.GetType().InvokeMember("Name", "GetProperty", $null, $_, $null) }
            if (-not ($members -contains $UserName)) { $status = 'Success' } else { $status = 'VerificationFailed' }
        }
    }
    catch {
        $err = $_.Exception.Message.Trim()
        if ($err -like "*bereits Mitglied*") { $status = 'AlreadyExists' }
        elseif ($err -like "*nicht Mitglied*") { $status = 'NotMember' }
        else {
            $msg = "ADSI-Fehler ($ComputerName): $err"
            Write-Warning $msg
            $GlobalErrorLog.Add($msg) | Out-Null
            $status = 'Failed'
        }
    }
    return $status
}

Function Create-RDPFile {
    param($ComputerName, $FilePath)
    
    $sb = [System.Text.StringBuilder]::new()
    $sb.AppendLine("screen mode id:i:2") | Out-Null
    $sb.AppendLine("full address:s:$ComputerName") | Out-Null
    $sb.AppendLine("prompt for credentials:i:1") | Out-Null
    $sb.AppendLine("redirectclipboard:i:1") | Out-Null
    
    try {
        Set-Content -Path $FilePath -Value $sb.ToString() -Encoding UTF8 -ErrorAction Stop
        return $true
    }
    catch {
        $msg = "Fehler RDP-Datei ($FilePath): $($_.Exception.Message)"
        Write-Warning $msg
        $GlobalErrorLog.Add($msg) | Out-Null
        return $false
    }
}

Function Get-RDPEmailRecipient {
    param($UserName)
    try {
        if (-not (Get-Module -Name ActiveDirectory)) { Import-Module ActiveDirectory -ErrorAction Stop }
        $adUser = Get-ADUser -Identity $UserName -Properties GivenName, Surname, EmailAddress -ErrorAction Stop
    } catch {
        $msg = "AD-Fehler ($UserName): $($_.Exception.Message)"
        Write-Warning $msg
        if ($Global:GlobalErrorLog) { $Global:GlobalErrorLog.Add($msg) | Out-Null }
        return $null
    }
    
    $to = $adUser.EmailAddress
    if ($adUser.GivenName -and $adUser.Surname) {
        $fullName = "$($adUser.GivenName) $($adUser.Surname)"
    } else {
        $fullName = $UserName
    }
    
    if (-not $to) {
        $msg = "Keine E-Mail im AD für $UserName."
        Write-Warning $msg
        if ($Global:GlobalErrorLog) { $Global:GlobalErrorLog.Add($msg) | Out-Null }
        return $null
    }
    
    return [PSCustomObject]@{ To = $to; FullName = $fullName }
}

Function Get-RDPEmailBody {
    param($FullName)
    $sb = [System.Text.StringBuilder]::new()
    $sb.AppendLine("<html><body style='font-family:Calibri, Arial, sans-serif; font-size:11pt; color:#333;'>") | Out-Null
    $sb.AppendLine("<p>Guten Tag $FullName,</p>") | Out-Null
    $sb.AppendLine("<p>anbei erhalten Sie Ihre persönlichen Zugangsdateien für den Remote-Desktop-Zugriff auf die Schulungsrechner.</p>") | Out-Null
    $sb.AppendLine("<div style='background-color:#f9f9f9; padding:15px; border-left: 4px solid #0078d4; margin: 10px 0;'>") | Out-Null
    $sb.AppendLine("<strong>Ihre Schritte zur Anmeldung:</strong>") | Out-Null
    $sb.AppendLine("<ol>") | Out-Null
    $sb.AppendLine("<li>Speichern Sie die angehängte(n) <b>.rdp-Datei(en)</b> auf Ihrem Desktop ab.</li>") | Out-Null
    $sb.AppendLine("<li>Starten Sie die Verbindung durch einen Doppelklick auf die Datei.</li>") | Out-Null
    $sb.AppendLine("<li>Geben Sie bei der Abfrage Ihre <b>gewohnten Windows-Zugangsdaten</b> ein (Benutzername und Kennwort, die Sie auch lokal verwenden).</li>") | Out-Null
    $sb.AppendLine("</ol>") | Out-Null
    $sb.AppendLine("</div>") | Out-Null
    $sb.AppendLine("<p>Sollten Probleme bei der Verbindung auftreten, wenden Sie sich bitte an den IT-Support.</p>") | Out-Null
    $sb.AppendLine("<br><p style='font-size:9pt; color:#888;'>Dies ist eine automatisch generierte Nachricht.</p>") | Out-Null
    $sb.AppendLine("</body></html>") | Out-Null
    return $sb.ToString()
}

Function Send-RDPEmailSmtp {
    param($UserName, $To, $Subject, $Body, $RDPFilePaths, $SmtpServer, $Credential, $Send)
    if (-not $Send) { return $false }
    if ($null -eq $Credential) {
        Write-Host " SMTP-Login erforderlich..." -ForegroundColor Yellow
        $Credential = Get-Credential -Message "SMTP Login ($SmtpServer)"
    }
    $p = @{ To=$To; From=$To; Subject=$Subject; Body=$Body; BodyAsHtml=$true; SmtpServer=$SmtpServer; Attachments=$RDPFilePaths; ErrorAction='Stop' }
    if ($Credential.UserName) { $p.Add('Credential', $Credential) }

    try {
        Send-MailMessage @p
        Write-Host (" [SMTP] Gesendet: {0}" -f $To) -ForegroundColor Cyan
        return $true
    } catch {
        Write-Warning "SMTP-Fehler ($UserName): $($_.Exception.Message)"
        return $false
    }
}

Function Send-RDPEmailOutlook {
    param($UserName, $To, $Subject, $Body, $RDPFilePaths, $MSGSavePath, $Send)
    try {
        try { $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') }
        catch { $outlook = New-Object -ComObject Outlook.Application }

        if (!$outlook) { throw "Outlook läuft nicht/nicht installiert." }

        $mail = $outlook.CreateItem(0)
        $mail.Subject = $Subject
        $mail.To = $To
        $mail.HTMLBody = $Body

        foreach ($f in $RDPFilePaths) {
            if (Test-Path $f) {
                [void]$mail.Attachments.Add($f) # Cast [void] unterdrückt Output
            }
        }
        
        if ($MSGSavePath) {
            [void]$mail.SaveAs($MSGSavePath, 3) # 3=olMsg. Speichern.
            Write-Host (" [MSG] Gespeichert: {0}" -f ($MSGSavePath | Split-Path -Leaf)) -ForegroundColor Green
        }
        if ($Send) {
            [void]$mail.Send()
            Write-Host (" [OUTLOOK] Gesendet: {0}" -f $To) -ForegroundColor Cyan
        }

        # Cleanup Versuch
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail) | Out-Null
        return $true
    } catch {
        $err = $_.Exception.Message
        Write-Warning "Outlook-Fehler ($UserName): $err"
        return $false
    }
}

Function Start-EmailWorkflow {
    <# Versendet E-Mails. Gibt $true/$false zurück. Unterdrückt COM-Ausgaben. #>
    param($UserName, $RDPFilePaths, $MSGSavePath, $Send, $SmtpServer, $Credential)

    # 1. AD Daten (Vorname Nachname)
    $recipient = Get-RDPEmailRecipient -UserName $UserName
    if (-not $recipient) { return $false }

    # 2. E-Mail Text
    $body = Get-RDPEmailBody -FullName $recipient.FullName
    $subj = "IT-Support: Ihre RDP-Verbindungsdaten"

    # 3. Versand
    if ($SmtpServer) {
        return Send-RDPEmailSmtp -UserName $UserName -To $recipient.To -Subject $subj -Body $body -RDPFilePaths $RDPFilePaths -SmtpServer $SmtpServer -Credential $Credential -Send $Send
    } else {
        return Send-RDPEmailOutlook -UserName $UserName -To $recipient.To -Subject $subj -Body $body -RDPFilePaths $RDPFilePaths -MSGSavePath $MSGSavePath -Send $Send
    }
}

Function Generate-SendMailsScript {
    <# Generiert das Helper-Skript für den User-Kontext #>
    param($UserList, $RDPFilePaths, $OutputPath, $SmtpServer, $SaveAsMsgOnly)
    
    Write-Verbose "Generiere Helper-Script 'sendMails.ps1'..."
    $resOut = Resolve-Path $OutputPath
    
    # Daten serialisieren
    $uStr = ($UserList | ForEach-Object { "'$_'" }) -join ","
    $rStr = ($RDPFilePaths | ForEach-Object { $p=Resolve-Path $_; "'$p'" }) -join ","
    $smtpStr = if ($SmtpServer) { "'$SmtpServer'" } else { "`$null" }
    $sendBool = if ($SaveAsMsgOnly) { "`$false" } else { "`$true" }
    
    # Funktions-Body kopieren
    $funcDefRecip = (Get-Command Get-RDPEmailRecipient).Definition
    $funcDefBody = (Get-Command Get-RDPEmailBody).Definition
    $funcDefSmtp = (Get-Command Send-RDPEmailSmtp).Definition
    $funcDefOutl = (Get-Command Send-RDPEmailOutlook).Definition
    $funcDefWork = (Get-Command Start-EmailWorkflow).Definition

    $sb = [System.Text.StringBuilder]::new()
    $sb.AppendLine("# AUTOMATISCH GENERIERTES SCRIPT - FUER USER-KONTEXT") | Out-Null
    $sb.AppendLine("# Ausfuehren als: NORMALER BENUTZER (Mit Outlook-Zugriff)") | Out-Null
    $sb.AppendLine("param(`$AltSmtp)") | Out-Null
    $sb.AppendLine("Function Get-RDPEmailRecipient { $funcDefRecip }") | Out-Null
    $sb.AppendLine("Function Get-RDPEmailBody { $funcDefBody }") | Out-Null
    $sb.AppendLine("Function Send-RDPEmailSmtp { $funcDefSmtp }") | Out-Null
    $sb.AppendLine("Function Send-RDPEmailOutlook { $funcDefOutl }") | Out-Null
    $sb.AppendLine("Function Start-EmailWorkflow { $funcDefWork }") | Out-Null
    $sb.AppendLine("") | Out-Null
    $sb.AppendLine("`$Users = @($uStr)") | Out-Null
    $sb.AppendLine("`$Files = @($rStr)") | Out-Null
    $sb.AppendLine("`$Smtp = $smtpStr") | Out-Null
    $sb.AppendLine("if (`$AltSmtp) { `$Smtp = `$AltSmtp }") | Out-Null
    $sb.AppendLine("`$DoSend = $sendBool") | Out-Null
    $sb.AppendLine("`$OutDir = '$resOut'") | Out-Null
    $sb.AppendLine("") | Out-Null
    $sb.AppendLine("Write-Host 'Starte E-Mail-Verarbeitung...' -ForegroundColor Cyan") | Out-Null
    $sb.AppendLine("foreach (`$u in `$Users) {") | Out-Null
    $sb.AppendLine("    `$m = Join-Path `$OutDir (`$u + '.msg')") | Out-Null
    $sb.AppendLine("    if (`$Smtp) { `$m = `$null }") | Out-Null
    $sb.AppendLine("    Write-Host ' > User:' `$u -NoNewline") | Out-Null
    $sb.AppendLine("    Start-EmailWorkflow -UserName `$u -RDPFilePaths `$Files -MSGSavePath `$m -Send `$DoSend -SmtpServer `$Smtp -Credential `$null") | Out-Null
    $sb.AppendLine("}") | Out-Null
    $sb.AppendLine("Read-Host 'Fertig. Enter druecken.'") | Out-Null

    try {
        $f = Join-Path $OutputPath "sendMails_$($GlobalTimestamp).ps1"
        Set-Content $f -Value $sb.ToString() -Encoding UTF8
        Write-Host " [INFO] Helper-Script für Outlook erstellt:" -ForegroundColor Gray
        Write-Host "        $f" -ForegroundColor White
    } catch {
        Write-Warning "Fehler beim Erstellen des Helper-Scripts."
    }
}

Function Write-Log {
    param($Data, $Type)
    if (-not (Test-Path $GlobalLogDir)) { New-Item $GlobalLogDir -Type Directory -Force | Out-Null }
    $f = Join-Path $GlobalLogDir "Log_${Type}_$($GlobalTimestamp).csv"
    $Data | Export-Csv -Path $f -NoType -Delimiter ';' -Encoding UTF8
    Write-Host (" [LOG] Protokoll: {0}" -f $f) -ForegroundColor DarkGreen
}

Function Load-CsvData {
    param($CsvPath)
    if (-not (Test-Path $CsvPath)) { Write-Error ("Datei fehlt: {0}" -f $CsvPath); return $null }
    try {
        $d = Import-Csv $CsvPath -Delimiter ';' -Encoding Default -ErrorAction Stop
        if (($d | Measure-Object).Count -eq 0) { Write-Error "Datei leer."; return $null }
        return $d
    } catch { Write-Error "CSV-Fehler: $($_.Exception.Message)"; return $null }
}

Function Write-ErrorLog {
    param([string]$OutPath)
    if (($GlobalErrorLog | Measure-Object).Count -gt 0) {
        try {
            if (-not (Test-Path $OutPath)) { New-Item -Path $OutPath -ItemType Directory -Force | Out-Null }
            $f = Join-Path $OutPath "ERROR.TXT"
            $GlobalErrorLog | Out-File $f -Encoding UTF8
            Write-Warning "Fehler aufgetreten. Details in: $f"
        } catch {}
    }
}

#==============================================================================
# WORKFLOWS
#==============================================================================

Function Start-RightsWorkflow {
    param($Set, $UserListPath, $ClientListPath, $UserCol, $ClientCol, $DoMail, $Smtp, $Cred, $OutPath)

    # Sicherheits-Check: Modus 1/2 MUSS Admin sein
    if (-not (Test-IsAdmin)) {
        Write-Error "KRITISCHER FEHLER: Um Rechte zu setzen, MUSS dieses Skript 'Als Administrator' ausgeführt werden."
        return
    }

    $Users = Load-CsvData $UserListPath
    $Clients = Load-CsvData $ClientListPath
    if (!$Users -or !$Clients) { return }

    if (-not (Get-Module -Name ActiveDirectory)) { try{Import-Module ActiveDirectory -ErrorAction Stop}catch{Write-Error "AD Modul fehlt";return} }
    $AD = Get-DynamicADDomain
    
    $Act = if ($Set) { 'Add' } else { 'Remove' }
    $Verb = if ($Set) { "Hinzufügen" } else { "Entfernen" }

    Write-Host ("=== MODUS: Rechte {0} ===" -f $Verb) -ForegroundColor Cyan
    
    # 1. Planung
    Write-Host " [1/3] Planung & Ping-Test..." -ForegroundColor Yellow
    $Plan = @()
    foreach ($c in $Clients) {
        $cN = $c.$ClientCol
        if (!$cN) { continue }
        if (Test-Connection $cN -Count 1 -Quiet) {
            foreach ($u in $Users) { if ($u.$UserCol) { $Plan += [PSCustomObject]@{ Client=$cN; User=$u.$UserCol; Action=$Act } } }
        } else {
            Write-Warning "Client $cN offline."
            $GlobalErrorLog.Add("Client $cN offline") | Out-Null
        }
    }
    
    $cnt = ($Plan | Measure-Object).Count
    if ($cnt -eq 0) { Write-Warning "Keine Aktionen möglich."; return }
    
    # 2. Bestätigung
    Write-Host (" [2/3] {0} Aktionen geplant." -f $cnt) -ForegroundColor Yellow
    $Plan | Format-Table -AutoSize
    
    $go = $false
    if ($PSCmdlet.ShouldProcess("$cnt Aktionen", "Ausführen")) {
        if ($PSBoundParameters.ContainsKey('Confirm') -and -not $Confirm) { $go = $true }
        else {
            $in = Read-Host " Starten? (J/N)"
            if ($in -eq 'J') { $go = $true }
        }
    }
    
    if ($go) {
        # 3. Ausführung
        Write-Host " [3/3] Ausführung..." -ForegroundColor Cyan
        $Rep = @()
        foreach ($p in $Plan) {
            $s = Invoke-RemoteGroupMembership -ComputerName $p.Client -UserName $p.User -Domain $AD.NetBIOS -Action $p.Action
            
            $c = "Red"; if ($s -eq 'Success') { $c="Green" } elseif ($s -match 'Already|NotMember') { $c="Gray" }
            Write-Host (" {0}: {1} -> {2}" -f $s, $p.User, $p.Client) -ForegroundColor $c
            
            $Rep += [PSCustomObject]@{ Client=$p.Client; User=$p.User; Action=$p.Action; Status=$s; Time=Get-Date }
        }
        Write-Log $Rep "Rights"
        Write-ErrorLog $GlobalLogDir
        
        # Chaining
        if ($DoMail) {
            $suc = $Rep | Where { $_.Status -in 'Success','AlreadyExists' }
            if (($suc | Measure-Object).Count -gt 0) {
                $ul = $suc | Select -Expand User -Unique
                $cl = $suc | Select -Expand Client -Unique
                Start-FileWorkflow -IsCombined $true -UserList $ul -ClientList $cl -OutputPath $OutPath -Smtp $Smtp -Cred $Cred -UserCol "User" -ClientCol "Client"
            }
        }
    }
}

Function Start-FileWorkflow {
    param($UserListPath, $ClientListPath, $LogPath, $FromLog, $UserList, $ClientList, $IsCombined, $OutputPath, $UserCol, $ClientCol, $SaveOnly, $Smtp, $Cred)

    # Sicherheits-Check: Outlook ohne SMTP darf NICHT als Admin laufen
    if (-not $Smtp -and (Test-IsAdmin)) {
        Write-Error "KRITISCHER FEHLER: Outlook-Versand darf NICHT 'Als Administrator' gestartet werden."
        Write-Warning "Bitte nutzen Sie das generierte 'sendMails.ps1' später als normaler Benutzer."
        # Wir brechen hier NICHT hart ab, sondern generieren nur die Dateien & das Helper-Script
        $SkipDirectMail = $true
    } else {
        $SkipDirectMail = $false
    }

    $U=$null; $C=$null
    if ($IsCombined) { $U=$UserList; $C=$ClientList; $UserCol="User"; $ClientCol="Client" }
    elseif ($FromLog) {
        $l = Load-CsvData $LogPath
        if (!$l) { return }
        $suc = $l | Where { $_.Status -in 'Success','AlreadyExists' }
        $U = $suc | Select -Expand User -Unique
        $C = $suc | Select -Expand Client -Unique
        $UserCol="User"; $ClientCol="Client"
    }
    else {
        $U = Load-CsvData $UserListPath
        $C = Load-CsvData $ClientListPath
        if (!$U -or !$C) { return }
    }
    
    if (-not (Get-Module -Name ActiveDirectory)) { Import-Module ActiveDirectory }
    
    if (!$OutputPath) { $OutputPath = Join-Path $PSScriptRoot "RDP_Output_$GlobalTimestamp" }
    if (-not (Test-Path $OutputPath)) { New-Item $OutputPath -Type Directory -Force | Out-Null }
    
    Write-Host "=== MODUS: Dateien & Mails ===" -ForegroundColor Cyan
    
    # RDPs
    $rdps = @()
    $cObjs = if ($IsCombined -or $FromLog) { $C } else { $C.$ClientCol }
    foreach ($c in $cObjs) {
        if (!$c) { continue }
        $p = Join-Path $OutputPath "$c.rdp"
        if (Create-RDPFile $c $p) { $rdps += $p }
    }
    if ($rdps.Count -eq 0) { Write-Error "Keine RDPs erstellt."; return }
    Write-Host " $(($rdps|Measure-Object).Count) RDP-Dateien erstellt." -ForegroundColor Green
    
    # Mails
    $doSend = -not $SaveOnly
    $Rep = @()
    $uObjs = if ($IsCombined -or $FromLog) { $U } else { $U.$UserCol }
    
    if ($SkipDirectMail) {
        Write-Warning "Überspringe direkten E-Mail-Versand (Admin-Modus). Generiere nur Helper-Script..."
    } else {
        Write-Host " Verarbeite E-Mails..."
        foreach ($un in $uObjs) {
            if (!$un) { continue }
            $mp = if (!$Smtp) { Join-Path $OutputPath "$un.msg" } else { $null }
            Write-Host " User: $un" -NoNewline
            $res = Start-EmailWorkflow -UserName $un -RDPFilePaths $rdps -MSGSavePath $mp -Send $doSend -SmtpServer $Smtp -Credential $Cred
            $Rep += [PSCustomObject]@{ User=$un; Action=(if($doSend){"Send"}else{"Save"}); Status=$res; Time=Get-Date }
        }
        Write-Log -Data $Rep -Type "Mails"
    }
    
    Write-ErrorLog $OutputPath
    
    Generate-SendMailsScript -UserList $uObjs -RDPFilePaths $rdps -OutputPath $OutputPath -SmtpServer $Smtp -SaveAsMsgOnly $SaveOnly
    Write-Host "Fertig. Daten in: $OutputPath" -ForegroundColor Green
}

#==============================================================================
# START
#==============================================================================
Write-Host ("Manage-RDPUsers v{0}" -f $Version) -ForegroundColor White

if ($PSCmdlet.ParameterSetName -ne 'Interactive') {
    if ($SetRDPRights -or $RemoveRDPRights) {
        Start-RightsWorkflow -Set $SetRDPRights -UserListPath $UserListPath -ClientListPath $ClientListPath -UserCol $UserColumn -ClientCol $ClientColumn -DoMail $SendEmail -Smtp $SmtpServer -Cred $Credential -OutPath $OutputPath
    }
    elseif ($GenerateRDPFiles) {
        Start-FileWorkflow -UserListPath $UserListPath -ClientListPath $ClientListPath -OutputPath $OutputPath -UserCol $UserColumn -ClientCol $ClientColumn -SaveOnly $SaveAsMsgOnly -Smtp $SmtpServer -Cred $Credential
    }
    elseif ($GenerateFromLog) {
        Start-FileWorkflow -FromLog $true -InputLogPath $InputLogPath -OutputPath $OutputPath -SaveOnly $SaveAsMsgOnly -Smtp $SmtpServer -Cred $Credential
    }
}
else {
    Show-MainMenu
    $sel = Read-Host "Wahl"
    if ($sel -eq '1') {
        $u = Read-Host "User-CSV"; $c = Read-Host "Client-CSV"
        Start-RightsWorkflow -Set $true -UserListPath $u -ClientListPath $c
    }
    if ($sel -eq '3') {
        $u = Read-Host "User-CSV"; $c = Read-Host "Client-CSV"
        Start-FileWorkflow -UserListPath $u -ClientListPath $c
    }
}
Write-Host "Ende."
