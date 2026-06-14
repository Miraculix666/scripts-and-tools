### aufiltern von L Kennungen die Länger als 39 Wochen nicht im AD angemeldet wurden, da sie nach 54 Wochen (ein Jahr) deaktiviert werden und in der Deponie landen 


# Define the distinguished names of the Organizational Units (OUs)
$ou81 = "OU=81,OU=Polizei-NRW-PB-PE-2012,DC=polizei,DC=nrw,DC=de"
$ou82 = "OU=82,OU=Polizei-NRW-PB-PE-2012,DC=polizei,DC=nrw,DC=de"

# Define the output file paths
$outputFilePath = "C:\Daten\Users_LastLogon_Report.txt"
$samAccountNamesFilePath = "C:\Daten\All_SAMAccountNames.txt"
$expiredUsersFilePath = "C:\Daten\Expired_Users_SAM.txt"

# Get users from OU 81 with "L110" or "L114" in the username
$usersOU81 = Get-ADUser -Filter {Enabled -eq $true -and (Name -like "L110*" -or Name -like "L114*")} -SearchBase $ou81

# Get users from OU 82 with "L110" or "L114" in the username
$usersOU82 = Get-ADUser -Filter {Enabled -eq $true -and (Name -like "L110*" -or Name -like "L114*")} -SearchBase $ou82

# Combine the results from both OUs
$allUsers = $usersOU81 + $usersOU82

# Display the results
if ($allUsers.Count -gt 0) {
    Write-Host "Users in specified OUs with 'L110' or 'L114' in the username:"
    $allUsers | Select-Object Name, SamAccountName, DistinguishedName | Format-Table -AutoSize

    # Write all SAMAccountNames in the given scope into a file
    $allUsers.SamAccountName | Out-File -Append -FilePath $samAccountNamesFilePath

    # Get all users and their last logon date
    $usersLastLogon = $allUsers | Get-ADUser -Properties LastLogonDate | Select-Object Name, SamAccountName, LastLogonDate

    # Sort users by the oldest last logon date first
    $usersLastLogon = $usersLastLogon | Sort-Object LastLogonDate

    # Calculate the date 39 weeks ago
    $thresholdDate = (Get-Date).AddDays(-39 * 7)

    # Format the results and mark dates older than 39 weeks 
    $results = $usersLastLogon | ForEach-Object {
        $lastLogonDate = $_.LastLogonDate
        if ($lastLogonDate -ne $null) {
            $formattedDate = Get-Date $lastLogonDate -Format "yyyy-MM-dd HH:mm:ss"
            $formattedDatemarked = if ($lastLogonDate -lt $thresholdDate) {  "xxxxx $formattedDate" } else { $formattedDate }
        } else {
            $formattedDatemarked = "xxxxx N/A"
        }
        [PSCustomObject]@{
            Name            = $_.Name
            SamAccountName  = $_.SamAccountName
            LastLogonDate   = $formattedDatemarked
        }
    }

    # Write all SAMAccountNames with logon older than 39 weeks into a file
    $expiredUsers = $results | Where-Object { $_.LastLogonDate -like "xxxxx*" }
    $expiredUsers.SamAccountName | Out-File -Append -FilePath $expiredUsersFilePath

    # Display the results
    if ($results.Count -gt 0) {
        Write-Host "Users and their last logon dates:"
        $results | Format-Table -AutoSize

        # Write results to a file
        $results | Export-Csv -Path $outputFilePath -NoTypeInformation
        Write-Host "Output saved to: $outputFilePath"
    } else {
        Write-Host "No users found with last logon dates."
    }
} else {
    Write-Host "No users found in the specified OUs with 'L110' or 'L114' in the username."
}

Write-Host "All SAMAccountNames have been written to: $samAccountNamesFilePath"
Write-Host "Expired SAMAccountNames have been written to: $expiredUsersFilePath"
----------------------------------------------
### This script looksup after deactiveted L-Kennungen in "Deponie" 

# Set the distinguished names of the Organizational Units (OU)
$ou81 = "OU=81-VIVA,OU=Deponie,DC=polizei,DC=nrw,DC=de"
$ou82 = "OU=82-VIVA,OU=Deponie,DC=polizei,DC=nrw,DC=de"

# Define the output file path
$outputFilePath = "c:\Daten\Deaktivierte_L_Kennung.csv"
$outputFilePathSAM = "c:\Daten\Deaktivierte_L_Kennung_SAM.txt"

# Get all deactivated users in OU 81 with "L110" or "L114" in the username
$deactivatedUsersOU81 = Get-ADUser -Filter {Enabled -eq $false -and (Name -like "L110*" -or Name -like "L114*")} -SearchBase $ou81

# Get all deactivated users in OU 82 with  "L110" or "L114" in the username
$deactivatedUsersOU82 = Get-ADUser -Filter {Enabled -eq $false -and (Name -like "L110*" -or Name -like "L114*")} -SearchBase $ou82

# Combine the results from both OUs
$allDeactivatedUsers = $deactivatedUsersOU81 + $deactivatedUsersOU82

# Sort the users first by OU, then by Name
$allDeactivatedUsers = $allDeactivatedUsers | Sort-Object @{Expression={$_ | ForEach-Object { $_.DistinguishedName -match "OU=(\d+)-VIVA"; [int]$matches[1] }}} , Name

# Display the results and write to CSV file
if ($allDeactivatedUsers.Count -gt 0) {
    Write-Host "Deactivated users in specified search path with names containing '110' or '114':"
    $allDeactivatedUsers | Select-Object Name, SamAccountName, DistinguishedName
    $allDeactivatedUsers | Select-Object Name, SamAccountName, DistinguishedName | Export-Csv -Path $outputFilePath -NoTypeInformation
    Write-Host "Output saved to: $outputFilePath"}

# Write SamAccountNames to the file
$allDeactivatedUsers.SamAccountName | Out-File -Append -FilePath $outputFilePathSAM
# Display message about the file
Write-Host "SamAccountNames have been written to: $outputFilePathSAM"
----------------------------------------------------------

### This script looksup after deactiveted L-Kennungen in "Deponie" 

# Set the distinguished names of the Organizational Units (OU)
$ouDeponie = "OU=Deponie,DC=polizei,DC=nrw,DC=de"

# Define the output file path
$outputFilePath = "c:\Daten\Deaktivierte_Kennung.csv"
$outputFilePathSAM = "c:\Daten\Deaktivierte_Kennung_SAM.txt"

# Get all deactivated users in OU Deponie
$deactivatedUsersDeponie = Get-ADUser -Filter {Enabled -eq $false } -SearchBase $ouDeponie

# Sort the users first by OU, then by Name
$allDeactivatedUsers = $deactivatedUsersDeponie | Sort-Object @{Expression={$_ | ForEach-Object { $_.DistinguishedName -match "OU=(\d+)-VIVA"; [int]$matches[1] }}} , Name

# Display the results and write to CSV file
if ($allDeactivatedUsers.Count -gt 0) {
    Write-Host "Deactivated users in Deponie:"
    $allDeactivatedUsers | Select-Object Name, SamAccountName, DistinguishedName
    $allDeactivatedUsers | Select-Object Name, SamAccountName, DistinguishedName | Export-Csv -Path $outputFilePath -NoTypeInformation
    Write-Host "Output saved to: $outputFilePath"}

# Write SamAccountNames to the file
$allDeactivatedUsers.SamAccountName | Out-File -Append -FilePath $outputFilePathSAM
# Display message about the file
Write-Host "SamAccountNames have been written to: $outputFilePathSAM"




$allDeactivatedUsers | Select-Object SamAccountName | .\Get-ADUserLockouts.ps1
-----------------------------------------------------------
Function Get-ADUserLockouts {
    [CmdletBinding(
        DefaultParameterSetName = 'All'
    )]
    param (
        [Parameter(
            ValueFromPipeline = $true,
            ParameterSetName = 'ByUser'
        )]
        [Microsoft.ActiveDirectory.Management.ADUser]$Identity
        ,
        [datetime]$StartTime
        ,
        [datetime]$EndTime
    )
    Begin{
        $filterHt = @{
            LogName = 'Security'
            ID = 4740
        }
        if ($PSBoundParameters.ContainsKey('StartTime')){
            $filterHt['StartTime'] = $StartTime
        }
        if ($PSBoundParameters.ContainsKey('EndTime')){
            $filterHt['EndTime'] = $EndTime
        }
        $PDCEmulator = (Get-ADDomain).PDCEmulator
        # Query the event log just once instead of for each user if using the pipeline
        $events = Get-WinEvent -ComputerName $PDCEmulator -FilterHashtable $filterHt
    }
    Process {
        if ($PSCmdlet.ParameterSetName -eq 'ByUser'){
            $user = Get-ADUser $Identity
            # Filter the events
            $output = $events | Where-Object {$_.Properties[0].Value -eq $user.SamAccountName}
        } else {
            $output = $events
        }
        foreach ($event in $output){
            [pscustomobject]@{
                UserName = $event.Properties[0].Value
                CallerComputer = $event.Properties[1].Value
                TimeStamp = $event.TimeCreated
            }
        }
    }
    End{}
}
------------------------------------------------------
"dsquery.exe user ""ou=Benutzer,ou=82,ou=Polizei-NRW-PB-PE-2012,dc=polizei,dc=nrw,dc=de"" -name L110* | DSMOD user -pwdneverexpires no -canchpwd no -mustchpwd no -pwd P2f2aL4!10"


"dsquery.exe user ""ou=Benutzer,ou=81,ou=Polizei-NRW-PB-PE-2012,dc=polizei,dc=nrw,dc=de"" -name L110* | DSMOD user -pwdneverexpires no -canchpwd no -mustchpwd no -pwd P2f2aL4!10"



Neue Abfrage ohne die Hochkommata:

dsquery.exe user "ou=Benutzer,ou=82,ou=Polizei-NRW-PB-PE-2012,dc=polizei,dc=nrw,dc=de" -name L110* | DSMOD user -pwdneverexpires no -canchpwd no -mustchpwd no -pwd P2f2aL4!10



Hier das Script für zusätzliche Entsperrung der Benutzerkonten:

dsquery.exe user "ou=Benutzer,ou=82,ou=Polizei-NRW-PB-PE-2012,dc=polizei,dc=nrw,dc=de" -name L110* | DSMOD user -pwdneverexpires no -canchpwd no -mustchpwd no -pwd P2f2aL4!10 -disabled no -limit 300



#Hier das Script für die User ab L11012* und Entsperrung:

dsquery.exe user "ou=Benutzer,ou=82,ou=Polizei-NRW-PB-PE-2012,dc=polizei,dc=nrw,dc=de" -name L11012* | DSMOD user -pwdneverexpires no -canchpwd no -mustchpwd no -pwd P2f2aL4!10 -disabled no 


#Hier das Sript für die inaktiven User(39 Wochen) ohne Limit:

dsquery.exe user "ou=Benutzer,ou=82,ou=Polizei-NRW-PB-PE-2012,dc=polizei,dc=nrw,dc=de" -name L110* -inactive 39 -limit 700 >"c:\daten\inactiveuser.txt"

dsquery.exe user "ou=Benutzer,ou=81,ou=Polizei-NRW-PB-PE-2012,dc=polizei,dc=nrw,dc=de" -name L110* -inactive 39 -limit 700 >"c:\daten\inactiveuser.txt"


#Hier das Script für das LAFP Neuss für die inaktive Anmeldung:

dsquery.exe user "ou=Benutzer,ou=82,ou=Polizei-NRW-PB-PE-2012,dc=polizei,dc=nrw,dc=de" -name L114* -inactive 39 -limit 700 >"c:\daten\inactiveuser.txt"

dsquery.exe user "ou=Benutzer,ou=81,ou=Polizei-NRW-PB-PE-2012,dc=polizei,dc=nrw,dc=de" -name L114* -inactive 39 -limit 700 >"c:\daten\inactiveuser.txt"
--------------------------------------------------------------------

# Define the path to the file containing SAMAccountNames of expired users
$expiredUsersFilePath = "C:\Daten\test.txt"

# Read SAMAccountNames from the file
$expiredUsers = Get-Content -Path $expiredUsersFilePath

# Set the new LastLogonDate
$newLastLogonDate = Get-Date

    # Set the new LastLogonDate for the user
    Set-ADUser -Identity $expiredUser -Replace LastLogonDate = $newLastLogonDate

    Write-Host "LastLogonDate updated for user: $expiredUser"

Write-Host "LastLogonDate update completed for expired users."
------------------------------------------------------------------------

#####ACHUTNG, funzt nicht!! erwischt falsche kennungen.
##### neuer Ansatztz mit Filter OU 81 & 82



# Input file path containing the list of affected users
$inputFilePath = "C:\Daten\Deaktivierte_L_Kennung_SAM.txt"

# Set the new password
$newPassword = ConvertTo-SecureString -String "P2f7aL4!01" -AsPlainText -Force

# Read SamAccountNames from the file
$samAccountNames = Get-Content -Path $inputFilePath

# Reset password for each user
foreach ($samAccountName in $samAccountNames) {
    try {
        Set-ADAccountPassword -Identity $samAccountName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $newPassword -Force) -ErrorAction Stop
        Write-Host "Password reset successful for user: $samAccountName" -ForegroundColor Green
    } catch {
        Write-Host "Failed to reset password for user: $samAccountName" -ForegroundColor Red
    }
}

Write-Host "Password reset completed for users in the file."


