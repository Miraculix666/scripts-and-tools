# Function to rename existing files with a running number
function Rename-ExistingFile {
    param (
        [string]$filePath
    )
    $counter = 1
    $newFilePath = $filePath -replace '\.txt$', "_OLD_$('{0:D2}' -f $counter).txt"
    while (Test-Path $newFilePath) {
        $counter++
        $newFilePath = $filePath -replace '\.txt$', "_OLD_$('{0:D2}' -f $counter).txt"
    }
    Rename-Item -Path $filePath -NewName $newFilePath
}

# Define the distinguished names of the Organizational Units (OUs)
$ou81 = "OU=81,OU=Polizei-NRW-PB-PE-2012,DC=polizei,DC=nrw,DC=de"
$ou82 = "OU=82,OU=Polizei-NRW-PB-PE-2012,DC=polizei,DC=nrw,DC=de"

# Define the output file paths
$outputFilePath = "C:\Daten\Users_LastLogon_Report.txt"
$samAccountNamesFilePath = "C:\Daten\All_SAMAccountNames.txt"
$expiredUsersFilePath = "C:\Daten\Expired_Users_SAM.txt"

# Check and rename existing files
if (Test-Path $outputFilePath) { Rename-ExistingFile -filePath $outputFilePath }
if (Test-Path $samAccountNamesFilePath) { Rename-ExistingFile -filePath $samAccountNamesFilePath }
if (Test-Path $expiredUsersFilePath) { Rename-ExistingFile -filePath $expiredUsersFilePath }

# Ensure the files are created anew
New-Item -Path $outputFilePath -ItemType File -Force
New-Item -Path $samAccountNamesFilePath -ItemType File -Force
New-Item -Path $expiredUsersFilePath -ItemType File -Force

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
    $allUsers | ForEach-Object { $_.SamAccountName | Out-File -Append -FilePath $samAccountNamesFilePath }

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

    # Output the results to the report file
    $results | Format-Table -AutoSize | Out-File -FilePath $outputFilePath
}

# Define the new password
$newPassword = Read-Host "Please enter the new password for the users" -AsSecureString

# Read SamAccountNames from the file
$samAccountNames = Get-Content -Path $inputFilePath

# Reset password for each user
foreach ($samAccountName in $samAccountNames) {
    try {
        Set-ADAccountPassword -Identity $samAccountName -Reset -NewPassword $newPassword -ErrorAction Stop
        Write-Host "Password reset successful for user: $samAccountName" -ForegroundColor Green
    } catch {
        Write-Host "Failed to reset password for user: $samAccountName" -ForegroundColor Red
    }
}

Write-Host "Password reset completed for users in the file."
