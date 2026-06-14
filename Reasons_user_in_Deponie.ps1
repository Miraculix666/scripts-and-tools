
### This script looksup after deactiveted L-Kennungen in "Deponie" an checks the reasons with .\Get-ADUserLockouts.ps1 from https://github.com/ThePoShWolf/Utilities/blob/master/ActiveDirectory/Get-ADUserLockouts.ps1

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
$allDeactivatedUsers | ForEach-Object { $_.SamAccountName } | Out-File -Append -FilePath $outputFilePathSAM
# Display message about the file
Write-Host "SamAccountNames have been written to: $outputFilePathSAM"




$allDeactivatedUsers | Select-Object SamAccountName | .\Get-ADUserLockouts.ps1
