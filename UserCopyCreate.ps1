# Active Directory User Management Script
# Author: Enhanced by Bolt
# Version: 2.4
# PowerShell Version: 5.1
# Description:
#   - Exports existing AD users to CSV based on a template user's OUs
#   - Creates new users interactively, via parameters, or from CSV
#   - Assigns groups based on template user
#   - Supports comprehensive logging and verbose output
#   - Implements German localization

[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $false, HelpMessage = "Template user for group assignments")]
    [string]$TemplateUser,

    [Parameter(Mandatory = $false, HelpMessage = "Path for exporting user data (CSV)")]
    [string]$ExportPath = "ADBenutzerExport.csv",

    [Parameter(Mandatory = $false, HelpMessage = "Path to CSV file for user import")]
    [string]$CsvPath,

    [Parameter(Mandatory = $false, HelpMessage = "Log file path")]
    [string]$LogPath = "ADBenutzerVerwaltung.log",

    [Parameter(Mandatory = $false, HelpMessage = "Default password for new users")]
    [SecureString]$DefaultPassword = (ConvertTo-SecureString "Willkommen2024!" -AsPlainText -Force),

    [Parameter(Mandatory = $false, HelpMessage = "Default OU for new users")]
    [string]$DefaultOU,

    [Parameter(Mandatory = $false, HelpMessage = "Export only template user")]
    [switch]$ExportTemplateOnly,

    [Parameter(Mandatory = $false, HelpMessage = "Enable verbose output")]
    [switch]$Verbose
)

# Set German culture for proper date/number formatting
$previousCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo("de-DE")

# Import required modules with error handling
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Verbose "ActiveDirectory-Modul erfolgreich geladen"
} catch {
    Write-Error "Fehler beim Laden des ActiveDirectory-Moduls: $_"
    exit 1
}

# Enhanced logging function with severity levels
function Write-CustomLog {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNUNG", "FEHLER", "DEBUG")]
        [string]$Level = "INFO"
    )

    $Timestamp = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    $LogMessage = "$Timestamp [$Level] $Message"

    # Write to log file
    Add-Content -Path $LogPath -Value $LogMessage -Encoding UTF8

    # Console output with color coding
    $Color = switch ($Level) {
        "INFO"    { "White" }
        "WARNUNG" { "Yellow" }
        "FEHLER"  { "Red" }
        "DEBUG"   { "Cyan" }
    }
    Write-Host $LogMessage -ForegroundColor $Color

    # Additional verbose output for debugging
    if ($Level -eq "DEBUG") {
        Write-Verbose $Message
    }
}

# Function to validate and get AD user with extended properties
function Get-ValidADUser {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Identity,
        [string]$Operation = "Abfrage"
    )

    try {
        $user = Get-ADUser -Identity $Identity -Properties MemberOf, DistinguishedName, GivenName, Surname, Department, Title, Manager, Office, OfficePhone, EmailAddress, Company, Description -ErrorAction Stop
        Write-Verbose "Benutzer '$Identity' erfolgreich gefunden"
        return $user
    } catch {
        Write-CustomLog "Fehler bei $Operation für Benutzer '$Identity': $_" -Level "FEHLER"
        return $null
    }
}

# Function to get all OUs from template user's path
function Get-TemplateUserOUs {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TemplateUser
    )

    $user = Get-ValidADUser -Identity $TemplateUser -Operation "OU-Ermittlung"
    if ($user) {
        try {
            $ouPath = ($user.DistinguishedName -split ',', 2)[1]
            Write-Verbose "Base OU path: $ouPath"

            # Get all OUs from the path
            $ous = @()
            $ous += Get-ADOrganizationalUnit -Filter * -SearchBase $ouPath -SearchScope Subtree |
                   Select-Object -ExpandProperty DistinguishedName

            # Add the template user's direct OU
            $ous += $ouPath

            $uniqueOUs = $ous | Select-Object -Unique
            Write-Verbose "Gefundene OUs: $($uniqueOUs.Count)"
            Write-Verbose ($uniqueOUs -join "`n")

            return $uniqueOUs
        } catch {
            Write-CustomLog "Fehler beim Abrufen der OUs: $_" -Level "FEHLER"
            return @($ouPath)
        }
    }
    return $null
}

# Function to compare group memberships
function Compare-GroupMembership {
    param (
        [string[]]$TemplateGroups,
        [string[]]$UserGroups
    )

    if (-not $TemplateGroups -or -not $UserGroups) {
        return $false
    }

    $templateGroupNames = $TemplateGroups | ForEach-Object { $_.Split(',')[0] }
    $userGroupNames = $UserGroups | ForEach-Object { $_.Split(',')[0] }

    $commonGroups = Compare-Object -ReferenceObject $templateGroupNames -DifferenceObject $userGroupNames -IncludeEqual -ExcludeDifferent
    return ($commonGroups.Count -gt 0)
}

# Function to export AD users
function Export-ADUsers {
    [CmdletBinding()]
    param()

    Write-CustomLog "Starte Benutzerexport basierend auf Template: $TemplateUser" -Level "INFO"

    $Template = Get-ValidADUser -Identity $TemplateUser -Operation "Template-Validierung"
    if (-not $Template) { return }

    try {
        Write-Verbose "Hole Template-Benutzer Gruppen und OUs"
        $templateGroups = $Template.MemberOf
        Write-Verbose "Template Gruppen: $($templateGroups.Count)"

        $templateOUs = Get-TemplateUserOUs -TemplateUser $TemplateUser
        if (-not $templateOUs) {
            throw "Keine OUs für Template-Benutzer gefunden"
        }

        Write-Verbose "Suche Benutzer in allen relevanten OUs"
        $allUsers = @()
        if ($ExportTemplateOnly) {
            # Export only the template user
            $allUsers += $Template
        } else {
            foreach ($ou in $templateOUs) {
                Write-Verbose "Durchsuche OU: $ou"
                try {
                    $usersInOU = Get-ADUser -Filter * -SearchBase $ou -SearchScope OneLevel `
                                -Properties SamAccountName, UserPrincipalName, Name, GivenName, Surname, `
                                          Department, Title, Manager, Office, OfficePhone, EmailAddress, `
                                          Company, Description, MemberOf, DistinguishedName

                    foreach ($user in $usersInOU) {
                        if (Compare-GroupMembership -TemplateGroups $templateGroups -UserGroups $user.MemberOf) {
                            $allUsers += $user
                        }
                    }
                } catch {
                    Write-CustomLog "Fehler beim Durchsuchen von OU '$ou': $_" -Level "WARNUNG"
                    continue
                }
            }
        }

        Write-Verbose "Gefundene Benutzer: $($allUsers.Count)"

        if ($allUsers.Count -eq 0) {
            Write-CustomLog "Keine Benutzer mit übereinstimmenden Gruppen gefunden" -Level "WARNUNG"
            return
        }

        $exportData = $allUsers | Select-Object @{
            Name='Benutzername'; Expression={$_.SamAccountName}
        }, @{
            Name='E-Mail'; Expression={$_.EmailAddress}
        }, @{
            Name='Vorname'; Expression={$_.GivenName}
        }, @{
            Name='Nachname'; Expression={$_.Surname}
        }, @{
            Name='Abteilung'; Expression={$_.Department}
        }, @{
            Name='Position'; Expression={$_.Title}
        }, @{
            Name='Vorgesetzter'; Expression={
                if ($_.Manager) {
                    try {
                        (Get-ADUser $_.Manager).SamAccountName
                    } catch {
                        $_.Manager
                    }
                } else { "" }
            }
        }, @{
            Name='Büro'; Expression={$_.Office}
        }, @{
            Name='Telefon'; Expression={$_.OfficePhone}
        }, @{
            Name='Firma'; Expression={$_.Company}
        }, @{
            Name='Beschreibung'; Expression={$_.Description}
        }, @{
            Name='OU'; Expression={($_.DistinguishedName -split ',',2)[1]}
        }, @{
            Name='Gruppen'; Expression={$_.MemberOf -join ';' }
        }

        $exportData | Export-Csv -Path $ExportPath -Delimiter ";" -NoTypeInformation -Encoding UTF8
        Write-CustomLog "Benutzerdaten erfolgreich exportiert nach: $ExportPath" -Level "INFO"
        Write-CustomLog "Anzahl exportierter Benutzer: $($exportData.Count)" -Level "INFO"
    } catch {
        Write-CustomLog "Fehler beim Exportieren der Benutzerdaten: $_" -Level "FEHLER"
    }
}

# Function to create new AD user
function New-CustomADUser {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true)]
        [string]$SamAccountName,
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $false)]
        [string]$GivenName,
        [Parameter(Mandatory = $false)]
        [string]$Surname,
        [Parameter(Mandatory = $false)]
        [string]$Department,
        [Parameter(Mandatory = $false)]
        [string]$Title,
        [Parameter(Mandatory = $false)]
        [string]$Manager,
        [Parameter(Mandatory = $false)]
        [string]$Office,
        [Parameter(Mandatory = $false)]
        [string]$OfficePhone,
        [Parameter(Mandatory = $false)]
        [string]$Company,
        [Parameter(Mandatory = $false)]
        [string]$Description,
        [Parameter(Mandatory = $true)]
        [string]$OU,
        [string[]]$Groups,
        [SecureString]$Password = $DefaultPassword
    )

    Write-Verbose "Erstelle neuen Benutzer: $Name in OU: $OU"

    try {
        if ($PSCmdlet.ShouldProcess($Name, "Benutzer erstellen")) {
            # Validate OU exists
            if (-not (Get-ADOrganizationalUnit -Filter {DistinguishedName -eq $OU})) {
                throw "Die angegebene OU existiert nicht: $OU"
            }

            $userParams = @{
                Name = $Name
                SamAccountName = $SamAccountName
                UserPrincipalName = $UserPrincipalName
                Path = $OU
                AccountPassword = $Password
                Enabled = $true
                ChangePasswordAtLogon = $true
            }

            if ($GivenName) { $userParams.GivenName = $GivenName }
            if ($Surname) { $userParams.Surname = $Surname }
            if ($Department) { $userParams.Department = $Department }
            if ($Title) { $userParams.Title = $Title }
            if ($Manager) { $userParams.Manager = $Manager }
            if ($Office) { $userParams.Office = $Office }
            if ($OfficePhone) { $userParams.OfficePhone = $OfficePhone }
            if ($Company) { $userParams.Company = $Company }
            if ($Description) { $userParams.Description = $Description }

            New-ADUser @userParams -ErrorAction Stop
            Write-CustomLog "Benutzer '$Name' ($SamAccountName) erfolgreich erstellt." -Level "INFO"

            if ($Groups) {
                foreach ($group in $Groups) {
                    try {
                        Add-ADGroupMember -Identity $group -Members $SamAccountName -ErrorAction Stop
                        Write-Verbose "Benutzer zur Gruppe '$group' hinzugefügt."
                    } catch {
                        Write-CustomLog "Fehler beim Hinzufügen von Benutzer '$SamAccountName' zu Gruppe '$group': $_" -Level "WARNUNG"
                    }
                }
            }
        }
    } catch {
        Write-CustomLog "Fehler bei der Erstellung des Benutzers '$Name': $_" -Level "FEHLER"
        throw
    }
}
