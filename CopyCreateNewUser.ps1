<#
.SYNOPSIS
    Erstellt AD-Benutzer basierend auf einem Vorlagenbenutzer und CSV-Daten.

.DESCRIPTION
    Dieses Skript erstellt neue AD-Benutzer basierend auf einem Vorlagenbenutzer und CSV-Daten.
    Es unterstützt auch das Auslesen der Daten des Vorlagenbenutzers in eine CSV-Datei.

.PARAMETER CsvPath
    Pfad zur CSV-Datei mit den Benutzerdaten.

.PARAMETER TemplateUser
    SAMAccountName des Vorlagenbenutzers.

.PARAMETER ExportTemplateOnly
    Schalter zum Auslesen der Daten des Vorlagenbenutzers in eine CSV-Datei.

.EXAMPLE
    .\CopyCreateNewUser.ps1 -CsvPath "C:\Users.csv" -TemplateUser "TemplateUser"

.EXAMPLE
    .\CopyCreateNewUser.ps1 -TemplateUser "TemplateUser" -ExportTemplateOnly
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [string]$CsvPath,
    [Parameter(Mandatory=$false)]
    [string]$TemplateUser,
    [Parameter(Mandatory=$false)]
    [switch]$ExportTemplateOnly,
    [Parameter(Mandatory=$false)]
    [string]$TargetOU
)

# Importiere das Active Directory-Modul
Import-Module ActiveDirectory

# Funktion zum Erstellen des Log-Verzeichnisses
function Create-LogDirectory {
    $logDir = "C:\ADUserCreationLogs"
    if (!(Test-Path $logDir)) {
        New-Item -ItemType Directory -Force -Path $logDir
    }
    return $logDir
}

# Funktionen für Logging und Ausgabe
function Write-LogMessage {
    param(
        [string]$Message,
        [ValidateSet('Info','Warning','Error','Success')]
        [string]$Type = 'Info'
)
    $logDir = Create-LogDirectory
    $logFile = Join-Path $logDir "ADUserCreation_$(Get-Date -Format 'yyyyMMdd').log"
    
    $colors = @{
        'Info' = 'Cyan'
        'Warning' = 'Yellow'
        'Error' = 'Red'
        'Success' = 'Green'
    }
    
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $logFile -Append
    Write-Verbose $Message
    Write-Host "[$timestamp] " -NoNewline
    Write-Host $Message -ForegroundColor $colors[$Type]
    
    # Logging in Datei
    $logPath = ".\ADUser_Creation_Log.txt"
    "[$timestamp] [$Type] $Message" | Out-File -FilePath $logPath -Append
}

# Funktion zum Exportieren der Template-Daten
function Export-TemplateUserData {
    param(
        [string]$TemplateUser,
        [string]$CsvPath
    )
    
    try {
        $user = Get-ADUser -Identity $TemplateUser -Properties *
        $exportProperties = @(
            'GivenName','Surname','Department','Title',
            'City','Country','Company','Office'
        )
        
        $userData = $user | Select-Object $exportProperties
        $userData | Export-Csv -Path ".\TemplateUser_Export.csv" -NoTypeInformation -Encoding UTF8
        Write-LogMessage "Template-Daten wurden exportiert nach TemplateUser_Export.csv" -Type Success
    }
    catch {
        Write-LogMessage "Fehler beim Exportieren der Template-Daten: $_" -Type Error
        exit
    }
}

# Funktion zum Erstellen eines neuen AD-Benutzers
function New-ADUserFromTemplate {
    param(
        [hashtable]$UserData,
        [string]$TemplateUser,
        [string]$TargetOU
    )
    $template = Get-ADUser -Identity $TemplateUser -Properties *
    
$securePassword = ConvertTo-SecureString $UserData.Password -AsPlainText -Force

    
$newUserParams = @{
        SamAccountName = $UserData.SamAccountName
        UserPrincipalName = "$($UserData.SamAccountName)@$((Get-ADDomain).DNSRoot)"
        Name = "$($UserData.GivenName) $($UserData.Surname)"
        GivenName = $UserData.GivenName
        Surname = $UserData.Surname
        DisplayName = $UserData.DisplayName
        Description = $UserData.Description
        Office = $UserData.Office
        Department = $UserData.Department
        Title = $UserData.Title
        Company = $UserData.Company
        EmailAddress = $UserData.EmailAddress
        StreetAddress = $UserData.StreetAddress
        City = $UserData.City
        State = $UserData.State
        PostalCode = $UserData.PostalCode
        Country = $UserData.Country
        OfficePhone = $UserData.OfficePhone
AccountPassword = $securePassword
Enabled = $true
        Instance = $template
        Path = $TargetOU
}

    
    # Optionale Parameter hinzufügen wenn vorhanden
    @('Department','Title','City','Country','Company','Office') | ForEach-Object {
        if ($UserData.$_) {
            $newUserParams[$_] = $UserData.$_
        }
    }
    
try {
New-ADUser @newUserParams
        Write-LogMessage "Benutzer $($UserData.SamAccountName) erfolgreich erstellt" -Type Success
}
catch {
        Write-LogMessage "Fehler beim Erstellen von $($UserData.SamAccountName): $_" -Type Error
}
}

if ($MyInvocation.InvocationName -ne '.') {
# Hauptprogramm
Write-LogMessage "Starte AD-Benutzerverarbeitung" -Type Info

# OU-Validierung
if (-not $TargetOU) {
    $TargetOU = Read-Host "Bitte geben Sie die Ziel-OU an (z.B. 'OU=Users,DC=domain,DC=com')"
}
else {
    if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$TargetOU'")) {
        Write-LogMessage "Die angegebene OU existiert nicht: $TargetOU" -Type Error
        return
    }
}

# Template-Export wenn gewünscht
if ($ExportTemplateOnly -and $TemplateUser) {
    Export-TemplateUserData -TemplateUser $TemplateUser -CsvPath $CsvPath
    return
}

# Verarbeitungsmodus bestimmen
if ($CsvPath) {
    # CSV-Modus
    $users = Import-Csv -Path $CsvPath -Encoding UTF8
    foreach ($user in $users) {
        if (-not $user.Password) {
            Write-LogMessage "Kein Passwort für Benutzer $($user.SamAccountName) angegeben" -Type Error
            continue
        }
        New-ADUserFromTemplate -UserData $user -TemplateUser $TemplateUser -TargetOU $TargetOU
    }
}
else {
    # Interaktiver Modus
    $userData = @{}
    $userData.GivenName = Read-Host "Vorname"
    $userData.Surname = Read-Host "Nachname"
    $userData.SamAccountName = Read-Host "SAM Account Name"
    $userData.Password = Read-Host "Passwort"
    
    if ($TemplateUser) {
        $template = Get-ADUser -Identity $TemplateUser -Properties *
        @('Department','Title','City','Country','Company','Office') | ForEach-Object {
            $userData[$_] = $template.$_
        }
    }
    
    New-ADUserFromTemplate -UserData $userData -TemplateUser $TemplateUser -TargetOU $TargetOU
}

Write-LogMessage "Verarbeitung abgeschlossen" -Type Success
}
