<#
.SYNOPSIS
Verwaltet Active Directory-Benutzer: Kopiert einzelne Benutzer, erstellt Benutzer aus CSV, wendet Eigenschaften an oder exportiert Benutzerdaten (allgemein oder spezifisch für L-Kennung).

.DESCRIPTION
Dieses Skript bietet Werkzeuge zur Verwaltung von Active Directory (AD) Benutzern.
Es kann verwendet werden, um:
- Einen einzelnen AD-Benutzer zu kopieren, einschließlich Gruppenzugehörigkeiten und optional der OU-Struktur (Modus: CopySingleUser).
- Mehrere AD-Benutzer basierend auf einer CSV-Datei zu erstellen (Modus: CreateUsersFromCSV). Optional können Attribute und Gruppen von einem Referenzbenutzer (Vorlage) übernommen werden.
- Ausgewählte Eigenschaften und Gruppenmitgliedschaften von einem Referenzbenutzer auf einen bereits existierenden Zielbenutzer anzuwenden (Modus: ApplyPropertiesToExistingUser).
- Benutzerdaten basierend auf Identitäts- und/oder OU-Filtern zu suchen und relevante Eigenschaften (inkl. OU, Gruppen) in eine CSV-Datei zu exportieren (Modus: ExportUserData).
- Spezifisch Benutzer mit L-Kennung (L110*/L114*) aus definierten OUs (standardmäßig '81', '82') zu exportieren (Modus: ExportLKennung).

Das Skript ist für PowerShell Version 5.1 optimiert und verwendet deutsche Lokalisierungseinstellungen für CSV-Exporte und Datums-/Zeitformate.
Es implementiert detaillierte Protokollierung und gibt standardmäßig ausführliche Meldungen aus (Verbose Output).
Log-Dateien und ein Benutzerbericht (CSV) über durchgeführte Aktionen werden standardmäßig im Verzeichnis des Skripts gespeichert.
Vor dem Erstellen oder Ändern von Benutzern wird eine detaillierte Bestätigungsaufforderung angezeigt (es sei denn, -Confirm:$false wird verwendet).

.PARAMETER CopySingleUser
Schalter, um den Modus zum Kopieren eines einzelnen Benutzers zu aktivieren.

.PARAMETER CreateUsersFromCSV
Schalter, um den Modus zur Erstellung von Benutzern aus einer CSV-Datei zu aktivieren.

.PARAMETER ApplyPropertiesToExistingUser
Schalter, um den Modus zum Anwenden von Eigenschaften/Gruppen auf einen existierenden Benutzer zu aktivieren.

.PARAMETER ExportUserData
Schalter, um den Modus zum Exportieren von Benutzerdaten mit allgemeinen Filtern zu aktivieren.

.PARAMETER ExportLKennung
Schalter, um den Modus zum Exportieren spezifischer L-Kennung-Benutzer aus definierten OUs zu aktivieren.

.PARAMETER ReferenceUserSamAccountName
Der SAMAccountName des Referenzbenutzers.
- Im Modus 'CopySingleUser': Der Quellbenutzer, der kopiert werden soll (Mandatory).
- Im Modus 'CreateUsersFromCSV': Der Vorlagenbenutzer, von dem optional Attribute/Gruppen übernommen werden (Optional).
- Im Modus 'ApplyPropertiesToExistingUser': Der Benutzer, dessen Eigenschaften/Gruppen als Quelle dienen (Mandatory).

.PARAMETER TargetUserSamAccountName
Der SAMAccountName des Zielbenutzers.
- Im Modus 'CopySingleUser': Der gewünschte SAMAccountName für den *neuen* (kopierten) Benutzer (Optional, wird interaktiv abgefragt).
- Im Modus 'ApplyPropertiesToExistingUser': Der SAMAccountName des *existierenden* Benutzers, der modifiziert werden soll (Mandatory).

.PARAMETER TargetUserPassword
Das initiale Passwort für den neuen (kopierten) Benutzer als SecureString (nur für Modus CopySingleUser). Wird interaktiv und sicher abgefragt, wenn nicht angegeben. Der Benutzer muss das Passwort bei der ersten Anmeldung ändern.

.PARAMETER TargetOU
Der Distinguished Name der Organisationseinheit (OU), in die der neue Benutzer kopiert/erstellt werden soll.
Für CopySingleUser: Optional; Standard ist die OU des Referenzbenutzers.
Für CreateUsersFromCSV: Optional; Überschreibt die OU aus der CSV oder vom Referenzbenutzer.
Für ApplyPropertiesToExistingUser / Export*: Nicht relevant.

.PARAMETER Force
Überschreibt einen bereits existierenden Zielbenutzer im Modus CopySingleUser, falls vorhanden. Standardmäßig wird ein Fehler ausgegeben.

.PARAMETER CsvPath
Der vollständige Pfad zur CSV-Datei, die die Daten für die zu erstellenden Benutzer enthält (nur für Modus CreateUsersFromCSV).
Erwartete Spalten (mindestens): SamAccountName, GivenName, Surname.
Optionale Spalten: Password (Klartext - NICHT EMPFOHLEN!), TargetOU, Enabled, Description, Office, Department, Title, Company, EmailAddress, StreetAddress, City, State, PostalCode, Country, OfficePhone. Siehe .NOTES für Details.

.PARAMETER DefaultPassword
Ein Standardpasswort (als SecureString), das für alle Benutzer aus der CSV verwendet wird, es sei denn, die CSV enthält eine 'Password'-Spalte. Wenn weder DefaultPassword noch eine 'Password'-Spalte vorhanden sind, wird ein zufälliges Passwort generiert (empfohlen).

.PARAMETER IdentityFilter
Ein Identitätsfilter (SamAccountName, Name, etc.) mit optionalen Wildcards (*), um die zu exportierenden Benutzer zu definieren (nur für Modus ExportUserData). Beispiel: 'Hans*', 'Benutzer1', '*Mueller'. Mandatory für ExportUserData.

.PARAMETER OUFilter
Ein Filter für den OU-Namen mit optionalen Wildcards (*), um die Suche für ExportUserData auf bestimmte OUs einzuschränken. Beispiel: '*Manag*', 'Vertrieb', '81'. Optional.

.PARAMETER SearchBaseOU
Der Distinguished Name einer spezifischen OU, um die Benutzersuche für den ExportUserData *alternativ* zu -OUFilter einzuschränken. Optional. (Hinweis: -OUFilter hat Vorrang, wenn beides angegeben wird).

.PARAMETER PropertiesToExport
Eine Liste von AD-Attributnamen (String-Array), die zusätzlich zu den Standardattributen exportiert werden sollen (für Modi ExportUserData und ExportLKennung). Standard: siehe .NOTES.

.PARAMETER ExportCsvPath
Pfad für die CSV-Exportdatei (nur für Modus ExportUserData). Mandatory.

.PARAMETER LKennungOUNames
Die Namen der OUs, die für den L-Kennung-Export durchsucht werden sollen (nur für Modus ExportLKennung). Standard: @('81', '82').

.PARAMETER LKennungLDAPFilter
Der LDAP-Filter, der für den L-Kennung-Export verwendet wird (nur für Modus ExportLKennung). Standard: "(|(sAMAccountName=L110*)(sAMAccountName=L114*))".

.PARAMETER LKennungExportCsvPath
Pfad für die CSV-Exportdatei im L-Kennung-Exportmodus (nur für Modus ExportLKennung). Mandatory.

.PARAMETER LogPath
Verzeichnis für die Log-Dateien. Standard ist das Verzeichnis, in dem das Skript liegt (`$PSScriptRoot`), oder das aktuelle Arbeitsverzeichnis (`$PWD`), wenn nicht aus einer Datei ausgeführt.

.PARAMETER LogLevel
Steuert die Detailtiefe der Log-Datei. Mögliche Werte: Error, Warning, Info, Verbose. Standard ist 'Info'.

.PARAMETER UserReportCsvPath
Pfad für die CSV-Datei, die einen Bericht über erstellte/kopierte/modifizierte Benutzer enthält (nicht für Export-Modi relevant). Standard ist das Skriptverzeichnis (`$PSScriptRoot` oder `$PWD`) mit dem Namen '{ScriptName}_UserReport_{Timestamp}.csv'. Der Bericht wird immer für die Modi Copy, Create, Apply erstellt.

.EXAMPLE
# Beispiel 1: Kopiert 'BenutzerA' zu 'BenutzerB' interaktiv (mit Verbose-Ausgabe und Bestätigung)
.\Enhanced-ADManagement.ps1 -CopySingleUser -ReferenceUserSamAccountName BenutzerA

.EXAMPLE
# Beispiel 2: Kopiert 'BenutzerA' zu 'BenutzerC', setzt Passwort, legt in spezifischer OU ab und überschreibt Ziel falls vorhanden (mit Bestätigung)
$password = ConvertTo-SecureString "P@sswOrd123!" -AsPlainText -Force
.\Enhanced-ADManagement.ps1 -CopySingleUser -ReferenceUserSamAccountName BenutzerA -TargetUserSamAccountName BenutzerC -TargetUserPassword $password -TargetOU "OU=NeueMitarbeiter,DC=firma,DC=local" -Force

.EXAMPLE
# Beispiel 3: Erstellt Benutzer aus CSV mit Standardpasswort und Gruppen/Attributen von 'TemplateUser' (mit Bestätigung für jeden Benutzer)
$defaultPass = ConvertTo-SecureString "Sommer2025!" -AsPlainText -Force
.\Enhanced-ADManagement.ps1 -CreateUsersFromCSV -CsvPath "C:\temp\neue_benutzer.csv" -ReferenceUserSamAccountName TemplateUser -DefaultPassword $defaultPass -LogLevel Info

.EXAMPLE
# Beispiel 4: Erstellt Benutzer aus CSV, verwendet Passwort aus CSV (WARNUNG: Unsicher!) und speichert Log in C:\Logs (mit Bestätigung)
.\Enhanced-ADManagement.ps1 -CreateUsersFromCSV -CsvPath "C:\temp\neue_benutzer_mit_passwort.csv" -TargetOU "OU=Vertrieb,DC=firma,DC=local" -LogPath "C:\Logs"

.EXAMPLE
# Beispiel 5: Wendet Abteilungs-, Büro- und Gruppeninformationen von 'RefUser' auf den existierenden Benutzer 'ExistingUser' an (mit Bestätigung)
.\Enhanced-ADManagement.ps1 -ApplyPropertiesToExistingUser -ReferenceUserSamAccountName RefUser -TargetUserSamAccountName ExistingUser

.EXAMPLE
# Beispiel 6: Exportiert alle Benutzer, deren SamAccountName mit 'L' beginnt (Allgemeiner Export)
.\Enhanced-ADManagement.ps1 -ExportUserData -IdentityFilter "L*" -ExportCsvPath "C:\temp\L_Benutzer_Export.csv"

.EXAMPLE
# Beispiel 7: Exportiert alle Benutzer, deren Name 'Hans' enthält, aus OUs, die 'Management' im Namen haben (Allgemeiner Export)
.\Enhanced-ADManagement.ps1 -ExportUserData -IdentityFilter "*Hans*" -OUFilter "*Management*" -ExportCsvPath "C:\temp\Hans_Management_Export.csv"

.EXAMPLE
# Beispiel 8: Exportiert spezifische L-Kennung-Benutzer (L110*/L114*) aus den Standard-OUs ('81', '82')
.\Enhanced-ADManagement.ps1 -ExportLKennung -LKennungExportCsvPath "C:\temp\Export_L-Kennung_Standard.csv"

.EXAMPLE
# Beispiel 9: Exportiert spezifische L-Kennung-Benutzer aus einer anderen OU ('99') mit Standardfilter
.\Enhanced-ADManagement.ps1 -ExportLKennung -LKennungOUNames @('99') -LKennungExportCsvPath "C:\temp\Export_L-Kennung_OU99.csv"

.NOTES
Autor: Gemini (basierend auf Nutzer-Input und Beispielen)
Version: 6.7
Datum: 2025-05-06
Benötigte Module: ActiveDirectory (wird durch #requires geprüft)
Benötigte Berechtigungen: Ausreichende AD-Berechtigungen zum Lesen von Benutzern und zum Erstellen/Modifizieren von Benutzern. Schreibrechte im Zielverzeichnis für Logs/Berichte/Exporte.

Bestätigungsaufforderungen:
- Das Skript verwendet die PowerShell-Standardmechanismen für Bestätigungen (`SupportsShouldProcess`).
- Vor jeder Aktion, die einen Benutzer erstellt oder modifiziert (Copy, Create, Apply), wird eine detaillierte Meldung angezeigt, was getan wird.
- Standardmäßig fragt PowerShell nach Bestätigung (J/N/A/H).
- Um alle Bestätigungen automatisch zu überspringen, verwenden Sie den Parameter `-Confirm:$false`.
- Um nur zu sehen, was getan würde, ohne Änderungen vorzunehmen, verwenden Sie den Parameter `-WhatIf`.

Modus ApplyPropertiesToExistingUser:
- Kopiert standardmäßig folgende Eigenschaften: Description, Office, StreetAddress, City, State, PostalCode, Country, Department, Company, Title, OfficePhone, EmailAddress.
- Fügt Gruppenmitgliedschaften hinzu, die der Referenzbenutzer hat, der Zielbenutzer aber nicht (ausgenommen 'Domain Users'). Bestehende Gruppen des Zielbenutzers bleiben erhalten.
- Ändert KEINE sicherheitsrelevanten oder Identitäts-Attribute wie Passwort, SID, SamAccountName, UPN, Enabled-Status.

Modi ExportUserData / ExportLKennung:
- Standardmäßig exportierte Eigenschaften: SamAccountName, Name, GivenName, Surname, DisplayName, UserPrincipalName, Enabled, DistinguishedName, OU (extrahiert), GroupNames (kommasepariert).
- Zusätzliche Eigenschaften können mit -PropertiesToExport angegeben werden. Immer mit exportiert werden 'MemberOf' (zur Gruppenauflösung) und 'DistinguishedName' (zur OU-Extraktion).
- Der Export erfolgt als CSV mit Semikolon als Trennzeichen und UTF8-Kodierung.
- Der Modus ExportLKennung sucht standardmäßig in OUs mit Namen '81' und '82' nach Benutzern mit SamAccountName 'L110*' oder 'L114*'. Dies kann über -LKennungOUNames und -LKennungLDAPFilter angepasst werden.
- Der Modus ExportUserData verwendet -IdentityFilter für die Benutzersuche (SamAccountName, Name etc.) und optional -OUFilter oder -SearchBaseOU, um den Suchbereich einzuschränken.

CSV-Format für CreateUsersFromCSV:
- Trennzeichen: Semikolon (;)
- Kodierung: UTF8
- Erforderliche Spalten: SamAccountName, GivenName (Vorname), Surname (Nachname)
- Empfohlene Spalten für volle Funktionalität: EmailAddress, TargetOU (Distinguished Name)
- Optionale Spalten (werden verwendet, wenn vorhanden): Password (Klartext - NICHT EMPFOHLEN!), Enabled (true/false), Description, Office, Department, Title, Company, StreetAddress, City, State, PostalCode, Country, OfficePhone.
- Wenn die Spalte 'Password' nicht vorhanden ist oder leer ist, wird -DefaultPassword verwendet. Wenn beides fehlt, wird ein sicheres, zufälliges Passwort generiert und MUSS geändert werden.
- Wenn die Spalte 'TargetOU' nicht vorhanden ist oder leer ist, wird der Wert von -TargetOU verwendet. Wenn dieser auch fehlt, wird die OU des ReferenceUser (falls angegeben) verwendet. Ansonsten Fehler.
- Wenn die Spalte 'Enabled' nicht vorhanden ist, wird der Benutzer standardmäßig aktiviert ($true).

Passwort-Sicherheit: Das Speichern von Klartext-Passwörtern in CSV-Dateien ist ein erhebliches Sicherheitsrisiko! Verwenden Sie bevorzugt den Parameter -DefaultPassword oder generieren Sie zufällige Passwörter und kommunizieren Sie diese sicher.

Standardpfade: Wenn -LogPath, -UserReportCsvPath, -ExportCsvPath oder -LKennungExportCsvPath nicht angegeben werden, versucht das Skript, die Dateien im selben Verzeichnis wie das Skript selbst (`$PSScriptRoot`) zu speichern. Wenn das Skript nicht aus einer Datei ausgeführt wird (z.B. im ISE oder direkt in der Konsole), wird stattdessen das aktuelle Arbeitsverzeichnis (`$PWD`) verwendet. Stellen Sie sicher, dass Schreibberechtigungen im Zielverzeichnis vorhanden sind.

Benutzerbericht: Der CSV-Bericht über durchgeführte Aktionen (Copy, Create, Apply) wird immer erstellt und im Verzeichnis gespeichert, das durch -UserReportCsvPath festgelegt ist (oder im Standardpfad).

Testen: Führen Sie das Skript zuerst in einer Testumgebung aus!

Quellen/Referenzen (vom Benutzer bereitgestellt):
https://petri.com/how-to-copy-active-directory-groups-from-one-user-to-another-with-powershell/
https://petri.com/how-to-copy-active-directory-users-with-powershell/
https://petri.com/create-new-active-directory-users-excel-powershell
https://blog.netwrix.com/bulk-user-creation-in-active-directory/
https://learn.microsoft.com/en-us/answers/questions/1035531/creating-new-ad-users-in-powershell-from-existing?page=2#answers
https://github.com/RichPrescott/UserCreation/blob/master/ANUC.ps1
https://community.spiceworks.com/t/bulk-create-active-directory-users-powershell-with-logs-less-rows-in-csv/974593
https://forums.powershell.org/t/copy-ad-user-not-from-template-just-mirror/18842/8
https://support.microsoft.com/en-us/windows/manage-user-accounts-in-windows-104dc19f-6430-4b49-6a2b-e4dbd1dcdf32
https://activedirectorypro.com/copy-group-membership-from-one-user-to-another-in-ad/
https://community.spiceworks.com/t/powershell-add-users-error/609976
https://community.spiceworks.com/t/how-to-copy-group-membership-from-one-user-to-another/1014279
https://notebooklm.google.com/notebook/9f6821c0-f64d-4adf-a462-c68fd050aea4?_gl=1*1p5hmk1*_ga*MTkyNjMzNzExNC4xNzM4NzYzMjU2*_ga_W0LDH41ZCB*MTc0NjE4MDcwMi4xMC4xLjE3NDYxODA3MDIuMC4wLjA.

.COMPONENT
ActiveDirectory

.ROLE
Administrator

.FUNCTIONALITY
User Account Management
User Data Export
#>

#requires -Version 5.1
#requires -Modules ActiveDirectory

[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'CopySingleUser')]
param(
    # --- Modus Schalter ---
    [Parameter(ParameterSetName = 'CopySingleUser', Mandatory = $true, HelpMessage = "Aktiviert den Modus zum Kopieren eines einzelnen Benutzers.")]
    [switch]$CopySingleUser,

    [Parameter(ParameterSetName = 'CreateUsersFromCSV', Mandatory = $true, HelpMessage = "Aktiviert den Modus zur Benutzererstellung aus CSV.")]
    [switch]$CreateUsersFromCSV,

    [Parameter(ParameterSetName = 'ApplyPropertiesToExistingUser', Mandatory = $true, HelpMessage = "Aktiviert Modus zum Anwenden von Eigenschaften auf existierenden Benutzer.")]
    [switch]$ApplyPropertiesToExistingUser,

    [Parameter(ParameterSetName = 'ExportUserData', Mandatory = $true, HelpMessage = "Aktiviert Modus zum Exportieren von Benutzerdaten mit allgemeinen Filtern.")]
    [switch]$ExportUserData,

    [Parameter(ParameterSetName = 'ExportLKennung', Mandatory = $true, HelpMessage = "Aktiviert Modus zum Exportieren spezifischer L-Kennung-Benutzer.")]
    [switch]$ExportLKennung,

    # --- Gemeinsame Parameter für Copy, Create, Apply ---
    [Parameter(ParameterSetName = 'CopySingleUser', Mandatory = $true, HelpMessage = "SAMAccountName des Quellbenutzers für die Kopie.")]
    [Parameter(ParameterSetName = 'CreateUsersFromCSV', Mandatory = $false, HelpMessage = "SAMAccountName des Vorlagenbenutzers für Attribute/Gruppen.")]
    [Parameter(ParameterSetName = 'ApplyPropertiesToExistingUser', Mandatory = $true, HelpMessage = "SAMAccountName des Benutzers, dessen Eigenschaften/Gruppen als Quelle dienen.")]
    [ValidateNotNullOrEmpty()]
    [string]$ReferenceUserSamAccountName,

    [Parameter(ParameterSetName = 'CopySingleUser', Mandatory = $false, HelpMessage = "Gewünschter SAMAccountName für den NEUEN Benutzer.")]
    [Parameter(ParameterSetName = 'ApplyPropertiesToExistingUser', Mandatory = $true, HelpMessage = "SAMAccountName des EXISTIERENDEN Benutzers, der modifiziert wird.")]
    [ValidateNotNullOrEmpty()]
    [string]$TargetUserSamAccountName, # Gilt für CopySingleUser (optional) und ApplyProperties (mandatory)

    [Parameter(ParameterSetName = 'CopySingleUser')]
    [Parameter(ParameterSetName = 'CreateUsersFromCSV')]
    [ValidateNotNullOrEmpty()]
    [string]$TargetOU, # Gilt für CopySingleUser und CreateUsersFromCSV

    # --- Parameter für CopySingleUser ---
    [Parameter(ParameterSetName = 'CopySingleUser', Mandatory = $false, HelpMessage = "Initiales Passwort für den neuen Benutzer (SecureString).")]
    [System.Security.SecureString]$TargetUserPassword,

    [Parameter(ParameterSetName = 'CopySingleUser', HelpMessage = "Überschreibt einen existierenden Zielbenutzer.")]
    [switch]$Force,

    # --- Parameter für CreateUsersFromCSV ---
    [Parameter(ParameterSetName = 'CreateUsersFromCSV', Mandatory = $true, HelpMessage = "Pfad zur CSV-Datei.")]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [Parameter(ParameterSetName = 'CreateUsersFromCSV', Mandatory = $false, HelpMessage = "Standardpasswort (SecureString) für CSV-Benutzer.")]
    [System.Security.SecureString]$DefaultPassword,

    # --- Parameter für ExportUserData ---
    [Parameter(ParameterSetName = 'ExportUserData', Mandatory = $true, HelpMessage = "Identitätsfilter (SamAccountName, Name etc.) mit optionalen Wildcards (*).")]
    [ValidateNotNullOrEmpty()]
    [string]$IdentityFilter,

    [Parameter(ParameterSetName = 'ExportUserData', Mandatory = $false, HelpMessage = "Filter für OU-Namen mit optionalen Wildcards (*), um die Suche einzuschränken.")]
    [string]$OUFilter,

    [Parameter(ParameterSetName = 'ExportUserData', Mandatory = $false, HelpMessage = "Alternativ zu -OUFilter: Der DN einer spezifischen OU als Suchbasis.")]
    [ValidateNotNullOrEmpty()]
    [string]$SearchBaseOU,

    [Parameter(ParameterSetName = 'ExportUserData', Mandatory = $true, HelpMessage = "Pfad für die CSV-Exportdatei.")]
    [ValidateNotNullOrEmpty()]
    [string]$ExportCsvPath,

    # --- Parameter für ExportLKennung ---
    [Parameter(ParameterSetName = 'ExportLKennung', Mandatory = $false, HelpMessage = "Namen der OUs für L-Kennung Export.")]
    [string[]]$LKennungOUNames = @('81', '82'), # KORREKTUR: Komma entfernt in v6.3

    [Parameter(ParameterSetName = 'ExportLKennung', Mandatory = $false, HelpMessage = "LDAP-Filter für L-Kennung Export.")]
    [string]$LKennungLDAPFilter = "(|(sAMAccountName=L110*)(sAMAccountName=L114*))", # Standard Filter

    [Parameter(ParameterSetName = 'ExportLKennung', Mandatory = $true, HelpMessage = "Pfad für die CSV-Exportdatei des L-Kennung Exports.")]
    [ValidateNotNullOrEmpty()]
    [string]$LKennungExportCsvPath,

    # --- Gemeinsame Parameter für Exporte ---
    [Parameter(ParameterSetName = 'ExportUserData')]
    [Parameter(ParameterSetName = 'ExportLKennung')]
    [string[]]$PropertiesToExport = @(), # Gilt für beide Export-Modi

    # --- Globale Parameter ---
    [Parameter(Mandatory = $false, HelpMessage = "Verzeichnis für Log-Dateien. Standard: Skriptverzeichnis oder aktuelles Verzeichnis.")]
    [string]$LogPath,

    [Parameter(Mandatory = $false, HelpMessage = "Detailtiefe der Log-Datei (Error, Warning, Info, Verbose).")]
    [ValidateSet('Error', 'Warning', 'Info', 'Verbose')]
    [string]$LogLevel = 'Info',

    [Parameter(Mandatory = $false, HelpMessage = "Pfad für den CSV-Aktionsbericht (Copy/Create/Apply). Standard: Skriptverzeichnis oder aktuelles Verzeichnis.")]
    [string]$UserReportCsvPath

)

begin {
    # --- Initialisierungen ---
    Write-Verbose "Beginne Initialisierung des Skripts."

    # Setze Kultur auf Deutsch für korrekte Formatierungen (z.B. CSV-Trennzeichen)
    try {
        [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
        [System.Threading.Thread]::CurrentThread.CurrentUICulture = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
        Write-Verbose "Kultur auf 'de-DE' gesetzt."
    }
    catch {
        Write-Warning "Konnte Kultur nicht auf 'de-DE' setzen: $_. Standardeinstellungen werden verwendet."
    }

    # Fehlerbehandlung standardmäßig auf Stop setzen
    $ErrorActionPreference = 'Stop'
    # KORREKTUR v6.4: Verbose Output standardmäßig aktivieren
    $VerbosePreference = 'Continue'


    # Bestimme Basisverzeichnis für Logs/Reports/Exports
    $basePath = $PSScriptRoot # Bevorzugt Skriptverzeichnis
    if (-not $basePath) {
        try {
            # Versuch, das Verzeichnis der aufrufenden Datei zu bekommen (funktioniert oft in ISE/VSCode)
            $basePath = Split-Path $MyInvocation.MyCommand.Path -Parent -ErrorAction Stop
            Write-Verbose "Variable `$PSScriptRoot ist leer. Verwende Verzeichnis der aufrufenden Datei als Basis: $basePath"
        } catch {
            # Fallback: Aktuelles Arbeitsverzeichnis
            $basePath = $PWD.Path
            Write-Verbose "Variable `$PSScriptRoot und Aufrufkontext nicht ermittelbar. Verwende aktuelles Arbeitsverzeichnis als Basis für Standardpfade: $basePath"
        }
    } else {
        Write-Verbose "Verwende Skriptverzeichnis als Basis für Standardpfade: $basePath"
    }

    # Bestimme Skriptnamen für Dateipräfixe
    $scriptBaseName = 'Enhanced-ADManagement' # Standard, falls nicht ermittelbar
    try {
        if ($MyInvocation.MyCommand.Name) {
            $scriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
        }
    } catch {
        Write-Warning "Konnte Skriptnamen nicht automatisch ermitteln. Verwende '$scriptBaseName' als Präfix."
    }
    Write-Verbose "Verwende '$scriptBaseName' als Präfix für Log-/Berichts-/Exportdateien."


    # Logging Setup
    $scriptStartTime = Get-Date
    # Standard-Logpfad setzen, wenn nicht angegeben
    if (-not $PSBoundParameters.ContainsKey('LogPath')) {
        $LogPath = $basePath
    }
    $logFileName = "{0}_{1}_{2}.log" -f $scriptBaseName, $PSCmdlet.ParameterSetName, $scriptStartTime.ToString('yyyyMMdd-HHmmss')
    try {
        # Sicherstellen, dass das Zielverzeichnis existiert
        if (-not (Test-Path $LogPath -PathType Container)) {
            Write-Verbose "Erstelle Log-Verzeichnis: $LogPath"
            New-Item -Path $LogPath -ItemType Directory -Force:$true -ErrorAction Stop | Out-Null
        }
        $global:fullLogPath = Join-Path -Path $LogPath -ChildPath $logFileName # Global machen für Zugriff in Funktionen/End-Block
        Write-Verbose "Log-Datei wird sein: $fullLogPath"
        # Schreibe initialen Log-Eintrag
         "[$($scriptStartTime.ToString('yyyy-MM-dd HH:mm:ss'))] [Info] Skript '$scriptBaseName.ps1' gestartet (Modus: $($PSCmdlet.ParameterSetName)). LogLevel: $LogLevel. Ausgeführt von: $($env:USERNAME)." | Out-File -FilePath $fullLogPath -Encoding UTF8 -Append
    }
    catch {
        Write-Error "Fehler beim Initialisieren des Loggings nach '$LogPath': $_. Breche Skript ab."
        # Da Logging essentiell ist, hier wirklich abbrechen
        exit 1
    }

    # User Report Setup (nur für Copy, Create, Apply Modi)
    $global:userReportData = $null
    $global:fullUserReportPath = $null
    if ($PSCmdlet.ParameterSetName -in @('CopySingleUser', 'CreateUsersFromCSV', 'ApplyPropertiesToExistingUser')) {
        $global:userReportData = [System.Collections.Generic.List[PSObject]]::new() # Global für Sammlung über Modi hinweg
        # Standard-Reportpfad setzen, wenn nicht angegeben
        if (-not $PSBoundParameters.ContainsKey('UserReportCsvPath')) {
            $UserReportCsvPath = $basePath
        }
        $reportFileName = "{0}_UserReport_{1}.csv" -f $scriptBaseName, $scriptStartTime.ToString('yyyyMMdd-HHmmss')
        try {
            # Sicherstellen, dass das Zielverzeichnis existiert
            if (-not (Test-Path $UserReportCsvPath -PathType Container)) {
                Write-Verbose "Erstelle Berichts-Verzeichnis: $UserReportCsvPath"
                New-Item -Path $UserReportCsvPath -ItemType Directory -Force:$true -ErrorAction Stop | Out-Null
            }
            $global:fullUserReportPath = Join-Path -Path $UserReportCsvPath -ChildPath $reportFileName
            Write-Verbose "Aktionsbericht wird erstellt: $fullUserReportPath"
        } catch {
            Write-Error "Fehler beim Initialisieren des Aktionsberichts-Pfades '$UserReportCsvPath': $_. Bericht kann nicht erstellt werden. Breche ab."
            exit 1 # Bericht ist obligatorisch für diese Modi
        }
    }


    # --- Hilfsfunktionen ---

    # Funktion zum Schreiben von Log-Einträgen
    function Write-Log {
        param(
            [Parameter(Mandatory = $true)]
            [ValidateSet('Error', 'Warning', 'Info', 'Verbose')]
            [string]$Level,

            [Parameter(Mandatory = $true)]
            [string]$Message
        )

        # Bestimme, ob die Nachricht basierend auf $LogLevel geloggt werden soll
        $logLevels = @{'Error' = 1; 'Warning' = 2; 'Info' = 3; 'Verbose' = 4 }
        $currentLogLevelValue = $logLevels[$LogLevel]
        $messageLogLevelValue = $logLevels[$Level]

        if ($messageLogLevelValue -le $currentLogLevelValue) {
            $logEntry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
            try {
                # Verwende den globalen Pfad
                Add-Content -Path $global:fullLogPath -Value $logEntry -Encoding UTF8 -ErrorAction Stop
            }
            catch {
                # Kritischer Fehler, wenn Logging fehlschlägt
                Write-Error "KRITISCH: Konnte Log-Eintrag nicht in '$($global:fullLogPath)' schreiben: $Message - Fehler: $_"
            }
        }

        # Zusätzliche Ausgabe auf der Konsole je nach Level
        switch ($Level) {
            'Error'   { Write-Error $Message }
            'Warning' { Write-Warning $Message }
            'Info'    { Write-Host "[INFO] $Message" -ForegroundColor Green } # Info messages green for visibility
            'Verbose' { Write-Verbose $Message } # Write-Verbose handles its own output based on $VerbosePreference/-Verbose switch
        }
    }

    # Funktion zum Hinzufügen von Daten zum Aktionsbericht (für Copy, Create, Apply)
    function Add-UserReportEntry {
        param(
            [Parameter(Mandatory = $true)]
            [string]$SamAccountName,
            [Parameter(Mandatory = $true)]
            [string]$Status, # z.B. "Erstellt", "Kopiert", "Fehler", "Modifiziert"
            [Parameter(Mandatory = $false)]
            [string]$Detail = ""
        )
        # Nur hinzufügen, wenn der Report für diesen Modus initialisiert wurde
        if ($global:userReportData -ne $null) {
            $reportObject = [PSCustomObject]@{
                Timestamp      = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                SamAccountName = $SamAccountName
                Status         = $Status
                Detail         = $Detail
            }
            $global:userReportData.Add($reportObject)
        }
    }


    # Funktion zum Kopieren eines AD-Benutzers mit Gruppen und OU
    function Copy-ADUserAdvanced {
        [CmdletBinding(SupportsShouldProcess = $true)]
        param(
            [Parameter(Mandatory = $true)]
            [Microsoft.ActiveDirectory.Management.ADUser]$SourceUser, # Ist der ReferenceUser

            [Parameter(Mandatory = $true)]
            [string]$TargetSamAccountName,

            [Parameter(Mandatory = $true)]
            [System.Security.SecureString]$Password,

            [Parameter(Mandatory = $false)]
            [string]$DestinationOU, # Optional: Ziel-OU

            [Parameter(Mandatory = $false)]
            [switch]$OverwriteTarget # Optional: Ziel überschreiben
        )

        Write-Log -Level Info -Message "Beginne Kopiervorgang von $($SourceUser.SamAccountName) nach $TargetSamAccountName."
        $targetUserExists = $false
        $existingTargetUser = $null
        try {
            $existingTargetUser = Get-ADUser -Filter "SamAccountName -eq '$TargetSamAccountName'" -ErrorAction SilentlyContinue
            if ($existingTargetUser) {
                $targetUserExists = $true
            }
        } catch {
             # Fehler beim Suchen ignorieren, weitermachen
            Write-Log -Level Warning -Message "Fehler beim Prüfen, ob Zielbenutzer '$TargetSamAccountName' existiert: $_"
        }

        if ($targetUserExists) {
            if (-not $OverwriteTarget) {
                $msg = "Zielbenutzer '$TargetSamAccountName' existiert bereits. Verwenden Sie -Force zum Überschreiben."
                Write-Log -Level Error -Message $msg
                Add-UserReportEntry -SamAccountName $TargetSamAccountName -Status "Fehler" -Detail $msg
                return $null # Fehler signalisieren
            } else {
                Write-Log -Level Warning -Message "Zielbenutzer '$TargetSamAccountName' existiert und wird überschrieben (-Force)."
                # Detailliertere ShouldProcess-Meldung für das Löschen
                $shouldProcessTarget = "den vorhandenen Benutzer '$TargetSamAccountName' (DN: $($existingTargetUser.DistinguishedName))"
                $shouldProcessAction = "Entfernen (wegen -Force)"
                if ($PSCmdlet.ShouldProcess($shouldProcessTarget, $shouldProcessAction)) {
                    try {
                        Remove-ADUser -Identity $existingTargetUser -Confirm:$false -ErrorAction Stop
                        Write-Log -Level Info -Message "Vorhandener Benutzer '$TargetSamAccountName' entfernt."
                    } catch {
                        $msg = "Fehler beim Entfernen des vorhandenen Benutzers '$TargetSamAccountName': $_"
                        Write-Log -Level Error -Message $msg
                        Add-UserReportEntry -SamAccountName $TargetSamAccountName -Status "Fehler" -Detail $msg
                        return $null
                    }
                } else {
                    $msg = "Entfernen des vorhandenen Benutzers '$TargetSamAccountName' übersprungen (ShouldProcess)."
                    Write-Log -Level Info -Message $msg
                    Add-UserReportEntry -SamAccountName $TargetSamAccountName -Status "Übersprungen" -Detail $msg
                    return $null
                }
            }
        }

        # OU bestimmen
        $finalOU = $DestinationOU # Aus Parameter verwenden
        if (-not $finalOU) {
            # Wenn nicht im Parameter, nimm die OU des Quellbenutzers
            $finalOU = ($SourceUser.DistinguishedName -split ',', 2)[1]
            Write-Verbose "Keine Ziel-OU (-TargetOU) angegeben, verwende Quell-OU: $finalOU"
        }

        # Prüfen, ob die Ziel-OU existiert
        try {
            if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$finalOU'" -ErrorAction Stop)) {
                 # Dies sollte eigentlich nicht passieren, da Get-ADOrganizationalUnit einen Fehler werfen sollte, wenn nicht gefunden. Doppelte Sicherheit.
                $msg = "Die Ziel-OU '$finalOU' existiert nicht."
                Write-Log -Level Error -Message $msg
                Add-UserReportEntry -SamAccountName $TargetSamAccountName -Status "Fehler" -Detail $msg
                return $null
            }
             Write-Verbose "Ziel-OU '$finalOU' ist gültig."
        } catch {
            $msg = "Fehler beim Überprüfen der Ziel-OU '$finalOU': $_"
            Write-Log -Level Error -Message $msg
            Add-UserReportEntry -SamAccountName $TargetSamAccountName -Status "Fehler" -Detail $msg
            return $null
        }


        # Parameter für New-ADUser vorbereiten
        $newUserParams = @{
            SamAccountName        = $TargetSamAccountName
            Name                  = $TargetSamAccountName # Standard Name = SamAccountName, kann später angepasst werden
            GivenName             = $SourceUser.GivenName
            Surname               = $SourceUser.Surname
            DisplayName           = "$($SourceUser.GivenName) $($SourceUser.Surname)" # Oder eine andere Logik
            UserPrincipalName     = "$TargetSamAccountName@$($env:USERDNSDOMAIN)" # Domain anpassen falls nötig!
            Path                  = $finalOU
            AccountPassword       = $Password
            ChangePasswordAtLogon = $true
            Enabled               = $true # Kopierte Benutzer standardmäßig aktivieren
            Description           = $SourceUser.Description # Beispiel für weitere Attribute
            Office                = $SourceUser.Office
            Department            = $SourceUser.Department
            Company               = $SourceUser.Company
            Title                 = $SourceUser.Title
            # Fügen Sie hier weitere Attribute hinzu, die kopiert werden sollen
            StreetAddress         = $SourceUser.StreetAddress
            City                  = $SourceUser.City
            State                 = $SourceUser.State
            PostalCode            = $SourceUser.PostalCode
            Country               = $SourceUser.Country
            OfficePhone           = $SourceUser.OfficePhone
            EmailAddress          = $SourceUser.EmailAddress
        }

        # Benutzer erstellen
        $newUser = $null
        # Detailliertere ShouldProcess-Meldung
        $shouldProcessTarget = "Benutzer '$TargetSamAccountName' (Kopie von '$($SourceUser.SamAccountName)')"
        $shouldProcessAction = "Erstellen in OU '$finalOU'"
        if ($PSCmdlet.ShouldProcess($shouldProcessTarget, $shouldProcessAction)) {
            try {
                Write-Log -Level Info -Message "Erstelle Benutzer '$TargetSamAccountName' in OU '$finalOU'."
                # Verwende NICHT -Instance hier, um mehr Kontrolle zu haben
                $newUser = New-ADUser @newUserParams -PassThru -ErrorAction Stop
                Write-Log -Level Info -Message "Benutzer '$($newUser.SamAccountName)' erfolgreich erstellt (SID: $($newUser.SID.Value))."
                Add-UserReportEntry -SamAccountName $newUser.SamAccountName -Status "Kopiert" -Detail "Von $($SourceUser.SamAccountName) nach OU '$finalOU'"
            }
            catch {
                $msg = "Fehler beim Erstellen des Benutzers '$TargetSamAccountName': $_"
                Write-Log -Level Error -Message $msg
                Add-UserReportEntry -SamAccountName $TargetSamAccountName -Status "Fehler" -Detail $msg
                return $null
            }
        } else {
            $msg = "Erstellung von '$TargetSamAccountName' übersprungen (ShouldProcess)."
            Write-Log -Level Info -Message $msg
            Add-UserReportEntry -SamAccountName $TargetSamAccountName -Status "Übersprungen" -Detail $msg
            return $null
        }

        # Gruppenmitgliedschaften kopieren
        try {
            $sourceGroups = Get-ADPrincipalGroupMembership -Identity $SourceUser -ErrorAction Stop
             # Filter optionale problematische Gruppen (z.B. 'Domain Users' wird oft automatisch hinzugefügt)
             $groupsToCopy = $sourceGroups | Where-Object {$_.Name -ne "Domain Users"} # Beispiel Filter

            if ($groupsToCopy) {
                $groupNames = $groupsToCopy.Name -join ', '
                Write-Log -Level Info -Message "Kopiere $($groupsToCopy.Count) Gruppenmitgliedschaften von $($SourceUser.SamAccountName) zu $($newUser.SamAccountName)."
                # Detailliertere ShouldProcess-Meldung
                $shouldProcessTarget = "Benutzer '$($newUser.SamAccountName)'"
                $shouldProcessAction = "Hinzufügen zu Gruppen ($($groupsToCopy.Count)): $groupNames"
                if ($PSCmdlet.ShouldProcess($shouldProcessTarget, $shouldProcessAction)) {
                    Add-ADPrincipalGroupMembership -Identity $newUser -MemberOf $groupsToCopy -ErrorAction Stop
                    Write-Log -Level Info -Message "Gruppenmitgliedschaften erfolgreich kopiert."
                    # Optional: Update Report
                    # Add-UserReportEntry -SamAccountName $newUser.SamAccountName -Status "Gruppen kopiert" -Detail "$($groupsToCopy.Count) Gruppen"
                } else {
                     Write-Log -Level Info -Message "Kopieren der Gruppenmitgliedschaften übersprungen (ShouldProcess)."
                }
            } else {
                Write-Log -Level Info -Message "Quellbenutzer $($SourceUser.SamAccountName) hat keine (relevanten) Gruppenmitgliedschaften zum Kopieren."
            }
        }
        catch {
            $msg = "Fehler beim Kopieren der Gruppenmitgliedschaften für '$($newUser.SamAccountName)': $_. Der Benutzer wurde erstellt, aber Gruppen fehlen möglicherweise."
            Write-Log -Level Warning -Message $msg
            Add-UserReportEntry -SamAccountName $newUser.SamAccountName -Status "Warnung" -Detail "Fehler beim Kopieren der Gruppen: $_"
            # Nicht abbrechen, Benutzer existiert ja schon
        }

        Write-Log -Level Info -Message "Kopiervorgang für $TargetSamAccountName abgeschlossen."
        return $newUser
    }

     # Funktion zum Erstellen eines Benutzers aus Daten (CSV-Zeile/Hashtable)
     function New-ADUserFromData {
         [CmdletBinding(SupportsShouldProcess = $true)]
         param(
             [Parameter(Mandatory = $true)]
             [hashtable]$UserData, # Enthält alle Infos aus CSV-Zeile

             [Parameter(Mandatory = $false)]
             [Microsoft.ActiveDirectory.Management.ADUser]$TemplateUser, # Optional: Template User Objekt (ReferenceUser)

             [Parameter(Mandatory = $false)]
             [System.Security.SecureString]$GlobalDefaultPassword, # Optional: Fallback Passwort

             [Parameter(Mandatory = $false)]
             [string]$GlobalTargetOU # Optional: Fallback OU
         )

         # Versuche SamAccountName zu bekommen, bevor geloggt wird
         $sam = $UserData.SamAccountName
         if (-not $sam) {
             # Wenn SamAccountName fehlt, können wir nicht viel tun
             $msg = "Fehlender Wert für 'SamAccountName' in den Daten. Überspringe Eintrag."
             Write-Log -Level Error -Message $msg
             Add-UserReportEntry -SamAccountName "(Unbekannt)" -Status "Fehler" -Detail $msg
             return $null
         }

         Write-Log -Level Info -Message "Beginne Verarbeitung zur Erstellung von Benutzer '$sam'."

         # --- Validierung der Pflichtfelder ---
         if (-not $UserData.GivenName) {
             $msg = "Fehlender Wert für 'GivenName' (Vorname) für '$sam'. Überspringe Eintrag."
             Write-Log -Level Error -Message $msg
             Add-UserReportEntry -SamAccountName $sam -Status "Fehler" -Detail $msg
             return $null
         }
         if (-not $UserData.Surname) {
             $msg = "Fehlender Wert für 'Surname' (Nachname) für '$sam'. Überspringe Eintrag."
             Write-Log -Level Error -Message $msg
             Add-UserReportEntry -SamAccountName $sam -Status "Fehler" -Detail $msg
             return $null
         }

         # --- Ziel-OU bestimmen (Priorität: Global Parameter > CSV > Template > Fehler) ---
         $finalOU = $null
         if ($GlobalTargetOU) {
             $finalOU = $GlobalTargetOU
             Write-Verbose "Verwende globale Ziel-OU '$finalOU' für '$sam'."
         } elseif ($UserData.ContainsKey('TargetOU') -and $UserData.TargetOU) {
             $finalOU = $UserData.TargetOU
             Write-Verbose "Verwende Ziel-OU aus Datenquelle '$finalOU' für '$sam'."
         } elseif ($TemplateUser) {
             $finalOU = ($TemplateUser.DistinguishedName -split ',', 2)[1]
             Write-Verbose "Verwende Ziel-OU vom Template-Benutzer '$finalOU' für '$sam'."
         } else {
             $msg = "Keine Ziel-OU für Benutzer '$sam' gefunden (weder in CSV, noch als Parameter, noch durch Template). Überspringe Eintrag."
             Write-Log -Level Error -Message $msg
             Add-UserReportEntry -SamAccountName $sam -Status "Fehler" -Detail $msg
             return $null
         }

         # --- Passwort bestimmen (Priorität: CSV > Global Parameter > Generieren) ---
         $finalPassword = $null
         $changePwdAtLogon = $true
         if ($UserData.ContainsKey('Password') -and $UserData.Password) {
             Write-Log -Level Warning -Message "Verwende Passwort aus Datenquelle für '$sam'. ACHTUNG: Klartextpasswörter sind ein Sicherheitsrisiko!"
             try {
                 $finalPassword = ConvertTo-SecureString $UserData.Password -AsPlainText -Force -ErrorAction Stop
             } catch {
                 $msg = "Fehler beim Konvertieren des Passworts aus der Datenquelle für '$sam': $_. Überspringe Eintrag."
                 Write-Log -Level Error -Message $msg
                 Add-UserReportEntry -SamAccountName $sam -Status "Fehler" -Detail $msg
                 return $null
             }
         } elseif ($GlobalDefaultPassword) {
             Write-Verbose "Verwende globales Standardpasswort für '$sam'."
             $finalPassword = $GlobalDefaultPassword
         } else {
             Write-Log -Level Info -Message "Generiere zufälliges Passwort für '$sam', da keines angegeben wurde."
             # Generiere sicheres, zufälliges Passwort
             try {
                 # Komplexeres Beispiel: Mind. 1 Groß, 1 Klein, 1 Zahl, 1 Sonderzeichen, Länge 14
                 $pwdChars = @()
                 $pwdChars += 65..90 | Get-Random # Großbuchstabe
                 $pwdChars += 97..122 | Get-Random # Kleinbuchstabe
                 $pwdChars += 48..57 | Get-Random # Zahl
                 $pwdChars += 33, 35, 36, 37, 38, 42, 64, 95 | Get-Random # Sonderzeichen !#$%&*@_
                 # Restliche Zeichen auffüllen (insgesamt 14)
                 $allChars = (48..57) + (65..90) + (97..122) + 33, 35, 36, 37, 38, 42, 64, 95
                 $pwdChars += $allChars | Get-Random -Count (14 - $pwdChars.Count)
                 # Mischen
                 $randomPassword = -join ($pwdChars | Get-Random -Count $pwdChars.Count | % {[char]$_})

                 $finalPassword = ConvertTo-SecureString $randomPassword -AsPlainText -Force -ErrorAction Stop
                 Write-Log -Level Info -Message "Zufälliges Passwort für '$sam' generiert. Benutzer MUSS es bei der ersten Anmeldung ändern."
             } catch {
                 $msg = "Fehler beim Generieren/Konvertieren des zufälligen Passworts für '$sam': $_. Überspringe Eintrag."
                 Write-Log -Level Error -Message $msg
                 Add-UserReportEntry -SamAccountName $sam -Status "Fehler" -Detail $msg
                 return $null
             }
         }

         # --- Prüfen ob Zielbenutzer existiert ---
         try {
             if (Get-ADUser -Filter "SamAccountName -eq '$sam'" -ErrorAction SilentlyContinue) {
                 $msg = "Benutzer '$sam' existiert bereits im AD. Überspringe Eintrag."
                 Write-Log -Level Error -Message $msg
                 Add-UserReportEntry -SamAccountName $sam -Status "Fehler" -Detail $msg
                 # KORREKTUR v6.7: Sofort abbrechen, wenn Benutzer existiert
                 return $null
             }
         } catch {
             # Fehler beim Suchen ist unwahrscheinlich, aber sicherheitshalber loggen und abbrechen
             $msg = "Fehler beim Prüfen, ob Benutzer '$sam' existiert: $_."
             Write-Log -Level Warning -Message $msg
             Add-UserReportEntry -SamAccountName $sam -Status "Fehler" -Detail $msg
             return $null
         }


        # --- Prüfen ob Ziel-OU existiert ---
        try {
            if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$finalOU'" -ErrorAction Stop)) {
                 $msg = "Die Ziel-OU '$finalOU' für '$sam' existiert nicht. Überspringe Eintrag."
                 Write-Log -Level Error -Message $msg
                 Add-UserReportEntry -SamAccountName $sam -Status "Fehler" -Detail $msg
                 return $null
            }
            Write-Verbose "Ziel-OU '$finalOU' für '$sam' ist gültig."
        } catch {
            $msg = "Fehler beim Überprüfen der Ziel-OU '$finalOU' für '$sam': $_. Überspringe Eintrag."
            Write-Log -Level Error -Message $msg
            Add-UserReportEntry -SamAccountName $sam -Status "Fehler" -Detail $msg
            return $null
        }

         # --- Benutzerparameter zusammenstellen ---
         $newUserParams = @{
             SamAccountName        = $sam
             Name                  = "$($UserData.GivenName) $($UserData.Surname)" # Standard Name
             GivenName             = $UserData.GivenName
             Surname               = $UserData.Surname
             DisplayName           = "$($UserData.GivenName) $($UserData.Surname)" # Standard DisplayName
             UserPrincipalName     = "$sam@$($env:USERDNSDOMAIN)" # Anpassen falls nötig
             Path                  = $finalOU
             AccountPassword       = $finalPassword
             ChangePasswordAtLogon = $changePwdAtLogon
             Enabled               = $true # Standardmäßig aktivieren
         }

         # Enabled-Status aus Datenquelle übernehmen, falls vorhanden
         if ($UserData.ContainsKey('Enabled')) {
             try {
                 $newUserParams.Enabled = [bool]::Parse($UserData.Enabled) # Sicherere Konvertierung
                 Write-Verbose "Setze 'Enabled' für '$sam' auf '$($newUserParams.Enabled)' basierend auf Datenquelle."
             } catch {
                 Write-Log -Level Warning -Message "Konnte 'Enabled'-Wert '$($UserData.Enabled)' für '$sam' nicht in Boolean konvertieren. Verwende Standard ($true)."
             }
         }

         # Weitere Attribute aus Datenquelle oder Template übernehmen
         $attributesToCheck = @(
             'Description', 'Office', 'Department', 'Title', 'Company',
             'EmailAddress', 'StreetAddress', 'City', 'State', 'PostalCode', 'Country', 'OfficePhone'
             # Fügen Sie hier weitere AD-Attribute hinzu, die unterstützt werden sollen
         )

         foreach ($attr in $attributesToCheck) {
             if ($UserData.ContainsKey($attr) -and $UserData.$attr) {
                 $newUserParams[$attr] = $UserData.$attr
                 Write-Verbose "Setze Attribut '$attr' für '$sam' aus Datenquelle."
             } elseif ($TemplateUser) {
                 # Nur übernehmen, wenn Attribut im Template existiert und nicht leer ist
                 if ($TemplateUser.PSObject.Properties.Match($attr).Count -gt 0 -and $TemplateUser.$attr -ne $null -and $TemplateUser.$attr -ne '') {
                    $newUserParams[$attr] = $TemplateUser.$attr
                    Write-Verbose "Setze Attribut '$attr' für '$sam' vom Template-Benutzer."
                 }
             }
         }

         # --- Benutzer erstellen (Jetzt NACH Prüfung auf Existenz) ---
         $newUser = $null
         # Detailliertere ShouldProcess-Meldung
         $shouldProcessTarget = "Benutzer '$sam' ($($newUserParams.Name))"
         $shouldProcessAction = "Erstellen in OU '$finalOU'"
         if ($PSCmdlet.ShouldProcess($shouldProcessTarget, $shouldProcessAction)) {
             try {
                 Write-Log -Level Info -Message "Erstelle Benutzer '$sam' in OU '$finalOU'."
                 $newUser = New-ADUser @newUserParams -PassThru -ErrorAction Stop
                 Write-Log -Level Info -Message "Benutzer '$($newUser.SamAccountName)' erfolgreich erstellt (SID: $($newUser.SID.Value))."
                 Add-UserReportEntry -SamAccountName $newUser.SamAccountName -Status "Erstellt" -Detail "In OU '$finalOU'"
             } catch {
                 $msg = "Fehler beim Erstellen des Benutzers '$sam': $_"
                 Write-Log -Level Error -Message $msg
                 Add-UserReportEntry -SamAccountName $sam -Status "Fehler" -Detail $msg
                 return $null # Abbruch für diesen Benutzer
             }
         } else {
             $msg = "Erstellung von '$sam' übersprungen (ShouldProcess)."
             Write-Log -Level Info -Message $msg
             Add-UserReportEntry -SamAccountName $sam -Status "Übersprungen" -Detail $msg
             return $null # Abbruch für diesen Benutzer
         }

         # --- Gruppen vom Template übernehmen ---
         if ($TemplateUser -and $newUser) { # Nur wenn Template vorhanden UND Benutzer erfolgreich erstellt wurde
             try {
                 $templateGroups = Get-ADPrincipalGroupMembership -Identity $TemplateUser -ErrorAction Stop
                 $groupsToCopy = $templateGroups | Where-Object {$_.Name -ne "Domain Users"} # Filter

                 if ($groupsToCopy) {
                     $groupNames = $groupsToCopy.Name -join ', '
                     Write-Log -Level Info -Message "Füge Benutzer '$($newUser.SamAccountName)' zu $($groupsToCopy.Count) Gruppen hinzu (basierend auf Template '$($TemplateUser.SamAccountName)')."
                     # Detailliertere ShouldProcess-Meldung
                     $shouldProcessTarget = "Benutzer '$($newUser.SamAccountName)'"
                     $shouldProcessAction = "Hinzufügen zu Gruppen ($($groupsToCopy.Count)): $groupNames"
                     if ($PSCmdlet.ShouldProcess($shouldProcessTarget, $shouldProcessAction)) {
                         Add-ADPrincipalGroupMembership -Identity $newUser -MemberOf $groupsToCopy -ErrorAction Stop
                         Write-Log -Level Info -Message "Gruppenmitgliedschaften für '$($newUser.SamAccountName)' erfolgreich hinzugefügt."
                         # Optional: Update Report
                         # Add-UserReportEntry -SamAccountName $newUser.SamAccountName -Status "Gruppen hinzugefügt" -Detail "$($groupsToCopy.Count) Gruppen von $($TemplateUser.SamAccountName)"
                     } else {
                         Write-Log -Level Info -Message "Hinzufügen der Gruppenmitgliedschaften übersprungen (ShouldProcess)."
                     }
                 } else {
                     Write-Log -Level Info -Message "Template-Benutzer $($TemplateUser.SamAccountName) hat keine (relevanten) Gruppenmitgliedschaften zum Hinzufügen."
                 }
             } catch {
                 $msg = "Fehler beim Hinzufügen der Gruppenmitgliedschaften vom Template für '$($newUser.SamAccountName)': $_. Der Benutzer wurde erstellt, aber Gruppen fehlen möglicherweise."
                 Write-Log -Level Warning -Message $msg
                 Add-UserReportEntry -SamAccountName $newUser.SamAccountName -Status "Warnung" -Detail "Fehler beim Hinzufügen der Gruppen: $_"
             }
         }

         Write-Log -Level Info -Message "Verarbeitung für Benutzer '$sam' abgeschlossen."
         return $newUser
     }

    # Hilfsfunktion zum Synchronisieren der Attribute
    function Sync-ADUserAttributes {
        [CmdletBinding(SupportsShouldProcess = $true)]
        param(
            [Parameter(Mandatory = $true)]
            [Microsoft.ActiveDirectory.Management.ADUser]$ReferenceUser,

            [Parameter(Mandatory = $true)]
            [Microsoft.ActiveDirectory.Management.ADUser]$TargetUser
        )

        $propertiesToApply = @(
            'Description', 'Office', 'StreetAddress', 'City', 'State',
            'PostalCode', 'Country', 'Department', 'Company', 'Title',
            'OfficePhone', 'EmailAddress'
        )

        $setParams = @{ Identity = $TargetUser }
        $changesDetected = $false
        $changedPropsList = @()

        foreach ($prop in $propertiesToApply) {
            if ($ReferenceUser.$prop -ne $TargetUser.$prop) {
                 if ($ReferenceUser.PSObject.Properties.Match($prop).Count -gt 0 -and $ReferenceUser.$prop -ne $null -and $ReferenceUser.$prop -ne '') {
                    Write-Verbose "Änderung für '$($TargetUser.SamAccountName)' bei Eigenschaft '$prop': '$($TargetUser.$prop)' -> '$($ReferenceUser.$prop)'"
                    $setParams[$prop] = $ReferenceUser.$prop
                    $changesDetected = $true
                    $changedPropsList += $prop
                } else {
                    Write-Verbose "Eigenschaft '$prop' ist im Referenzbenutzer '$($ReferenceUser.SamAccountName)' leer, null oder nicht vorhanden, wird für '$($TargetUser.SamAccountName)' nicht überschrieben."
                }
            }
        }

        if ($changesDetected) {
            $shouldProcessTarget = "Benutzer '$($TargetUser.SamAccountName)'"
            $shouldProcessAction = "Eigenschaften anwenden (Quelle: $($ReferenceUser.SamAccountName)): $($changedPropsList -join ', ')"
            if ($PSCmdlet.ShouldProcess($shouldProcessTarget, $shouldProcessAction)) {
                try {
                    Set-ADUser @setParams -ErrorAction Stop
                    Write-Log -Level Info -Message "Eigenschaften erfolgreich auf '$($TargetUser.SamAccountName)' angewendet."
                    Add-UserReportEntry -SamAccountName $TargetUser.SamAccountName -Status "Modifiziert" -Detail "Eigenschaften angewendet von $($ReferenceUser.SamAccountName)"
                } catch {
                    $msg = "Fehler beim Anwenden der Eigenschaften auf '$($TargetUser.SamAccountName)': $_"
                    Write-Log -Level Error -Message $msg
                    Add-UserReportEntry -SamAccountName $TargetUser.SamAccountName -Status "Fehler" -Detail "Fehler beim Anwenden der Eigenschaften: $_"
                }
            } else {
                Write-Log -Level Info -Message "Anwenden der Eigenschaften auf '$($TargetUser.SamAccountName)' übersprungen (ShouldProcess)."
                Add-UserReportEntry -SamAccountName $TargetUser.SamAccountName -Status "Übersprungen" -Detail "Anwenden der Eigenschaften übersprungen (ShouldProcess)"
            }
        } else {
            Write-Log -Level Info -Message "Keine unterschiedlichen Eigenschaften zum Anwenden auf '$($TargetUser.SamAccountName)' gefunden."
        }

        return $changesDetected
    }

    # Hilfsfunktion zum Synchronisieren der Gruppenmitgliedschaften
    function Sync-ADUserGroups {
        [CmdletBinding(SupportsShouldProcess = $true)]
        param(
            [Parameter(Mandatory = $true)]
            [Microsoft.ActiveDirectory.Management.ADUser]$ReferenceUser,

            [Parameter(Mandatory = $true)]
            [Microsoft.ActiveDirectory.Management.ADUser]$TargetUser,

            [Parameter(Mandatory = $true)]
            [bool]$AttributesChanged
        )

        try {
            Write-Verbose "Prüfe Gruppenmitgliedschaften für '$($TargetUser.SamAccountName)' (Quelle: $($ReferenceUser.SamAccountName))."
            $referenceGroups = Get-ADPrincipalGroupMembership -Identity $ReferenceUser -ErrorAction Stop
            $targetGroups = Get-ADPrincipalGroupMembership -Identity $TargetUser -ErrorAction Stop

            # Gruppen filtern (z.B. Domain Users ausschließen)
            $referenceGroupsFiltered = $referenceGroups | Where-Object {$_.Name -ne "Domain Users"}

            # Gruppen finden, die der Referenzbenutzer hat, der Zielbenutzer aber nicht
            $groupsToAdd = Compare-Object -ReferenceObject $referenceGroupsFiltered -DifferenceObject $targetGroups -Property DistinguishedName -PassThru | Where-Object {$_.SideIndicator -eq '<='}

            if ($groupsToAdd) {
                $groupNames = $groupsToAdd.Name -join ', '
                Write-Log -Level Info -Message "Füge $($groupsToAdd.Count) fehlende Gruppenmitgliedschaften zu '$($TargetUser.SamAccountName)' hinzu (Quelle: $($ReferenceUser.SamAccountName))."

                $shouldProcessTarget = "Benutzer '$($TargetUser.SamAccountName)'"
                $shouldProcessAction = "Hinzufügen zu Gruppen ($($groupsToAdd.Count)): $groupNames"
                if ($PSCmdlet.ShouldProcess($shouldProcessTarget, $shouldProcessAction)) {
                    Add-ADPrincipalGroupMembership -Identity $TargetUser -MemberOf $groupsToAdd -ErrorAction Stop
                    Write-Log -Level Info -Message "Fehlende Gruppenmitgliedschaften erfolgreich zu '$($TargetUser.SamAccountName)' hinzugefügt."
                    Add-UserReportEntry -SamAccountName $TargetUser.SamAccountName -Status "Modifiziert" -Detail "$($groupsToAdd.Count) Gruppen hinzugefügt von $($ReferenceUser.SamAccountName)"
                } else {
                    Write-Log -Level Info -Message "Hinzufügen fehlender Gruppenmitgliedschaften zu '$($TargetUser.SamAccountName)' übersprungen (ShouldProcess)."
                    Add-UserReportEntry -SamAccountName $TargetUser.SamAccountName -Status "Übersprungen" -Detail "Hinzufügen fehlender Gruppen übersprungen (ShouldProcess)"
                }
            } else {
                Write-Log -Level Info -Message "Keine fehlenden Gruppenmitgliedschaften bei '$($TargetUser.SamAccountName)' gefunden (basierend auf $($ReferenceUser.SamAccountName))."
                 # Report-Eintrag nur, wenn auch bei Eigenschaften keine Änderung erfolgte
                 if(-not $AttributesChanged) {
                    Add-UserReportEntry -SamAccountName $TargetUser.SamAccountName -Status "Keine Änderung" -Detail "Keine Eigenschafts- oder Gruppenänderungen von $($ReferenceUser.SamAccountName) nötig"
                 }
            }
        } catch {
             $msg = "Fehler beim Verarbeiten der Gruppenmitgliedschaften für '$($TargetUser.SamAccountName)': $_"
             Write-Log -Level Warning -Message $msg
             Add-UserReportEntry -SamAccountName $TargetUser.SamAccountName -Status "Warnung" -Detail "Fehler beim Verarbeiten der Gruppen: $_"
        }
    }

    # Funktion zum Anwenden von Eigenschaften auf einen existierenden Benutzer
    function Apply-ADUserProperties {
        [CmdletBinding(SupportsShouldProcess = $true)]
        param(
            [Parameter(Mandatory = $true)]
            [Microsoft.ActiveDirectory.Management.ADUser]$ReferenceUser, # Quelle der Eigenschaften

            [Parameter(Mandatory = $true)]
            [Microsoft.ActiveDirectory.Management.ADUser]$TargetUser # Ziel der Modifikation
        )

        Write-Log -Level Info -Message "Beginne Anwenden von Eigenschaften von '$($ReferenceUser.SamAccountName)' auf '$($TargetUser.SamAccountName)'."

        # --- Eigenschaften anwenden ---
        $attributesChanged = Sync-ADUserAttributes -ReferenceUser $ReferenceUser -TargetUser $TargetUser

        # --- Gruppenmitgliedschaften hinzufügen ---
        Sync-ADUserGroups -ReferenceUser $ReferenceUser -TargetUser $TargetUser -AttributesChanged $attributesChanged

        Write-Log -Level Info -Message "Anwenden von Eigenschaften/Gruppen auf '$($TargetUser.SamAccountName)' abgeschlossen."

    }


    Write-Verbose "Initialisierung abgeschlossen. Wechsle zur Prozess-Phase."
} # End Begin Block

process {
    Write-Verbose "Beginne Prozess-Phase. Ausgewählter Modus: $($PSCmdlet.ParameterSetName)"

    switch ($PSCmdlet.ParameterSetName) {
        # --- Modus: CopySingleUser ---
        'CopySingleUser' {
            Write-Log -Level Info -Message "Starte Modus: CopySingleUser"

            # 1. Referenzbenutzer (Quelle) validieren
            $referenceUserObject = $null
            try {
                # Verwende den konsolidierten Parameter $ReferenceUserSamAccountName
                $referenceUserObject = Get-ADUser -Identity $ReferenceUserSamAccountName -Properties * -ErrorAction Stop # Lade alle Properties für die Kopie
                Write-Log -Level Info -Message "Referenzbenutzer (Quelle) '$($referenceUserObject.SamAccountName)' gefunden."
            } catch {
                $msg = "Referenzbenutzer (Quelle) '$ReferenceUserSamAccountName' konnte nicht gefunden werden: $_"
                Write-Log -Level Error -Message $msg
                Add-UserReportEntry -SamAccountName $ReferenceUserSamAccountName -Status "Fehler" -Detail $msg
                return # Abbruch des Modus
            }

            # 2. Zielbenutzernamen abfragen, wenn nicht angegeben
            if (-not $TargetUserSamAccountName) {
                try {
                    $TargetUserSamAccountName = Read-Host -Prompt "Bitte geben Sie den gewünschten SAMAccountName für den NEUEN Benutzer ein"
                    if (-not $TargetUserSamAccountName) { throw "Eingabe darf nicht leer sein."}
                } catch {
                     $msg = "Ungültige Eingabe für Zielbenutzernamen: $_"
                     Write-Log -Level Error -Message $msg
                     Add-UserReportEntry -SamAccountName "(Unbekannt)" -Status "Fehler" -Detail $msg
                     return
                }
            }

            # 3. Passwort abfragen, wenn nicht angegeben
            if (-not $TargetUserPassword) {
                 try {
                     $TargetUserPassword = Read-Host -Prompt "Bitte geben Sie das initiale Passwort für '$TargetUserSamAccountName' ein" -AsSecureString
                     if ($TargetUserPassword.Length -eq 0) { throw "Passwort darf nicht leer sein."} # Einfache Prüfung
                 } catch {
                     $msg = "Ungültige Eingabe für Passwort: $_"
                     Write-Log -Level Error -Message $msg
                     Add-UserReportEntry -SamAccountName $TargetUserSamAccountName -Status "Fehler" -Detail $msg
                     return
                 }
            }

            # 4. Kopierfunktion aufrufen
            $newUser = Copy-ADUserAdvanced -SourceUser $referenceUserObject `
                                          -TargetSamAccountName $TargetUserSamAccountName `
                                          -Password $TargetUserPassword `
                                          -DestinationOU $TargetOU `
                                          -OverwriteTarget:$Force
                                          # KORREKTUR v6.6: Explizite Parameterübergabe entfernt

            if ($newUser) {
                Write-Log -Level Info -Message "Benutzer '$($newUser.SamAccountName)' erfolgreich kopiert."
                # Report-Eintrag wird bereits in Copy-ADUserAdvanced hinzugefügt
            } else {
                Write-Log -Level Error -Message "Fehler beim Kopieren des Benutzers '$ReferenceUserSamAccountName'."
                # Fehler und Report-Eintrag wurden bereits in Copy-ADUserAdvanced geloggt/hinzugefügt
            }
        } # End CopySingleUser

        # --- Modus: CreateUsersFromCSV ---
        'CreateUsersFromCSV' {
            Write-Log -Level Info -Message "Starte Modus: CreateUsersFromCSV"
            $SuccessCount = 0
            $FailCount = 0

            # 1. CSV-Datei einlesen
            $usersData = @()
            try {
                Write-Log -Level Info -Message "Lese CSV-Datei: $CsvPath"
                $usersData = Import-Csv -Path $CsvPath -Delimiter ';' -Encoding UTF8 -ErrorAction Stop
                Write-Log -Level Info -Message "$($usersData.Count) Einträge in CSV-Datei gefunden."
                # Prüfung auf essentielle Header
                if($usersData.Count -gt 0){
                    $headers = $usersData[0].PSObject.Properties.Name
                    if(-not ($headers -contains 'SamAccountName' -and $headers -contains 'GivenName' -and $headers -contains 'Surname')){
                       throw "CSV-Datei '$CsvPath' fehlen essentielle Header: SamAccountName, GivenName, Surname."
                    }
                }

            } catch {
                $msg = "Fehler beim Lesen oder Validieren der CSV-Datei '$CsvPath': $_"
                Write-Log -Level Error -Message $msg
                Add-UserReportEntry -SamAccountName "(CSV)" -Status "Fehler" -Detail $msg
                return # Abbruch des Modus
            }

            if ($usersData.Count -eq 0) {
                 $msg = "CSV-Datei '$CsvPath' ist leer oder enthält keine Daten."
                 Write-Log -Level Warning -Message $msg
                 Add-UserReportEntry -SamAccountName "(CSV)" -Status "Warnung" -Detail $msg
                 return
            }

            # 2. Referenzbenutzer (Template) laden (wenn angegeben)
            $referenceUserObject = $null
            # Verwende den konsolidierten Parameter $ReferenceUserSamAccountName
            if ($ReferenceUserSamAccountName) {
                try {
                    $referenceUserObject = Get-ADUser -Identity $ReferenceUserSamAccountName -Properties * -ErrorAction Stop # Alle Properties laden
                    Write-Log -Level Info -Message "Referenzbenutzer (Template) '$($referenceUserObject.SamAccountName)' gefunden und geladen."
                } catch {
                    Write-Log -Level Warning -Message "Referenzbenutzer (Template) '$ReferenceUserSamAccountName' konnte nicht gefunden werden: $_. Gruppen und Attribute werden nicht vom Template übernommen."
                    # Nicht abbrechen, aber weitermachen ohne Template
                }
            } else {
                 Write-Log -Level Info -Message "Kein Referenzbenutzer (Template) angegeben (-ReferenceUserSamAccountName). Attribute/Gruppen werden nur aus CSV übernommen."
            }

            # 3. Jeden Eintrag in der CSV verarbeiten
            Write-Log -Level Info -Message "Beginne Verarbeitung der CSV-Einträge..."
            foreach ($userRow in $usersData) {
                $userDataHash = @{}
                # Konvertiere PSObject aus Import-Csv in Hashtable für einfachere Handhabung
                $userRow.PSObject.Properties | ForEach-Object { $userDataHash[$_.Name] = $_.Value }

                $newUser = New-ADUserFromData -UserData $userDataHash `
                                              -TemplateUser $referenceUserObject `
                                              -GlobalDefaultPassword $DefaultPassword `
                                              -GlobalTargetOU $TargetOU
                                              # KORREKTUR v6.6: Explizite Parameterübergabe entfernt

                if ($newUser) {
                    $SuccessCount++
                } else {
                    $FailCount++
                    # Fehler und Report-Eintrag wurden bereits in New-ADUserFromData geloggt/hinzugefügt
                }
            } # End foreach userRow

            Write-Log -Level Info -Message "CSV-Verarbeitung abgeschlossen. Erfolgreich erstellt: $SuccessCount, Fehlgeschlagen/Übersprungen: $FailCount."

        } # End CreateUsersFromCSV

        # --- Modus: ApplyPropertiesToExistingUser ---
        'ApplyPropertiesToExistingUser' {
            Write-Log -Level Info -Message "Starte Modus: ApplyPropertiesToExistingUser"

            # 1. Referenzbenutzer (Quelle) validieren
            $referenceUserObject = $null
            try {
                $referenceUserObject = Get-ADUser -Identity $ReferenceUserSamAccountName -Properties * -ErrorAction Stop # Alle Properties laden
                Write-Log -Level Info -Message "Referenzbenutzer (Quelle) '$($referenceUserObject.SamAccountName)' gefunden."
            } catch {
                $msg = "Referenzbenutzer (Quelle) '$ReferenceUserSamAccountName' konnte nicht gefunden werden: $_"
                Write-Log -Level Error -Message $msg
                Add-UserReportEntry -SamAccountName $ReferenceUserSamAccountName -Status "Fehler" -Detail $msg
                return # Abbruch des Modus
            }

            # 2. Zielbenutzer validieren (muss existieren)
            $targetUserObject = $null
            try {
                # Verwende den konsolidierten Parameter $TargetUserSamAccountName
                $targetUserObject = Get-ADUser -Identity $TargetUserSamAccountName -Properties * -ErrorAction Stop # Alle Properties laden
                Write-Log -Level Info -Message "Zielbenutzer '$($targetUserObject.SamAccountName)' gefunden."
            } catch {
                $msg = "Zielbenutzer '$TargetUserSamAccountName' konnte nicht gefunden werden oder existiert nicht: $_"
                Write-Log -Level Error -Message $msg
                Add-UserReportEntry -SamAccountName $TargetUserSamAccountName -Status "Fehler" -Detail $msg
                return # Abbruch des Modus
            }

            # 3. Funktion zum Anwenden der Eigenschaften aufrufen
            Apply-ADUserProperties -ReferenceUser $referenceUserObject `
                                   -TargetUser $targetUserObject
                                   # KORREKTUR v6.6: Explizite Parameterübergabe entfernt

            # Report-Einträge werden innerhalb von Apply-ADUserProperties hinzugefügt
            Write-Log -Level Info -Message "Modus ApplyPropertiesToExistingUser abgeschlossen für Ziel '$($targetUserObject.SamAccountName)'."

        } # End ApplyPropertiesToExistingUser

        # --- Modus: ExportUserData (Allgemeiner Export) ---
        'ExportUserData' {
             Write-Log -Level Info -Message "Starte Modus: ExportUserData"

             # 1. Eigenschaften für Get-ADUser zusammenstellen
             $defaultExportProperties = @(
                 'SamAccountName', 'Name', 'GivenName', 'Surname', 'DisplayName',
                 'UserPrincipalName', 'Enabled', 'DistinguishedName'
                 # 'MemberOf' wird immer benötigt
             )
             $allPropertiesToGet = ($defaultExportProperties + $PropertiesToExport + 'MemberOf', 'DistinguishedName' | Select-Object -Unique)
             Write-Verbose "Folgende Eigenschaften werden für den Export abgefragt: $($allPropertiesToGet -join ', ')"

             # 2. Suchbasen bestimmen (OUFilter oder SearchBaseOU)
             $searchBases = @()
             $domainDN = (Get-ADDomain).DistinguishedName
             if ($OUFilter) {
                 Write-Log -Level Info -Message "Suche OUs mit Filter '$OUFilter'..."
                 try {
                     $foundOUs = Get-ADOrganizationalUnit -Filter "Name -like '$OUFilter'" -SearchBase $domainDN -SearchScope Subtree -ErrorAction Stop
                     if ($foundOUs) {
                         $searchBases = $foundOUs.DistinguishedName
                         Write-Log -Level Info -Message "$($searchBases.Count) passende OUs gefunden: $($searchBases -join '; ')"
                     } else {
                         Write-Log -Level Warning -Message "Keine OUs für Filter '$OUFilter' gefunden."
                         # Suche in der gesamten Domäne fortsetzen, wenn keine passende OU gefunden wurde
                         $searchBases = @($domainDN)
                         Write-Log -Level Warning -Message "Durchsuche stattdessen die gesamte Domäne."
                     }
                 } catch {
                     Write-Log -Level Error -Message "Fehler beim Suchen von OUs mit Filter '$OUFilter': $_"
                     return # Abbruch bei Fehler in OU-Suche
                 }
             } elseif ($SearchBaseOU) {
                 Write-Verbose "Verwende spezifische SearchBase OU: $SearchBaseOU"
                 $searchBases = @($SearchBaseOU)
             } else {
                 Write-Verbose "Kein OU-Filter oder SearchBaseOU angegeben, durchsuche gesamte Domäne."
                 $searchBases = @($domainDN)
             }

             # 3. Benutzer suchen (innerhalb der bestimmten Suchbasen)
             $allFoundUsers = @() # Standard PowerShell Array
             $getADUserParams = @{ Properties = $allPropertiesToGet }

             # Filter für Get-ADUser vorbereiten
             if ($IdentityFilter -notlike '*?' -and $IdentityFilter -notlike '*`*' -and $IdentityFilter -notlike '`**') {
                 # Versuche zuerst Identity, wenn keine Wildcards enthalten sind (potenziell schneller)
                  try {
                      Write-Verbose "Versuche Get-ADUser -Identity '$IdentityFilter'"
                      $userById = Get-ADUser -Identity $IdentityFilter @getADUserParams -ErrorAction Stop
                      # Stelle sicher, dass der gefundene Benutzer in einer der erlaubten SearchBases liegt
                      $userIsInAllowedOU = $false
                      foreach($base in $searchBases){
                          # Prüfe ob der DN des Benutzers mit ",$base" endet (korrekte Prüfung für OU-Zugehörigkeit)
                          if($userById.DistinguishedName -like "*,$base"){
                              $userIsInAllowedOU = $true
                              break
                          }
                      }
                      if($userIsInAllowedOU){
                          $allFoundUsers += $userById # Fügt das einzelne Objekt hinzu
                          Write-Log -Level Info -Message "Benutzer '$IdentityFilter' über -Identity gefunden (in erlaubter OU)."
                      } else {
                           Write-Log -Level Info -Message "Benutzer '$IdentityFilter' über -Identity gefunden, aber nicht in einer der spezifizierten Suchbasen ($($searchBases -join '; '))."
                           # Setze Filter für Fallback-Suche, falls Identity gefunden wurde, aber nicht in der richtigen OU
                           $getADUserParams.Filter = "SamAccountName -eq '$IdentityFilter' -or Name -eq '$IdentityFilter' -or UserPrincipalName -eq '$IdentityFilter'"
                      }
                  } catch {
                      Write-Verbose "Benutzer '$IdentityFilter' nicht über -Identity gefunden, versuche Filter..."
                      # Wenn -Identity fehlschlägt, Fallback auf -Filter
                      $getADUserParams.Filter = "SamAccountName -eq '$IdentityFilter' -or Name -eq '$IdentityFilter' -or UserPrincipalName -eq '$IdentityFilter'"
                  }
             } else {
                 # Verwende -Filter mit -like für Wildcards
                 $filterString = "SamAccountName -like '$IdentityFilter' -or Name -like '$IdentityFilter' -or DisplayName -like '$IdentityFilter'"
                 $getADUserParams.Filter = $filterString
                 Write-Verbose "Verwende Filter: $filterString"
             }

             # Wenn Benutzer nicht schon über -Identity gefunden wurde ODER Identity gefunden wurde aber nicht in der richtigen OU, suche mit -Filter in den Suchbasen
             if ($allFoundUsers.Count -eq 0 -and $getADUserParams.ContainsKey('Filter')) {
                 Write-Log -Level Info -Message "Suche Benutzer mit Filter '$($getADUserParams.Filter)' in $($searchBases.Count) Suchbasis(en)..."
                 foreach ($base in $searchBases) {
                     $getADUserParams.SearchBase = $base
                     try {
                         Write-Verbose "Durchsuche Basis: $base"
                         $usersInBase = Get-ADUser @getADUserParams -ErrorAction SilentlyContinue # Fehler hier nicht kritisch, weiter mit nächster Basis
                         if ($usersInBase) {
                             Write-Verbose "$($usersInBase.Count) Benutzer in Basis '$base' gefunden."
                             $allFoundUsers += $usersInBase
                         }
                     } catch {
                         Write-Log -Level Warning -Message "Fehler beim Durchsuchen der Basis '$base': $_"
                     }
                 }
             }

             # Duplikate entfernen
             $uniqueFoundUsers = $allFoundUsers | Select-Object -Unique -Property DistinguishedName
             Write-Log -Level Info -Message "Insgesamt $($uniqueFoundUsers.Count) eindeutige Benutzer gefunden."

             if ($uniqueFoundUsers.Count -eq 0) {
                 Write-Log -Level Warning -Message "Keine Benutzer für die angegebenen Kriterien gefunden."
                 return
             }

             # 4. Daten für den Export aufbereiten
             $exportData = [System.Collections.Generic.List[PSObject]]::new()
             Write-Log -Level Info -Message "Bereite Daten für den Export vor..."
             foreach ($userDN in ($uniqueFoundUsers | Select-Object -ExpandProperty DistinguishedName)) { # Iteriere über DNs
                 try {
                     # Hole das vollständige Benutzerobjekt erneut
                     $user = Get-ADUser -Identity $userDN -Properties $allPropertiesToGet -ErrorAction Stop
                     Write-Verbose "Verarbeite Benutzer für Export: $($user.SamAccountName)"

                     $userExportObject = [ordered]@{
                         SamAccountName = $user.SamAccountName
                         Name = $user.Name
                         GivenName = $user.GivenName
                         Surname = $user.Surname
                         DisplayName = $user.DisplayName
                         UserPrincipalName = $user.UserPrincipalName
                         Enabled = $user.Enabled
                         DistinguishedName = $user.DistinguishedName
                         OU = ($user.DistinguishedName -split ',', 2)[1] # OU extrahieren
                     }
                     foreach ($prop in $PropertiesToExport) {
                         if ($user.PSObject.Properties.Match($prop).Count -gt 0) { $userExportObject[$prop] = $user.$prop } else { $userExportObject[$prop] = $null }
                     }
                     $groupNames = @()
                     try {
                         if ($user.MemberOf) { $groupNames = $user.MemberOf | ForEach-Object { try { (Get-ADGroup $_ -ErrorAction Stop).Name } catch { Write-Verbose "Konnte Gruppe '$_' nicht auflösen."; "FehlerhafteGruppe:$_" } } | Sort-Object }
                     } catch { Write-Log -Level Warning -Message "Fehler beim Auflösen der Gruppen für '$($user.SamAccountName)': $_" }
                     $userExportObject['GroupNames'] = $groupNames -join ','
                     $exportData.Add([PSCustomObject]$userExportObject)
                 } catch {
                      Write-Log -Level Warning -Message "Fehler beim erneuten Abrufen oder Verarbeiten des Benutzers mit DN '$userDN': $_"
                 }
             } # End foreach userDN

             # 5. Nach CSV exportieren
             if($exportData.Count -eq 0){
                 Write-Log -Level Warning -Message "Keine Daten zum Exportieren nach der Aufbereitung vorhanden (möglicherweise Fehler bei der Eigenschaftsextraktion)."
                 return
             }
             $finalExportPath = $ExportCsvPath # Mandatory Parameter

             Write-Log -Level Info -Message "Exportiere $($exportData.Count) Benutzerdatensätze nach '$finalExportPath'."
             if ($PSCmdlet.ShouldProcess($finalExportPath, "Benutzerdaten exportieren")) {
                 try {
                     $exportDir = Split-Path -Path $finalExportPath -Parent
                     if (-not (Test-Path $exportDir -PathType Container)) {
                         Write-Verbose "Erstelle Export-Verzeichnis: $exportDir"
                         New-Item -Path $exportDir -ItemType Directory -Force:$true -ErrorAction Stop | Out-Null
                     }
                     $exportData | Export-Csv -Path $finalExportPath -Delimiter ';' -NoTypeInformation -Encoding UTF8 -Force -ErrorAction Stop
                     Write-Log -Level Info -Message "Export erfolgreich abgeschlossen: $finalExportPath"
                 } catch {
                     Write-Log -Level Error -Message "Fehler beim Exportieren der Daten nach '$finalExportPath': $_"
                 }
             } else {
                 Write-Log -Level Info -Message "Export nach '$finalExportPath' übersprungen (ShouldProcess)."
             }

        } # End ExportUserData

        # --- Modus: ExportLKennung (Spezifischer Export) ---
        'ExportLKennung' {
             Write-Log -Level Info -Message "Starte Modus: ExportLKennung"

             # 1. Ziel-OUs finden
             $targetOUDNs = @()
             $domainDN = (Get-ADDomain).DistinguishedName
             Write-Log -Level Info -Message "Suche nach OUs: $($LKennungOUNames -join ', ')"
             foreach ($ouName in $LKennungOUNames) {
                 try {
                     # Suche OU basierend auf dem Namen, unabhängig von der Ebene
                     $foundOUs = Get-ADOrganizationalUnit -Filter "Name -eq '$ouName'" -SearchBase $domainDN -SearchScope Subtree -ErrorAction Stop
                     if ($foundOUs) {
                         # Es könnten mehrere OUs mit gleichem Namen existieren, füge alle hinzu
                         foreach($foundOU in $foundOUs){
                             $targetOUDNs += $foundOU.DistinguishedName
                             Write-Verbose "OU '$ouName' gefunden: $($foundOU.DistinguishedName)"
                         }
                     } else {
                         Write-Log -Level Warning -Message "OU mit Namen '$ouName' nicht gefunden."
                     }
                 } catch {
                      Write-Log -Level Warning -Message "Fehler beim Suchen der OU '$ouName': $_"
                 }
             }
             # Eindeutige DNs sicherstellen
             $targetOUDNs = $targetOUDNs | Select-Object -Unique

             if ($targetOUDNs.Count -eq 0) {
                 Write-Log -Level Error -Message "Keine der angegebenen OUs gefunden ($($LKennungOUNames -join ', ')). Export wird abgebrochen."
                 return
             }
             Write-Log -Level Info -Message "Gefundene OU DNs für die Suche: $($targetOUDNs -join '; ')"

             # 2. Eigenschaften für Get-ADUser zusammenstellen (wie bei ExportUserData)
             $defaultExportProperties = @(
                 'SamAccountName', 'Name', 'GivenName', 'Surname', 'DisplayName',
                 'UserPrincipalName', 'Enabled', 'DistinguishedName'
             )
             $allPropertiesToGet = ($defaultExportProperties + $PropertiesToExport + 'MemberOf', 'DistinguishedName' | Select-Object -Unique)
             Write-Verbose "Folgende Eigenschaften werden für den L-Kennung Export abgefragt: $($allPropertiesToGet -join ', ')"

             # 3. Benutzer in den gefundenen OUs suchen
             $allFoundUsers = @() # Standard PowerShell Array verwenden
             Write-Log -Level Info -Message "Suche Benutzer mit LDAP-Filter '$LKennungLDAPFilter' in den gefundenen OUs..."
             foreach ($ouDN in $targetOUDNs) {
                 try {
                     Write-Verbose "Durchsuche OU: $ouDN"
                     # SearchScope Subtree, um auch Benutzer in Unter-OUs zu finden
                     $usersInOU = Get-ADUser -LDAPFilter $LKennungLDAPFilter -SearchBase $ouDN -Properties $allPropertiesToGet -SearchScope Subtree -ErrorAction Stop
                     if ($usersInOU) {
                         Write-Verbose "$($usersInOU.Count) Benutzer in OU '$ouDN' (und Unter-OUs) gefunden."
                         # KORREKTUR (v6.1): += verwenden statt AddRange
                         $allFoundUsers += $usersInOU
                     } else {
                          Write-Verbose "Keine passenden Benutzer in OU '$ouDN' (und Unter-OUs) gefunden."
                     }
                 } catch {
                     # Fehler beim Durchsuchen loggen und mit nächster OU fortfahren
                     $errorMessage = "Fehler beim Durchsuchen der OU '$ouDN' mit Filter '$LKennungLDAPFilter': $_"
                     Write-Log -Level Warning -Message $errorMessage # Fehler wird nun geloggt
                 }
             }

             # Duplikate entfernen, falls OUs verschachtelt waren oder Benutzer in mehreren gefunden wurden
             # KORREKTUR (v6.2): Eindeutige DNs extrahieren, DANN die vollen Objekte neu laden
             $uniqueUserDNs = $allFoundUsers | Select-Object -ExpandProperty DistinguishedName -Unique
             Write-Log -Level Info -Message "Insgesamt $($uniqueUserDNs.Count) eindeutige Benutzer-DNs gefunden."

             if ($uniqueUserDNs.Count -eq 0) {
                 Write-Log -Level Warning -Message "Keine Benutzer für die angegebenen Kriterien gefunden."
                 return
             }

             # 4. Daten für den Export aufbereiten (JETZT mit vollständigen Objekten)
             $exportData = [System.Collections.Generic.List[PSObject]]::new()
             Write-Log -Level Info -Message "Bereite Daten für den L-Kennung Export vor..."
             foreach ($dn in $uniqueUserDNs) {
                 try {
                     # Hole das vollständige Benutzerobjekt erneut
                     $user = Get-ADUser -Identity $dn -Properties $allPropertiesToGet -ErrorAction Stop
                     Write-Verbose "Verarbeite Benutzer für Export: $($user.SamAccountName)"

                     $userExportObject = [ordered]@{
                         SamAccountName = $user.SamAccountName
                         Name = $user.Name
                         GivenName = $user.GivenName
                         Surname = $user.Surname
                         DisplayName = $user.DisplayName
                         UserPrincipalName = $user.UserPrincipalName
                         Enabled = $user.Enabled
                         DistinguishedName = $user.DistinguishedName
                         OU = ($user.DistinguishedName -split ',', 2)[1] # OU extrahieren
                     }
                     foreach ($prop in $PropertiesToExport) {
                         if ($user.PSObject.Properties.Match($prop).Count -gt 0) { $userExportObject[$prop] = $user.$prop } else { $userExportObject[$prop] = $null }
                     }
                     $groupNames = @()
                     try {
                         if ($user.MemberOf) { $groupNames = $user.MemberOf | ForEach-Object { try { (Get-ADGroup $_ -ErrorAction Stop).Name } catch { Write-Verbose "Konnte Gruppe '$_' nicht auflösen."; "FehlerhafteGruppe:$_" } } | Sort-Object }
                     } catch { Write-Log -Level Warning -Message "Fehler beim Auflösen der Gruppen für '$($user.SamAccountName)': $_" }
                     $userExportObject['GroupNames'] = $groupNames -join ','
                     $exportData.Add([PSCustomObject]$userExportObject)
                 } catch {
                      Write-Log -Level Warning -Message "Fehler beim erneuten Abrufen oder Verarbeiten des Benutzers mit DN '$dn': $_"
                 }
             } # End foreach dn

             # 5. Nach CSV exportieren
             if($exportData.Count -eq 0){
                 Write-Log -Level Warning -Message "Keine Daten zum Exportieren nach der Aufbereitung vorhanden (möglicherweise Fehler bei der Eigenschaftsextraktion)."
                 return
             }
             $finalExportPath = $LKennungExportCsvPath # Mandatory Parameter für diesen Modus

             Write-Log -Level Info -Message "Exportiere $($exportData.Count) L-Kennung Benutzerdatensätze nach '$finalExportPath'."
             if ($PSCmdlet.ShouldProcess($finalExportPath, "L-Kennung Benutzerdaten exportieren")) {
                 try {
                     $exportDir = Split-Path -Path $finalExportPath -Parent
                     if (-not (Test-Path $exportDir -PathType Container)) {
                         Write-Verbose "Erstelle Export-Verzeichnis: $exportDir"
                         New-Item -Path $exportDir -ItemType Directory -Force:$true -ErrorAction Stop | Out-Null
                     }
                     $exportData | Export-Csv -Path $finalExportPath -Delimiter ';' -NoTypeInformation -Encoding UTF8 -Force -ErrorAction Stop
                     Write-Log -Level Info -Message "L-Kennung Export erfolgreich abgeschlossen: $finalExportPath"
                 } catch {
                     Write-Log -Level Error -Message "Fehler beim Exportieren der L-Kennung Daten nach '$finalExportPath': $_"
                 }
             } else {
                 Write-Log -Level Info -Message "L-Kennung Export nach '$finalExportPath' übersprungen (ShouldProcess)."
             }

        } # End ExportLKennung

        default {
            # Sollte nicht passieren bei korrekter Parameternutzung
            $msg = "Unbekannter oder keiner der Hauptmodi wurde ausgewählt. Verwenden Sie -CopySingleUser, -CreateUsersFromCSV, -ApplyPropertiesToExistingUser, -ExportUserData oder -ExportLKennung."
            Write-Log -Level Error -Message $msg
            Add-UserReportEntry -SamAccountName "(Skript)" -Status "Fehler" -Detail $msg
        }
    } # End Switch ParameterSetName

    Write-Verbose "Prozess-Phase abgeschlossen."

} # End Process Block

end {
    $scriptEndTime = Get-Date
    $duration = New-TimeSpan -Start $scriptStartTime -End $scriptEndTime
    Write-Verbose "Beginne End-Phase."

    # Aktionsbericht schreiben, wenn Daten vorhanden (nur für Copy, Create, Apply)
    if ($global:userReportData -ne $null -and $global:userReportData.Count -gt 0) {
        Write-Log -Level Info -Message "Schreibe Aktionsbericht nach '$($global:fullUserReportPath)'..."
        if ($PSCmdlet.ShouldProcess($global:fullUserReportPath, "Aktionsbericht exportieren")) {
            try {
                $global:userReportData | Export-Csv -Path $global:fullUserReportPath -Delimiter ';' -NoTypeInformation -Encoding UTF8 -Force -ErrorAction Stop
                Write-Log -Level Info -Message "Aktionsbericht erfolgreich geschrieben."
            } catch {
                Write-Log -Level Error -Message "Fehler beim Schreiben des Aktionsberichts nach '$($global:fullUserReportPath)': $_"
            }
        } else {
             Write-Log -Level Info -Message "Schreiben des Aktionsberichts übersprungen (ShouldProcess)."
        }
    } elseif ($global:userReportData -ne $null) { # Report wurde initialisiert, aber keine Daten
         Write-Log -Level Info -Message "Keine Daten für Aktionsbericht vorhanden."
    }

    Write-Log -Level Info -Message "Skriptausführung beendet. Gesamtdauer: $($duration.ToString('g'))"
    Write-Log -Level Info -Message "Log-Datei: $fullLogPath"
    Write-Verbose "End-Phase abgeschlossen."
} # End End Block
