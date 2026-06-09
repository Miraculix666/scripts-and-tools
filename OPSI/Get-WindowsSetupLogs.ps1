<#
.SYNOPSIS
Sammelt alle relevanten Protokolldateien einer unbeaufsichtigten Windows-Installation.

.DESCRIPTION
Dieses Skript kopiert wichtige Log-Dateien aus den Installationsphasen (Setup, Sysprep, OOBE)
von einem lokalen oder remote erreichbaren Windows-System. Es nutzt den Admin-Share (C$)
für den Remote-Zugriff und unterstützt die Übergabe von alternativen Anmeldeinformationen.
Die Ergebnisse werden in einem lokalen Ordner gespeichert, der nach dem Quell-Computer benannt ist.

.PARAMETER ComputerName
Der Hostname oder die IP-Adresse des Zielcomputers. Bei Weglassen wird der lokale Computer verwendet.

.PARAMETER DestinationPath
Der Basispfad, in dem die gesammelten Logs gespeichert werden sollen.
Wenn nicht angegeben, ist der Standard der Ordner 'CollectedLogs' im Verzeichnis des Skripts.

.PARAMETER Credential
Ein PSCredential-Objekt mit alternativen Anmeldeinformationen für den Remote-Zugriff.

.PARAMETER Silent
Unterdrückt die verbose Ausgabe (Logging) des Skripts. Standardmäßig ist Verbose aktiviert.

.EXAMPLE
# Lokale Logs sammeln (Zielordner: Skriptpfad\CollectedLogs\LOKALER_PC)
.\Get-WindowsSetupLogs.ps1

.EXAMPLE
# Remote-Logs sammeln mit interaktiver Abfrage der Anmeldeinformationen
.\Get-WindowsSetupLogs.ps1 -ComputerName 'SRV01'

.EXAMPLE
# Remote-Logs sammeln unter Verwendung spezifischer Credentials
$creds = Get-Credential -UserName 'DOMAIN\Admin' -Message 'Geben Sie das Passwort für den Remote-Zugriff ein'
.\Get-WindowsSetupLogs.ps1 -ComputerName 'SRV02' -Credential $creds

.NOTES
Autor: PS-Coding
Version: 1.2 (Fehlerbehebung: Fehlerhafte Pfadberechnung bei leerem $PSCommandPath korrigiert.)
Quellen: AI-discovered (Typische Windows Setup Log-Pfade)
         User-provided (Remote-Zugriff über Admin-Share, dynamische Credential-Verwendung)
Sprache/Lokalisierung: Deutsch (Verbose-Ausgabe, Pfadnamen)
#>
[CmdletBinding(DefaultParameterSetName = 'Remote', SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'Remote')]
    [string]$ComputerName = $env:COMPUTERNAME,

    [Parameter(Mandatory = $false)]
    [string]$DestinationPath, # Standardwert hier entfernt, wird im Skriptkörper gesetzt

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter(Mandatory = $false)]
    [switch]$Silent
)

# --- 1. Konfiguration und Initialisierung ---

# Aktiviert Verbose-Ausgabe, wenn der Silent-Parameter nicht gesetzt ist
if (-not $Silent) {
    $VerbosePreference = 'Continue'
}
Write-Verbose "--- Skript-Initialisierung gestartet (PS-Coding Log Collector) ---"

# Korrektur: Setzt den dynamischen Standardpfad, falls der Benutzer ihn nicht angegeben hat.
if (-not $PSBoundParameters.ContainsKey('DestinationPath')) {
    
    # $PSScriptRoot ist die sicherste Methode in PS 3.0+ (was 5.1 einschließt), um den Ordnerpfad des Skripts zu erhalten.
    $ScriptDir = $PSScriptRoot
    
    if (-not $ScriptDir) {
        # Fallback für spezielle Ausführungskontexte, bei denen $PSScriptRoot nicht verfügbar ist.
        # Wir verwenden den aktuellen Ausführungspfad (Get-Location) als letzten Ausweg.
        try {
            $ScriptDir = (Get-Location -ErrorAction Stop).Path
            Write-Warning "ACHTUNG: Skriptpfad (\$PSScriptRoot) war nicht verfügbar. Verwende den aktuellen Ausführungspfad ('$ScriptDir') als Basis für die Logs."
        } catch {
             Write-Error "Schwerer Fehler: Der Ausführungspfad konnte nicht ermittelt werden. Bitte geben Sie den Parameter -DestinationPath manuell an."
             exit 1
        }
    }
    
    $DestinationPath = Join-Path -Path $ScriptDir -ChildPath 'CollectedLogs'
    Write-Verbose "Standard-Zielpfad basierend auf Skriptpfad festgelegt: $($DestinationPath)"
}

# Definiert die typischen Speicherorte für Windows Setup Logs relativ zum Systemlaufwerk
$LogSources = @(
    '\Windows\Panther'
    '\Windows\Panther\UnattendGC'
    '\Windows\System32\sysprep\Panther'
    '\Windows\Setup\Scripts'
    '\Windows\Inf'
    '\Windows\Debug'
)

# Erstellt den vollständigen Zielordnerpfad
$TargetHostName = $ComputerName.ToUpper().Trim() # Hostname in Großbuchstaben für Konsistenz
$FinalDestination = Join-Path -Path $DestinationPath -ChildPath $TargetHostName

Write-Verbose "Zielcomputer: $($ComputerName)"
Write-Verbose "Zielpfad für die Logs: $($FinalDestination)"

# --- 2. Pfad- und Zugriffsprüfung ---

# Überprüft, ob der Zielordner erstellt werden kann
if (-not (Test-Path -Path $FinalDestination)) {
    Write-Verbose "Erstelle Zielverzeichnis: $($FinalDestination)"
    try {
        $null = New-Item -Path $FinalDestination -ItemType Directory -ErrorAction Stop
    } catch {
        Write-Error "Fehler beim Erstellen des Zielordners '$FinalDestination'. Überprüfen Sie die Berechtigungen."
        Write-Error $_.Exception.Message
        exit 1
    }
}

# Bestimmt den UNC-Pfad zum Admin-Share
$UNCPath = "\\$ComputerName\c$"
Write-Verbose "Versuche Remote-Zugriff über UNC-Pfad: $($UNCPath)"

# --- 3. Credential- und Verbindungs-Handling ---

function Test-UNCPathAccess {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Cred
    )

    Write-Verbose "Teste Zugriff auf '$Path'..."

    # Wenn Credentials vorhanden sind, verwende sie zur temporären Authentifizierung
    if ($Cred) {
        Write-Verbose "Verwende übergebene Credentials für den Zugriff."
        try {
            # Erfordert die PowerShell Remoting (WinRM) Umgebung
            $UserName = $Cred.UserName
            $Password = $Cred.GetNetworkCredential().Password
            
            # Verwendung von 'net use' zur temporären Authentifizierung des UNC-Pfades
            # Dies ist robuster in Umgebungen ohne PS-Remoting
            Write-Verbose "Führe 'net use' aus, um Verbindung herzustellen..."
            
            $netUseResult = net.exe use $Path $Password "/user:$UserName" /persistent:no 2>&1
            
            if ($LASTEXITCODE -ne 0) {
                 Write-Verbose "net use fehlgeschlagen (ExitCode $LASTEXITCODE). Ergebnis: $($netUseResult | Out-String)"
                 return $false
            }
            Write-Verbose "Netzwerkverbindung temporär erfolgreich hergestellt."
            return $true
        } catch {
            Write-Warning "Fehler bei der Credential-Authentifizierung: $($_.Exception.Message)"
            return $false
        }
    }

    # Testet den Zugriff ohne explizite Credentials (aktueller Benutzer)
    try {
        # Einfacher Test-Path ohne Credentials
        $TestPath = Test-Path -Path $Path -ErrorAction Stop
        if ($TestPath) {
            Write-Verbose "Zugriff mit aktuellen Benutzerrechten erfolgreich."
            return $true
        }
    } catch {
        Write-Verbose "Zugriff mit aktuellen Benutzerrechten fehlgeschlagen: $($_.Exception.Message)"
        return $false
    }
    return $false
}

$IsAccessOK = $false

# 3a. Versuch mit übergebenen Credentials
if ($Credential) {
    $IsAccessOK = Test-UNCPathAccess -Path $UNCPath -Cred $Credential
}

# 3b. Versuch mit aktuellen Benutzerrechten (nur wenn 3a fehlgeschlagen ist)
if (-not $IsAccessOK) {
    Write-Verbose "Prüfe Zugriff mit aktuellen Benutzerrechten..."
    $IsAccessOK = Test-UNCPathAccess -Path $UNCPath
}

# 3c. Interaktive Abfrage, wenn alle Versuche fehlschlagen und keine Credentials übergeben wurden
if (-not $IsAccessOK -and -not $Credential) {
    Write-Warning "Zugriff auf '$ComputerName' mit aktuellen Berechtigungen fehlgeschlagen."
    
    # Interaktiver Modus: Warte auf neue Credentials
    while (-not $IsAccessOK) {
        try {
            $credPrompt = Read-Host "Möchten Sie alternative Anmeldeinformationen eingeben? (J/N)"
            if ($credPrompt -notmatch '^[Jj]$') {
                Write-Error "Zugriff verweigert. Das Skript wird abgebrochen."
                exit 1
            }
            
            $Credential = Get-Credential -Message "Geben Sie die Credentials für den Zugriff auf '$ComputerName' ein."
            $IsAccessOK = Test-UNCPathAccess -Path $UNCPath -Cred $Credential
        } catch {
            Write-Error "Fehler bei der Credential-Eingabe: $($_.Exception.Message)"
        }
    }
} elseif (-not $IsAccessOK) {
     # Fehler, wenn Credentials übergeben, aber Zugriff fehlgeschlagen ist
     Write-Error "Der Remote-Zugriff auf '$ComputerName' mit den bereitgestellten Credentials ist fehlgeschlagen. Das Skript wird abgebrochen."
     exit 1
}

# --- 4. Log-Dateien kopieren (Aktion) ---

Write-Verbose "`nStarte Kopiervorgang der Log-Dateien von $ComputerName..."

$FileCounter = 0
foreach ($Source in $LogSources) {
    $RemoteSourcePath = Join-Path -Path $UNCPath -ChildPath $Source
    
    # Der Ordnerpfad relativ zum Ziel (z.B. Panther, Inf)
    $RelativeDestination = Split-Path -Path $Source -Leaf
    $CurrentDestination = Join-Path -Path $FinalDestination -ChildPath $RelativeDestination

    Write-Verbose "Prüfe Remote-Pfad: $($RemoteSourcePath)"

    # Test-Path muss auf dem UNC-Pfad ausgeführt werden, der jetzt authentifiziert sein sollte
    if (Test-Path -Path $RemoteSourcePath) {
        Write-Verbose "Verzeichnis gefunden. Kopiere Dateien nach: $($CurrentDestination)"
        
        # Erstellt das Unterverzeichnis im Ziel, falls nötig
        if (-not (Test-Path -Path $CurrentDestination)) {
             $null = New-Item -Path $CurrentDestination -ItemType Directory -Force -ErrorAction Stop
        }

        # Kopiert alle Dateien im Verzeichnis rekursiv
        try {
            # Hier wird -Recurse verwendet, um z.B. Unterordner wie 'UnattendGC' in Panther zu finden
            $ItemsToCopy = Get-ChildItem -Path $RemoteSourcePath -File -Recurse -ErrorAction SilentlyContinue
            
            if ($ItemsToCopy.Count -gt 0) {
                 foreach ($Item in $ItemsToCopy) {
                    # Berechne den relativen Pfad der Datei zum RemoteSourcePath
                    $RelativePath = $Item.FullName.Substring($RemoteSourcePath.Length).TrimStart('\')
                    
                    # Erzeuge den vollständigen Zielpfad, um die Unterordnerstruktur beizubehalten
                    $TargetFilePath = Join-Path -Path $CurrentDestination -ChildPath $RelativePath

                    # Stelle sicher, dass der Ziel-Unterordner existiert
                    $TargetFolder = Split-Path -Path $TargetFilePath -Parent
                    if (-not (Test-Path -Path $TargetFolder)) {
                         $null = New-Item -Path $TargetFolder -ItemType Directory -Force -ErrorAction Stop
                    }

                    # Der Kopiervorgang
                    Copy-Item -Path $Item.FullName -Destination $TargetFilePath -Force -ErrorAction Stop
                    $FileCounter++
                 }
                 Write-Verbose "Insgesamt $($ItemsToCopy.Count) Dateien aus '$Source' zur Kopie gefunden."
            } else {
                Write-Verbose "Keine Dateien im Ordner '$Source' gefunden. (Nur Dateien, keine Ordner)"
            }
        } catch {
            Write-Warning "Kopierfehler im Ordner '$Source': $($_.Exception.Message)"
        }

    } else {
        Write-Verbose "Remote-Pfad '$Source' nicht gefunden oder nicht lesbar. Überspringe."
    }
}

# Optional: Trennt die temporäre Netzwerkverbindung (falls 'net use' verwendet wurde)
# ACHTUNG: Der 'net use' Befehl bleibt oft aktiv, bis das Skript beendet ist oder explizit getrennt wird.
# Wir versuchen die Trennung nur, wenn eine Credential verwendet wurde.
if ($Credential) {
    Write-Verbose "Trenne temporäre Netzwerkverbindung zu '$UNCPath' (falls eingerichtet)."
    # Verwende 'net use /delete', falls die Verbindung von net use erstellt wurde
    net.exe use $UNCPath /delete 2>&1 | Out-Null
}

# --- 5. Abschluss und Erfolgsmeldung ---

Write-Host "---"
if ($FileCounter -gt 0) {
    # Deutsche Lokalisierung für das Datum
    $GermanDate = (Get-Date).ToString("dd.MM.yyyy HH:mm:ss")
    Write-Host "✅ Erfolg ($GermanDate): Protokolle von '$ComputerName' erfolgreich gesammelt."
    Write-Host "Insgesamt wurden $FileCounter Dateien kopiert."
    Write-Host "Speicherort: '$FinalDestination'"
} else {
    Write-Warning "Es wurden keine Protokolldateien gefunden oder kopiert."
    Write-Host "Prüfen Sie, ob '$ComputerName' erreichbar ist und ob die Log-Ordner existieren."
}

Write-Verbose "--- Skript-Ausführung beendet. ---"
