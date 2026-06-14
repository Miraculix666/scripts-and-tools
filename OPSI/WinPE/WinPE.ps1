<#
    .SYNOPSIS
    Standardkonformer WinPE Builder V14.0 für OPSI 4.3 - OPSI Edition
    
    .CHANGELOG
    V14.0 (2026-02-13):
    - OPSI-konforme Struktur (winpe/ und winpe_uefi/)
    - Automatischer Toolkit-Download mit Smart-Caching
    - BCD-Repair für UEFI + BIOS (dual-boot ready)
    - WinPE-Komponenten Integration (PowerShell, WMI, NetFx)
    - Interaktives LaunchMenu.ps1 für Toolkit-Zugang
    - Deutsche Lokalisierung (Kultur, Tastatur, UI)
    - ADK Auto-Detection
    - ISO-Cache-Logik für schnellere Builds
    
    V13.5 (2026-02-13):
    - Automatische Elevation (UAC) wenn ohne Admin-Rechte gestartet
    - Umfassende Parameter-Validierung
    - Strukturiertes Logging mit Zeitstempeln
    - Verbesserte Fehlerbehandlung
    - Fortschrittsanzeigen für lange Operationen
    
    ANFORDERUNGEN:
    - Windows ADK für Windows 11 installiert
    - Erstellung WinPE Ordnerstruktur & ISO
    - Integration Tools (Explorer++, Micro, NPP, Sysinternals)
    - Caching-Logik für ISO-Inhalte
    - Verbose Debugging & Inspect-Kopie (_WIM_CONTENT_INSPECT)
    - Fehlerhandling für DISM (Locking/Permissions)
#>

param (
    [Parameter(Mandatory = $true)] 
    [string]$opsiexportpath,
    
    [Parameter(Mandatory = $true)] 
    [string]$WorkingPath,
    
    [Parameter(Mandatory = $true)] 
    [string]$DriverSource,
    
    [Parameter(Mandatory = $true)] 
    [string]$sourceiso,
    
    [Parameter(Mandatory = $false)]
    [string]$ToolsSource = "AUTO",  # AUTO = Download, Pfad = Lokale Kopie
    
    [Parameter(Mandatory = $false)]
    [switch]$UseCache,  # ISO-Cache wiederverwenden
    
    [Parameter(Mandatory = $false)]
    [switch]$SkipToolkitDownload,  # Toolkit-Download überspringen
    
    [Parameter(Mandatory = $false)]
    [string]$ADKPath,  # ADK-Pfad überschreiben
    
    [Parameter(Mandatory = $false)]
    [switch]$SkipBCDRepair,  # BCD-Repair überspringen
    
    [Parameter(Mandatory = $false)]
    [switch]$KeepMountOpen  # Für Debugging: Mount nicht schließen
)

$ErrorActionPreference = "Stop"
$VerbosePreference = "Continue"

# Pfad-Definitionen
$MountPath = Join-Path $WorkingPath "mount"
$BootWimMod = Join-Path $WorkingPath "boot_modified.wim"
$InspectPath = Join-Path $opsiexportpath "_WIM_CONTENT_INSPECT"
$IsoFile = Join-Path $WorkingPath "OPSI_WinPE_LTSC.iso"
$LogFile = Join-Path $WorkingPath "WinPE_Build_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$ToolkitPath = Join-Path $WorkingPath "ExternalTools"

# Globale Variablen
$script:ADKRoot = $null

# Logging-Funktion
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'SUCCESS')][string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "[$timestamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $logMessage -ErrorAction SilentlyContinue
}

Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "   OPSI WinPE Builder V14.0 (OPSI Edition)    " -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "Log-Datei: $LogFile" -ForegroundColor Gray
Write-Host "" 
Write-Log "OPSI WinPE Builder V14.0 gestartet"

# --- 1. ADMIN-CHECK & AUTOMATIC ELEVATION ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "[ELEVATION] Administrator-Rechte erforderlich. Starte Skript neu mit Elevation..." -ForegroundColor Yellow
    
    # Reconstruct argument list with all bound parameters
    $argList = @()
    $argList += '-NoProfile'
    $argList += '-ExecutionPolicy'
    $argList += 'Bypass'
    $argList += '-File'
    $argList += "`"$($MyInvocation.MyCommand.Path)`""
    
    # Add all bound parameters
    $MyInvocation.BoundParameters.GetEnumerator() | ForEach-Object {
        $argList += "-$($_.Key)"
        if ($_.Value -is [switch]) {
            if ($_.Value) { $argList += "`$true" }
        }
        else {
            $argList += "`"$($_.Value)`""
        }
    }
    
    # Launch elevated process
    try {
        Start-Process -FilePath "powershell.exe" -ArgumentList $argList -Verb RunAs
        Write-Host "[ELEVATION] Elevated Instanz gestartet. Diese Fenster wird geschlossen." -ForegroundColor Green
        exit 0
    }
    catch {
        Write-Error "Fehler beim Starten der elevated Instanz: $_"
        exit 1
    }
}

Write-Host "[ADMIN] Script läuft mit Administrator-Rechten." -ForegroundColor Green
Write-Log "Administrator-Rechte bestätigt"

# --- FUNCTION: Get-ADKPath ---
function Get-ADKPath {
    if ($ADKPath -and (Test-Path $ADKPath)) { 
        Write-Log "Verwende benutzerdefinierten ADK-Pfad: $ADKPath"
        return $ADKPath 
    }
    
    Write-Host " [ADK] Suche Windows Deployment Kit..." -ForegroundColor Yellow
    Write-Log "Suche ADK-Installation"
    
    $Roots = @(
        "C:\Program Files (x86)\Windows Kits\10",
        "C:\Program Files\Windows Kits\10",
        "C:\Program Files (x86)\Windows Kits\11",
        "C:\Program Files\Windows Kits\11"
    )
    
    foreach ($root in $Roots) {
        $adkPath = Join-Path $root "Assessment and Deployment Kit"
        if (Test-Path $adkPath) {
            Write-Host " [OK] ADK gefunden: $root" -ForegroundColor Green
            Write-Log "ADK gefunden: $adkPath"
            return $adkPath
        }
    }
    
    $errMsg = "Windows ADK nicht gefunden! Bitte ADK für Windows 11 installieren."
    Write-Log $errMsg -Level ERROR
    throw $errMsg
}

# --- FUNCTION: Get-ToolkitFiles ---
function Get-ToolkitFiles {
    param([string]$DestPath)
    
    Write-Host " [TOOLKIT] Bereite Werkzeuge vor..." -ForegroundColor Cyan
    Write-Log "Starte Toolkit-Download/Cache-Check"
    
    $Toolkit = @{
        "NotepadPP"    = "https://github.com/notepad-plus-plus/notepad-plus-plus/releases/download/v8.6.2/npp.8.6.2.portable.x64.zip"
        "ExplorerPP"   = "https://explorerplusplus.com/software/explorer++_1.3.5_x64.zip"
        "Sysinternals" = "https://download.sysinternals.com/files/SysinternalsSuite.zip"
        "NirSoft"      = "https://www.nirsoft.net/packages/x64tools.zip"
    }
    
    if (-not (Test-Path $DestPath)) {
        New-Item $DestPath -ItemType Directory -Force | Out-Null
    }
    
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    
    foreach ($app in $Toolkit.Keys) {
        $destDir = Join-Path $DestPath $app
        
        # Smart-Caching: Wenn Ordner existiert und Dateien enthält, überspringen
        if (Test-Path $destDir) {
            $fileCount = (Get-ChildItem $destDir -Recurse -File).Count
            if ($fileCount -gt 0) {
                Write-Host " [CACHE] $app bereits vorhanden ($fileCount Dateien)" -ForegroundColor Gray
                Write-Log "$app gecacht ($fileCount Dateien)" -Level INFO
                continue
            }
        }
        
        try {
            $zipFile = Join-Path $DestPath "$app.zip"
            Write-Host " [DOWNLOAD] $app..." -ForegroundColor Cyan
            Write-Log "Lade $app herunter von: $($Toolkit[$app])" -Level INFO
            
            Invoke-WebRequest $Toolkit[$app] -OutFile $zipFile -UserAgent $userAgent -TimeoutSec 120 -ErrorAction Stop
            [System.IO.Compression.ZipFile]::ExtractToDirectory($zipFile, $destDir)
            Remove-Item $zipFile -Force
            
            Write-Host " [OK] $app heruntergeladen und extrahiert" -ForegroundColor Green
            Write-Log "$app erfolgreich heruntergeladen" -Level SUCCESS
        }
        catch {
            Write-Log ("Fehler beim Download von " + $app + ": $_") -Level ERROR
            Write-Warning "Download fehlgeschlagen: $app (wird übersprungen)"
        }
    }
}

# --- FUNCTION: Add-WinPEComponents ---
function Add-WinPEComponents {
    param(
        [string]$MountPath,
        [string]$ADKRoot
    )
    
    Write-Host " [COMPONENTS] Integriere WinPE Optional Components..." -ForegroundColor Cyan
    Write-Log "Starte WinPE-Komponenten Integration"
    
    $Components = @(
        "WinPE-WMI",
        "WinPE-NetFX",
        "WinPE-Scripting",
        "WinPE-PowerShell",
        "WinPE-StorageWMI",
        "WinPE-DismCmdlets"
    )
    
    $OCPath = Join-Path $ADKRoot "Windows Preinstallation Environment\amd64\WinPE_OCs"
    if (-not (Test-Path $OCPath)) {
        Write-Log "WinPE Optional Components Pfad nicht gefunden: $OCPath" -Level WARNING
        Write-Warning "WinPE OCs nicht gefunden - überspringe Komponentenintegration"
        return
    }
    
    Write-Log "WinPE OCs Pfad: $OCPath"
    
    foreach ($comp in $Components) {
        $cabFile = Get-ChildItem $OCPath -Filter "*$comp*.cab" -Recurse |
        Where-Object { $_.FullName -notmatch "de-de|en-us|fr-fr|es-es" } |
        Select-Object -First 1
        
        if ($cabFile) {
            Write-Host " [CAB] $comp..." -ForegroundColor Gray
            Write-Log "Integriere $comp von: $($cabFile.FullName)" -Level INFO
            
            try {
                Add-WindowsPackage -Path $MountPath -PackagePath $cabFile.FullName -IgnoreCheck -ErrorAction Stop | Out-Null
                Write-Log "$comp erfolgreich integriert" -Level SUCCESS
            }
            catch {
                Write-Log ("Fehler beim Integrieren von " + $comp + ": $_") -Level ERROR
                Write-Warning "Fehler bei $comp (wird übersprungen)"
            }
        }
        else {
            Write-Log "$comp.cab nicht gefunden" -Level WARNING
        }
    }
    
    # Deutsche Sprachpakete hinzufügen
    Write-Host " [LANG] Füge deutsche Sprachpakete hinzu..." -ForegroundColor Gray
    foreach ($comp in $Components) {
        $langCab = Get-ChildItem $OCPath -Filter "*$comp*de-de.cab" -Recurse | Select-Object -First 1
        if ($langCab) {
            Write-Log "Integriere deutsches Language Pack: $comp" -Level INFO
            Add-WindowsPackage -Path $MountPath -PackagePath $langCab.FullName -IgnoreCheck -ErrorAction SilentlyContinue | Out-Null
        }
    }
    
    Write-Host " [OK] WinPE Komponenten integriert" -ForegroundColor Green
    Write-Log "WinPE-Komponenten Integration abgeschlossen" -Level SUCCESS
}

# --- FUNCTION: Set-GermanLocale ---
function Set-GermanLocale {
    param([string]$MountPath)
    
    Write-Host " [LOCALE] Setze deutsche Lokalisierung..." -ForegroundColor Cyan
    Write-Log "Setze deutsche Lokalisierung" -Level INFO
    
    try {
        # Kultur (Datum, Zeit, Währung, Zahlen)
        $null = Dism /Image:$MountPath /Set-AllIntl:de-DE /LogLevel:1
        Write-Log "AllIntl auf de-DE gesetzt" -Level SUCCESS
        
        # Eingabegebietsschema (Tastatur: Deutsch)
        $null = Dism /Image:$MountPath /Set-InputLocale:0407:00000407 /LogLevel:1
        Write-Log "InputLocale auf Deutsch gesetzt (0407:00000407)" -Level SUCCESS
        
        # Systemgebietsschema
        $null = Dism /Image:$MountPath /Set-SysLocale:de-DE /LogLevel:1
        Write-Log "SysLocale auf de-DE gesetzt" -Level SUCCESS
        
        # UI-Sprache
        $null = Dism /Image:$MountPath /Set-UILang:de-DE /LogLevel:1
        Write-Log "UILang auf de-DE gesetzt" -Level SUCCESS
        
        Write-Host " [OK] Deutsche Lokalisierung gesetzt" -ForegroundColor Green
        Write-Log "Deutsche Lokalisierung erfolgreich abgeschlossen" -Level SUCCESS
    }
    catch {
        Write-Log ("Fehler bei Lokalisierung: " + $_) -Level ERROR
        Write-Warning "Lokalisierung teilweise fehlgeschlagen"
    }
}

# --- FUNCTION: Get-LaunchMenuContent ---
function Get-LaunchMenuContent {
    return @'
function Show-Menu {
    Clear-Host
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "      OPSI WinPE Toolkit Launcher V14.0        " -ForegroundColor Cyan
    Write-Host "===============================================" -ForegroundColor Cyan
    
    try {
        $ip = (Get-NetIPAddress -AddressFamily IPv4 | 
               Where-Object { $_.InterfaceAlias -notmatch 'Loopback' } | 
               Select-Object -First 1).IPAddress
    } catch { $ip = "Offline" }
    
    Write-Host " IP-Adresse: $ip" -ForegroundColor Yellow
    Write-Host "-----------------------------------------------"
    Write-Host " [1] Notepad++             [2] Explorer++"
    Write-Host " [3] Sysinternals Suite    [4] NirSoft x64 Tools"
    Write-Host " [5] PowerShell            [6] Command Prompt"
    Write-Host " [7] Registry Editor       [8] Disk Management"
    Write-Host " [9] OPSI Client Agent     [0] Netzwerk-Info"
    Write-Host " [R] REBOOT                [S] SHUTDOWN"
    Write-Host "==============================================="
}

$ToolkitPath = "X:\Program Files\Toolkit"
$ExplorerPP = "$ToolkitPath\ExplorerPP\Explorer++.exe"

do {
    Show-Menu
    $choice = Read-Host "Auswahl"
    
    switch ($choice.ToUpper()) {
        '1' { 
            $npp = "$ToolkitPath\NotepadPP\notepad++.exe"
            if (Test-Path $npp) {
                Start-Process $npp
            } else { Write-Host "Notepad++ nicht gefunden!" -ForegroundColor Red; Start-Sleep 1 }
        }
        '2' { 
            if (Test-Path $ExplorerPP) {
                Start-Process $ExplorerPP
            } else { 
                Start-Process explorer.exe 
            }
        }
        '3' { 
            $sysinternals = "$ToolkitPath\Sysinternals"
            if (Test-Path $sysinternals) {
                if (Test-Path $ExplorerPP) {
                    Start-Process $ExplorerPP $sysinternals
                } else {
                    Start-Process explorer.exe $sysinternals
                }
            } else { Write-Host "Sysinternals nicht gefunden!" -ForegroundColor Red; Start-Sleep 1 }
        }
        '4' { 
            $nirsoft = "$ToolkitPath\NirSoft"
            if (Test-Path $nirsoft) {
                if (Test-Path $ExplorerPP) {
                    Start-Process $ExplorerPP $nirsoft
                } else {
                    Start-Process explorer.exe $nirsoft
                }
            } else { Write-Host "NirSoft nicht gefunden!" -ForegroundColor Red; Start-Sleep 1 }
        }
        '5' { Start-Process powershell.exe -Wait }
        '6' { Start-Process cmd.exe }
        '7' { Start-Process regedit.exe }
        '8' { Start-Process cmd.exe -ArgumentList "/k diskpart" }
        '9' { 
            $opsiClient = "X:\opsi-client-agent\opsi-client-agent.exe"
            if (Test-Path $opsiClient) {
                Write-Host "Starte OPSI Client Agent..." -ForegroundColor Green
                Start-Process $opsiClient "/configure"
            } else { 
                Write-Host "OPSI Client Agent nicht gefunden!" -ForegroundColor Yellow
                Write-Host "Pfad: $opsiClient" -ForegroundColor Gray
                Start-Sleep 2 
            }
        }
        '0' { 
            Clear-Host
            Write-Host "=== Netzwerk-Informationen ===" -ForegroundColor Cyan
            ipconfig /all
            Write-Host ""
            Read-Host "Enter drücken zum Fortfahren..."
        }
        'R' { wpeutil reboot }
        'S' { wpeutil shutdown }
    }
} while ($true)
'@
}

# --- FUNCTION: Get-StartnetContent ---
function Get-StartnetContent {
    return @"
@echo off
wpeinit
echo.
echo ============================================
echo    OPSI WinPE V14.0 - System wird gestartet
echo ============================================
echo.

REM Füge Toolkit zum PATH hinzu
set PATH=%PATH%;X:\Program Files\Toolkit\NotepadPP
set PATH=%PATH%;X:\Program Files\Toolkit\ExplorerPP

REM Versuche OPSI Client Agent zu starten
if exist X:\opsi-client-agent\opsi-client-agent.exe (
    echo [OPSI] Starte OPSI Client Agent...
    X:\opsi-client-agent\opsi-client-agent.exe /configure
) else (
    echo [INFO] OPSI Client Agent nicht gefunden
    echo [INFO] Starte Toolkit-Launcher...
    powershell.exe -ExecutionPolicy Bypass -File X:\Windows\System32\LaunchMenu.ps1
)

cmd.exe
"@
}

# --- FUNCTION: New-LauncherMenu ---
function New-LauncherMenu {
    param([string]$MountPath)

    Write-Host " [MENU] Erstelle Launcher-Menü..." -ForegroundColor Cyan
    Write-Log "Erstelle LaunchMenu.ps1"

    $MenuScript = Get-LaunchMenuContent

    $MenuPath = Join-Path $MountPath "Windows\System32\LaunchMenu.ps1"
    $MenuScript | Out-File $MenuPath -Encoding ASCII -Force
    Write-Log "LaunchMenu.ps1 erstellt: $MenuPath" -Level SUCCESS

    # Anpassung startnet.cmd
    $StartnetPath = Join-Path $MountPath "Windows\System32\startnet.cmd"
    $Startnet = Get-StartnetContent

    $Startnet | Out-File $StartnetPath -Encoding ASCII -Force
    Write-Log "startnet.cmd erstellt/angepasst: $StartnetPath" -Level SUCCESS
    
    Write-Host " [OK] Launcher-Menü erstellt" -ForegroundColor Green
}

# --- FUNCTION: Repair-BCDConfiguration ---
function Repair-BCDConfiguration {
    param([string]$MediaPath)
    
    if ($SkipBCDRepair) {
        Write-Host " [BCD] BCD-Repair übersprungen (SkipBCDRepair)" -ForegroundColor Gray
        Write-Log "BCD-Repair übersprungen" -Level INFO
        return
    }
    
    Write-Host " [BCD] Repariere Boot Configuration (UEFI + BIOS)..." -ForegroundColor Yellow
    Write-Log "Starte BCD-Repair" -Level INFO
    
    $bcdLocations = @(
        @{Path = "Boot\BCD"; Loader = "\windows\system32\boot\winload.exe"; Type = "BIOS" },
        @{Path = "EFI\Microsoft\Boot\BCD"; Loader = "\windows\system32\boot\winload.efi"; Type = "UEFI" }
    )
    $ramdiskGuid = "{7619dcc8-fafe-11d9-b411-000476eba25f}"
    
    foreach ($loc in $bcdLocations) {
        $bcdPath = Join-Path $MediaPath $loc.Path
        
        Write-Host " [$($loc.Type)] Konfiguriere BCD..." -ForegroundColor Cyan
        Write-Log "BCD-Repair für $($loc.Type): $bcdPath" -Level INFO
        
        if (-not (Test-Path $bcdPath)) {
            Write-Log "BCD nicht gefunden: $($loc.Path), versuche Template" -Level WARNING
            
            # Versuche BCDTemplate zu kopieren
            $templatePath = $bcdPath -replace "\\BCD$", "\BCDTemplate"
            if (Test-Path $templatePath) {
                Copy-Item $templatePath $bcdPath -Force
                Write-Log "BCD aus Template erstellt: $($loc.Type)" -Level INFO
            }
            else {
                # Erstelle neue BCD
                $bcdDir = Split-Path $bcdPath
                if (-not (Test-Path $bcdDir)) { New-Item $bcdDir -ItemType Directory -Force | Out-Null }
                
                & bcdedit.exe /createstore $bcdPath | Out-Null
                & bcdedit.exe /store $bcdPath /create "{bootmgr}" /d "Windows Boot Manager" | Out-Null
                & bcdedit.exe /store $bcdPath /create "{default}" /d "OPSI WinPE" /application osloader | Out-Null
                Write-Log "Neue BCD erstellt: $($loc.Type)" -Level INFO
            }
        }
        
        # Setze BCD-Einträge
        if (Test-Path $bcdPath) {
            Set-ItemProperty $bcdPath -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue
        }
        
        try {
            # Boot Manager konfigurieren
            & bcdedit.exe /store $bcdPath /set "{bootmgr}" device boot | Out-Null
            & bcdedit.exe /store $bcdPath /set "{bootmgr}" displayorder "{default}" | Out-Null
            & bcdedit.exe /store $bcdPath /set "{bootmgr}" timeout 10 | Out-Null
            & bcdedit.exe /store $bcdPath /set "{bootmgr}" locale de-DE | Out-Null
            
            # OS Loader konfigurieren
            & bcdedit.exe /store $bcdPath /set "{default}" device "ramdisk=[boot]\sources\boot.wim,$ramdiskGuid" | Out-Null
            & bcdedit.exe /store $bcdPath /set "{default}" osdevice "ramdisk=[boot]\sources\boot.wim,$ramdiskGuid" | Out-Null
            & bcdedit.exe /store $bcdPath /set "{default}" path "$($loc.Loader)" | Out-Null
            & bcdedit.exe /store $bcdPath /set "{default}" systemroot "\windows" | Out-Null
            & bcdedit.exe /store $bcdPath /set "{default}" locale de-DE | Out-Null
            & bcdedit.exe /store $bcdPath /set "{default}" winpe Yes | Out-Null
            & bcdedit.exe /store $bcdPath /set "{default}" detecthal Yes | Out-Null
            & bcdedit.exe /store $bcdPath /set "{default}" nointegritychecks Yes | Out-Null
            
            # Ramdisk konfigurieren
            $ramdiskExists = & bcdedit.exe /store $bcdPath /enum all 2>&1 | Select-String $ramdiskGuid
            if (-not $ramdiskExists) {
                & bcdedit.exe /store $bcdPath /create $ramdiskGuid /d "Ramdisk Options" /device | Out-Null
            }
            & bcdedit.exe /store $bcdPath /set $ramdiskGuid ramdisksdidevice boot | Out-Null
            & bcdedit.exe /store $bcdPath /set $ramdiskGuid ramdisksdipath "\Boot\boot.sdi" | Out-Null
            
            Write-Host " [OK] $($loc.Type) BCD repariert" -ForegroundColor Green
            Write-Log "$($loc.Type) BCD erfolgreich konfiguriert" -Level SUCCESS
        }
        catch {
            Write-Log ("Fehler bei " + $loc.Type + " BCD-Konfiguration: " + $_) -Level ERROR
            Write-Warning "BCD-Konfiguration für $($loc.Type) fehlgeschlagen"
        }
    }
    
    Write-Log "BCD-Repair abgeschlossen" -Level SUCCESS
}

# --- FUNCTION: Export-OPSIStructure ---
function Export-OPSIStructure {
    param(
        [string]$SourcePath,
        [string]$OpsiPath
    )
    
    Write-Host " [OPSI] Exportiere OPSI-Depot-Struktur..." -ForegroundColor Yellow
    Write-Log "Starte OPSI-Struktur Export nach: $OpsiPath" -Level INFO
    
    if (-not (Test-Path $OpsiPath)) {
        New-Item $OpsiPath -ItemType Directory -Force | Out-Null
        Write-Log "OPSI-Export-Verzeichnis erstellt: $OpsiPath"
    }
    
    foreach ($dirName in @("winpe", "winpe_uefi")) {
        $target = Join-Path $OpsiPath $dirName
        
        Write-Host " [EXPORT] $dirName Struktur..." -ForegroundColor Cyan
        Write-Log "Exportiere nach: $target" -Level INFO
        
        # Bereinige Zielordner
        if (Test-Path $target) {
            Write-Log "Lösche vorhandenen Ordner: $target"
            Remove-Item $target -Recurse -Force -ErrorAction SilentlyContinue
        }
        New-Item $target -ItemType Directory -Force | Out-Null
        
        # Kopiere komplette Struktur
        Write-Host " [COPY] Kopiere Dateien..." -ForegroundColor Gray
        Get-ChildItem -Path "$SourcePath\*" -Force -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                Copy-Item $_.FullName -Destination $target -Recurse -Force -ErrorAction Stop
                Write-Log "Kopiert: $($_.Name)" -Level INFO
            }
            catch {
                Write-Log ("Fehler beim Kopieren von " + $_.Name + ": " + $_) -Level WARNING
            }
        }
        
        # Validierung kritischer Dateien
        $critical = @(
            "bootmgr",
            "bootmgr.efi",
            "Boot\BCD",
            "Boot\boot.sdi",
            "EFI\Microsoft\Boot\BCD",
            "sources\boot.wim"
        )
        
        $missing = @()
        foreach ($file in $critical) {
            $fullPath = Join-Path $target $file
            if (-not (Test-Path $fullPath)) {
                $missing += $file
            }
        }
        
        if ($missing.Count -gt 0) {
            Write-Log "Fehlende Dateien in ${dirName}: $($missing -join ', ')" -Level WARNING
            Write-Warning "Fehlende Dateien in $dirName : $($missing -join ', ')"
        }
        else {
            Write-Host " [OK] $dirName vollständig exportiert" -ForegroundColor Green
            Write-Log "$dirName Export erfolgreich, alle kritischen Dateien vorhanden" -Level SUCCESS
        }
    }
    
    Write-Log "OPSI-Struktur Export abgeschlossen" -Level SUCCESS
}

# ============================================================================
# HAUPTPROGRAMM
# ============================================================================

try {
    # --- 2. PARAMETER-VALIDIERUNG ---
    Write-Host "`n[VALIDATION] Validiere Parameter..." -ForegroundColor Cyan
    Write-Log "Starte Parameter-Validierung"
    
    # Validate Source ISO
    if (-not (Test-Path $sourceiso)) {
        $errMsg = "Source ISO nicht gefunden: $sourceiso"
        Write-Log $errMsg -Level ERROR
        throw $errMsg
    }
    Write-Host " [OK] Source ISO: $sourceiso" -ForegroundColor Green
    Write-Log "Source ISO validiert: $sourceiso"
    
    # Validate Driver Source
    if (-not (Test-Path $DriverSource)) {
        $errMsg = "Driver Source nicht gefunden: $DriverSource"
        Write-Log $errMsg -Level ERROR
        throw $errMsg
    }
    Write-Host " [OK] Driver Source: $DriverSource" -ForegroundColor Green
    Write-Log "Driver Source validiert: $DriverSource"
    
    # Create/Validate Working Path
    if (-not (Test-Path $WorkingPath)) {
        try {
            New-Item -ItemType Directory -Path $WorkingPath -Force | Out-Null
            Write-Host " [OK] WorkingPath erstellt: $WorkingPath" -ForegroundColor Green
            Write-Log "WorkingPath erstellt: $WorkingPath"
        }
        catch {
            $errMsg = "Kann WorkingPath nicht erstellen: $WorkingPath - $_"
            Write-Log $errMsg -Level ERROR
            throw $errMsg
        }
    }
    else {
        Write-Host " [OK] WorkingPath vorhanden: $WorkingPath" -ForegroundColor Green
        Write-Log "WorkingPath validiert: $WorkingPath"
    }
    
    # Check disk space
    $drive = Split-Path -Qualifier $WorkingPath
    if ($drive) {
        $disk = Get-PSDrive -Name $drive.TrimEnd(':')
        $freeSpaceGB = [math]::Round($disk.Free / 1GB, 2)
        if ($disk.Free -lt 10GB) {
            Write-Host " [WARNING] Wenig Speicherplatz auf ${drive}: $freeSpaceGB GB" -ForegroundColor Yellow
            Write-Host " [WARNING] Empfohlen: Mindestens 10 GB" -ForegroundColor Yellow
            Write-Log "Wenig Speicherplatz: $freeSpaceGB GB" -Level WARNING
        }
        else {
            Write-Host " [OK] Freier Speicherplatz auf ${drive}: $freeSpaceGB GB" -ForegroundColor Green
            Write-Log "Freier Speicherplatz: $freeSpaceGB GB"
        }
    }
    
    Write-Host " [VALIDATION] Parameter validiert`n" -ForegroundColor Green
    
    # --- 3. ADK DETECTION ---
    Write-Progress -Activity "WinPE Build V14.0" -Status "Suche ADK..." -PercentComplete 5
    $script:ADKRoot = Get-ADKPath
    Write-Progress -Activity "WinPE Build V14.0" -Status "ADK gefunden" -PercentComplete 10
    
    # --- 4. ROBUSTER CLEANUP ---
    Write-Host "`n[CLEANUP] Bereinige Mount-Punkte und Verzeichnisse..." -ForegroundColor Yellow
    Write-Log "Starte Cleanup"
    Write-Progress -Activity "WinPE Build V14.0" -Status "Cleanup..." -PercentComplete 12
    
    $mountedImages = Get-WindowsImage -Mounted
    foreach ($img in $mountedImages) {
        if ($img.Path -eq $MountPath) {
            Write-Host " [CLEANUP] Aktiver Mount gefunden: $MountPath. Unmount wird erzwungen..." -ForegroundColor Red
            Write-Log "Erzwinge Unmount von: $MountPath" -Level WARNING
            Dism /Unmount-Wim /MountDir:$MountPath /Discard | Out-Null
        }
    }
    
    if (Test-Path $MountPath) {
        Remove-Item $MountPath -Recurse -Force -ErrorAction SilentlyContinue
    }
    
    # Erstelle frische Ordner
    New-Item -ItemType Directory -Path $MountPath, $InspectPath, $opsiexportpath -Force | Out-Null
    Write-Log "Cleanup abgeschlossen, Verzeichnisstruktur erstellt"
    
    # --- 5. ISO EXTRAKTION & CACHING ---
    Write-Progress -Activity "WinPE Build V14.0" -Status "ISO-Extraktion..." -PercentComplete 15
    
    $wimCheckPath = Join-Path $opsiexportpath "sources\boot.wim"
    $needsIsoExtraction = -not $UseCache -or -not (Test-Path $wimCheckPath)
    
    if ($needsIsoExtraction) {
        Write-Host "`n[ISO] Extrahiere Basis-Struktur von ISO..." -ForegroundColor Cyan
        Write-Log "Starte ISO-Extraktion von: $sourceiso"
        Write-Progress -Activity "WinPE Build V14.0" -Status "Mounte ISO..." -PercentComplete 16
        
        $mountResult = Mount-DiskImage -ImagePath $sourceiso -PassThru
        $driveLetter = ($mountResult | Get-Volume).DriveLetter
        if (!$driveLetter) { 
            $errMsg = "ISO konnte nicht gemountet werden"
            Write-Log $errMsg -Level ERROR
            throw $errMsg
        }
        
        $isoPath = "$($driveLetter):\"
        Write-Host " [ISO] Kopiere Dateien von $isoPath nach $opsiexportpath ..." -ForegroundColor Cyan
        Write-Log "Kopiere ISO-Dateien von: $isoPath"
        Write-Progress -Activity "WinPE Build V14.0" -Status "Kopiere ISO-Inhalt..." -PercentComplete 18
        
        Copy-Item "$isoPath*" $opsiexportpath -Recurse -Force
        Dismount-DiskImage -ImagePath $sourceiso | Out-Null
        
        Write-Log "ISO-Extraktion abgeschlossen"
        Write-Progress -Activity "WinPE Build V14.0" -Status "ISO extrahiert" -PercentComplete 25
    }
    else {
        Write-Host "`n[CACHE] Nutze gecachte ISO-Struktur in $opsiexportpath" -ForegroundColor Green
        Write-Log "Nutze gecachte ISO-Struktur"
        Write-Progress -Activity "WinPE Build V14.0" -Status "Cache verwendet" -PercentComplete 25
    }
    
    # --- 6. TOOLKIT VORBEREITUNG ---
    Write-Progress -Activity "WinPE Build V14.0" -Status "Toolkit vorbereiten..." -PercentComplete 28
    
    if (-not $SkipToolkitDownload) {
        if ($ToolsSource -eq "AUTO") {
            Write-Host "`n[TOOLKIT] Auto-Download-Modus aktiviert" -ForegroundColor Cyan
            Get-ToolkitFiles -DestPath $ToolkitPath
        }
        elseif (Test-Path $ToolsSource) {
            Write-Host "`n[TOOLKIT] Verwende lokales Toolkit: $ToolsSource" -ForegroundColor Green
            $ToolkitPath = $ToolsSource
            Write-Log "Lokales Toolkit: $ToolsSource"
        }
        else {
            Write-Warning "ToolsSource nicht gefunden: $ToolsSource (wird als AUTO behandelt)"
            Write-Log "ToolsSource ungültig, verwende AUTO" -Level WARNING
            Get-ToolkitFiles -DestPath $ToolkitPath
        }
    }
    else {
        Write-Host "`n[TOOLKIT] Toolkit-Download übersprungen" -ForegroundColor Gray
        Write-Log "Toolkit-Download übersprungen (SkipToolkitDownload)" -Level INFO
    }
    
    # --- 7. WIM MODIFIKATION ---
    $SourceWim = Join-Path $opsiexportpath "sources\boot.wim"
    Write-Host "`n[MODIFY] Bereite boot.wim zur Bearbeitung vor..." -ForegroundColor Cyan
    Write-Log "Starte WIM Modifikation"
    Write-Progress -Activity "WinPE Build V14.0" -Status "Kopiere boot.wim..." -PercentComplete 32
    
    Copy-Item $SourceWim $BootWimMod -Force
    Set-ItemProperty -Path $BootWimMod -Name IsReadOnly -Value $false
    
    Write-Host " [MODIFY] Mounte WinPE (Index 1)..." -ForegroundColor Cyan
    Write-Log "Mounte boot.wim Index 1"
    Write-Progress -Activity "WinPE Build V14.0" -Status "Mounte WIM..." -PercentComplete 35
    
    try {
        Mount-WindowsImage -ImagePath $BootWimMod -Index 1 -Path $MountPath -ErrorAction Stop | Out-Null
        Write-Log "WIM erfolgreich gemountet"
    }
    catch {
        $errMsg = "Fehler beim Mounten des WIM: $_"
        Write-Log $errMsg -Level ERROR
        throw $errMsg
    }
    
    # --- 8. WINPE KOMPONENTEN INTEGRATION ---
    Write-Progress -Activity "WinPE Build V14.0" -Status "Integriere WinPE-Komponenten..." -PercentComplete 40
    Add-WinPEComponents -MountPath $MountPath -ADKRoot $script:ADKRoot
    
    # --- 9. TREIBER INTEGRATION ---
    Write-Host "`n[DRIVERS] Injektiere Treiber aus $DriverSource..." -ForegroundColor Cyan
    Write-Log "Injektiere Treiber aus: $DriverSource"
    Write-Progress -Activity "WinPE Build V14.0" -Status "Füge Treiber hinzu..." -PercentComplete 48
    
    try {
        Dism /Image:$MountPath /Add-Driver /Driver:$DriverSource /Recurse /ForceUnsigned | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "DISM Add-Driver fehlgeschlagen mit Exit-Code: $LASTEXITCODE" }
        Write-Host " [OK] Treiber erfolgreich injektiert" -ForegroundColor Green
        Write-Log "Treiber erfolgreich injektiert" -Level SUCCESS
    }
    catch {
        $errMsg = "Fehler beim Injektieren der Treiber: $_"
        Write-Log $errMsg -Level ERROR
        Write-Warning $errMsg
    }
    
    # --- 10. TOOLKIT INJECTION ---
    if (-not $SkipToolkitDownload -and (Test-Path $ToolkitPath)) {
        Write-Host "`n[INJECT] Injektiere Toolkit ins WIM..." -ForegroundColor Cyan
        Write-Progress -Activity "WinPE Build V14.0" -Status "Injektiere Toolkit..." -PercentComplete 55
        
        $WimToolkitPath = Join-Path $MountPath "Program Files\Toolkit"
        if (-not (Test-Path $WimToolkitPath)) { 
            New-Item $WimToolkitPath -ItemType Directory -Force | Out-Null 
        }
        
        Get-ChildItem $ToolkitPath -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                Copy-Item $_.FullName -Destination $WimToolkitPath -Recurse -Force
                Write-Log "Toolkit injiziert: $($_.Name)" -Level SUCCESS
            }
            catch {
                Write-Log ("Fehler beim Injektieren von " + $_.Name + ": " + $_) -Level WARNING
            }
        }
        
        Write-Host " [OK] Toolkit ins WIM injektiert" -ForegroundColor Green
    }
    
    # --- 11. LAUNCHER MENÜ & LOKALISIERUNG ---
    Write-Progress -Activity "WinPE Build V14.0" -Status "Erstelle Menü & Lokalisierung..." -PercentComplete 62
    New-LauncherMenu -MountPath $MountPath
    Set-GermanLocale -MountPath $MountPath
    
    # --- 12. INSPECT-KOPIE ---
    Write-Host "`n[INSPECT] Erstelle einsehbare Kopie des WIM-Inhalts..." -ForegroundColor Yellow
    Write-Log "Erstelle Inspect-Kopie nach: $InspectPath"
    Write-Progress -Activity "WinPE Build V14.0" -Status "Erstelle Inspect-Kopie..." -PercentComplete 68
    
    robocopy $MountPath $InspectPath /E /R:0 /W:0 /XJ /NFL /NDL /NJH /NJS | Out-Null
    Write-Log "Inspect-Kopie erstellt"
    
    # --- 13. SPEICHERN & UNMOUNT ---
    if ($KeepMountOpen) {
        Write-Host "`n[DEBUG] Mount bleibt offen: $MountPath" -ForegroundColor Magenta
        Write-Log "Mount bleibt offen für Debugging (KeepMountOpen)" -Level INFO
        Write-Host "Drücke Enter wenn bereit zum Fortfahren..." -ForegroundColor Yellow
        Read-Host
    }
    
    Write-Host "`n[SAVE] Schließe WIM (Commit)..." -ForegroundColor Cyan
    Write-Log "Speichere und unmounte WIM"
    Write-Progress -Activity "WinPE Build V14.0" -Status "Speichere WIM..." -PercentComplete 75
    
    try {
        Dismount-WindowsImage -Path $MountPath -Save -ErrorAction Stop | Out-Null
        Write-Log "WIM erfolgreich gespeichert"
    }
    catch {
        $errMsg = "Fehler beim Unmounten des WIM: $_"
        Write-Log $errMsg -Level ERROR
        
        Write-Host " [ERROR] Versuche Discard als Fallback..." -ForegroundColor Red
        try {
            Dismount-WindowsImage -Path $MountPath -Discard -ErrorAction Stop | Out-Null
        }
        catch {
            Write-Log "Auch Discard fehlgeschlagen: $_" -Level ERROR
        }
        throw $errMsg
    }
    
    # Kopiere modifizierte WIM zurück
    Write-Progress -Activity "WinPE Build V14.0" -Status "Kopiere modifizierte WIM..." -PercentComplete 80
    Copy-Item $BootWimMod $SourceWim -Force
    Write-Log "Modifizierte WIM zurückkopiert"
    
    # --- 14. BCD REPAIR ---
    Write-Progress -Activity "WinPE Build V14.0" -Status "Repariere BCD..." -PercentComplete 85
    Repair-BCDConfiguration -MediaPath $opsiexportpath
    
    # --- 15. OPSI STRUKTUR EXPORT ---
    Write-Progress -Activity "WinPE Build V14.0" -Status "Exportiere OPSI-Struktur..." -PercentComplete 92
    Export-OPSIStructure -SourcePath $opsiexportpath -OpsiPath $opsiexportpath
    
    # --- 16. ISO GENERIERUNG (Optional) ---
    Write-Host "`n[ISO] Generiere bootfähiges ISO-Image..." -ForegroundColor Cyan
    Write-Log "Starte ISO-Generierung"
    Write-Progress -Activity "WinPE Build V14.0" -Status "Generiere ISO..." -PercentComplete 96
    
    $oscdimgPath = Get-ChildItem $script:ADKRoot -Filter "oscdimg.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    
    if ($oscdimgPath) {
        $etfs = Join-Path $opsiexportpath "boot\etfsboot.com"
        $efi = Join-Path $opsiexportpath "efi\microsoft\boot\efisys.bin"
        
        if (Test-Path $etfs) {
            try {
                # Bootfähiges ISO mit legacy + UEFI Support
                if (Test-Path $efi) {
                    $bootData = "2#p0,e,b`"$etfs`"#pEF,e,b`"$efi`""
                    $oscdimgArgs = "-n -m -o -u2 -udfver102 -bootdata:$bootData `"$opsiexportpath`" `"$IsoFile`""
                }
                else {
                    $oscdimgArgs = "-n -m -b`"$etfs`" `"$opsiexportpath`" `"$IsoFile`""
                }
                
                Write-Log "Verwende oscdimg.exe: $($oscdimgPath.FullName)"
                Start-Process -FilePath $oscdimgPath.FullName -ArgumentList $oscdimgArgs -Wait -NoNewWindow
                
                Write-Host " [OK] ISO erfolgreich erstellt: $IsoFile" -ForegroundColor Green
                Write-Log "ISO erfolgreich erstellt: $IsoFile" -Level SUCCESS
            }
            catch {
                $errMsg = "Fehler bei ISO-Generierung: $_"
                Write-Log $errMsg -Level ERROR
                Write-Warning $errMsg
            }
        }
        else {
            Write-Warning "etfsboot.com nicht gefunden, ISO-Generierung übersprungen"
            Write-Log "etfsboot.com nicht gefunden" -Level WARNING
        }
    }
    else {
        Write-Warning "oscdimg.exe nicht gefunden, ISO-Generierung übersprungen"
        Write-Log "oscdimg.exe nicht gefunden" -Level WARNING
    }
    
    # --- FERTIG ---
    Write-Progress -Activity "WinPE Build V14.0" -Status "Abgeschlossen" -PercentComplete 100 -Completed
    
    Write-Host "`n===============================================" -ForegroundColor Cyan
    Write-Host "   ERFOLGREICH ABGESCHLOSSEN                  " -ForegroundColor Green
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "`n[DEBUG] Einsehbare Kopie: $InspectPath" -ForegroundColor Gray
    Write-Host "[DEBUG] Log-Datei: $LogFile" -ForegroundColor Gray
    Write-Host "[DEBUG] OPSI Export: $opsiexportpath" -ForegroundColor Gray
    if (Test-Path $IsoFile) {
        Write-Host "[DEBUG] ISO-Datei: $IsoFile" -ForegroundColor Gray
    }
    
    Write-Log "WinPE Build V14.0 erfolgreich abgeschlossen" -Level SUCCESS
    Write-Host "`n[ERFOLG] WinPE V14.0 OPSI Edition Build abgeschlossen!" -ForegroundColor Green
    Write-Host "`nHinweise:" -ForegroundColor Yellow
    Write-Host " - OPSI-Strukturen: winpe/ und winpe/_uefi/" -ForegroundColor Yellow
    Write-Host " - Secure Boot muss ggf. deaktiviert werden" -ForegroundColor Yellow
    Write-Host " - opsi-set-rights auf dem OPSI-Server ausführen!" -ForegroundColor Yellow
}
catch {
    Write-Host "`n!!! KRITISCHER FEHLER !!!" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Log "Kritischer Fehler: $($_.Exception.Message)" -Level ERROR
    
    # Cleanup bei Fehler
    if (Test-Path $MountPath) {
        $mounted = Get-WindowsImage -Mounted | Where-Object { $_.Path -eq $MountPath }
        if ($mounted) {
            Write-Host "Versuche Mount zu bereinigen..." -ForegroundColor Yellow
            try {
                Dismount-WindowsImage -Path $MountPath -Discard -ErrorAction Stop | Out-Null
                Write-Log "Mount nach Fehler bereinigt"
            }
            catch {
                Write-Log "Fehler beim Bereinigen des Mounts: $_" -Level ERROR
            }
        }
    }
    
    exit 1
}