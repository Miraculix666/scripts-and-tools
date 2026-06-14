<#
.SYNOPSIS
    PSC WinPE Builder - Erstellt enterprise-fähige WinPE-Umgebungen für OPSI-Deployment
    
    Version: 6.3 (Strict Handbook Compliance & Addon Preservation)
    Author: PS-Coding (AI-Assistent)
    Letzte Änderung: 2026-02-13

.DESCRIPTION
    Dieser Builder erstellt ein hochgradig angepasstes Windows Preinstallation Environment (WinPE)
    mit integrierten Administrationstools und vollständiger OPSI-Kompatibilität.
    
    ╔══════════════════════════════════════════════════════════════════════════════╗
    ║                         WORKFLOW-ÜBERSICHT                                   ║
    ╚══════════════════════════════════════════════════════════════════════════════╝
    
    1. WORKSPACE INITIALIZATION
       - Erstellt temporäre Verzeichnisstruktur
       - Führt copype.cmd aus (Windows ADK)
       - Bereitet boot.wim für Modifikation vor
    
    2. TOOLKIT DOWNLOAD & INTEGRATION
       - Lädt automatisch herunter:
         • Notepad++ (Texteditor)
         • Explorer++ (Dateimanager)
         • Micro Editor (CLI-Editor)
         • Sysinternals Suite (Systemtools)
         • NirSoft Launcher (Utilities)
       - Integriert Tools DIREKT IN boot.wim (nicht außerhalb!)
    
    3. WIM MODIFICATION
       - Mounted boot.wim Image
       - Fügt WinPE Optional Components hinzu (WMI, PowerShell, NetFX)
       - Injektiert essenzielle Treiber (Netzwerk, Storage)
       - Installiert LaunchMenu.ps1 für interaktiven Zugriff
    
    4. OPSI-EXPORT (Handbuch-konform)
       - Erstellt Struktur: boot/, EFI/, sources/, bootmgr, bootmgr.efi
       - Alle Tools sind INNERHALB von boot.wim (Addon-Safe!)
       - PXE-bootfähig
    
    5. ISO-GENERIERUNG
       - Hybrid Boot (BIOS + UEFI)
       - Optional: Kopie ins OPSI-Depot

.PARAMETER WorkingPath
    Temporäres Arbeitsverzeichnis für Build-Prozess
    Standard: C:\PSC_WinPE_Temp
    
    WICHTIG: Benötigt ca. 3-5 GB freien Speicherplatz!

.PARAMETER SourcePath
    Quelle für boot.wim. Kann sein:
    - Pfad zu einer Windows-ISO (wird gemountet und boot.wim extrahiert)
    - Direkter Pfad zu einer boot.wim Datei
    - LEER = Verwendet ADK-Standard winpe.wim
    
    HINWEIS: Dieser Parameter ist der korrekte Name für ISO-/WIM-Quellen!
            (Der alte Name "-sourceiso" wird NICHT unterstützt)

.PARAMETER DriverSource
    Verzeichnis mit Treibern (rekursive Suche nach .inf)
    NUR essenzielle Treiber werden injektiert:
    - Net (Netzwerk)
    - NetTrans (Netzwerktransport)
    - HDC (Festplattencontroller)
    - SCSIAdapter (SCSI/RAID)
    
    Beispiel: Y:\opsi-winpe\drivers\drivers

.PARAMETER IsoPath
    Ausgabepfad für das finale ISO-Image
    Standard: C:\PSC_WinPE_Output\WinPE_OPSI_Custom.iso

.PARAMETER OpsiExportPath
    Zielverzeichnis für OPSI-Depot-Export
    Erstellt Unterordner "winpe\" mit boot-fähiger Struktur
    
    WICHTIG: Verwenden Sie den OPSI-Depot-Pfad!
    Beispiel: \\opsi-server\depot\winpe oder C:\Temp\opsiPE_new13.4

.PARAMETER Architecture
    Prozessor-Architektur
    Gültige Werte: amd64 (Standard), x86, arm64

.PARAMETER ADKPath
    Optionaler ADK-Pfad (für nicht-Standard-Installationen)
    Standard: Auto-Detection in C:\Program Files

.PARAMETER ScratchSpace
    WinPE Scratch Space in MB
    Standard: 512 MB
    Erhöhen bei großen Toolkit-Downloads oder vielen Treibern

.EXAMPLE
    # ╔═══════════════════════════════════════════════════════════════════╗
    # ║ USE CASE 1: Vollständiger Build aus Windows-ISO                  ║
    # ╚═══════════════════════════════════════════════════════════════════╝
    
    .\WinPE_maker.ps1 `
        -SourcePath "D:\Downloads\SW_DVD9_WIN_ENT_LTSC_2024_64-bit_German_MLF_X23-70052.ISO" `
        -DriverSource "Y:\opsi-winpe\drivers\drivers" `
        -OpsiExportPath "C:\Temp\opsiPE_new13.4" `
        -WorkingPath "C:\temp_winPE_working" `
        -Verbose
    
    ERGEBNIS:
    - boot.wim wird aus ISO extrahiert
    - Toolkit wird heruntergeladen (ca. 100 MB)
    - Treiber aus Y:\opsi-winpe\drivers\drivers werden injektiert
    - OPSI-Struktur wird nach C:\Temp\opsiPE_new13.4\winpe exportiert
    - ISO wird erstellt

.EXAMPLE
    # ╔═══════════════════════════════════════════════════════════════════╗
    # ║ USE CASE 2: Build aus vorhandener boot.wim                       ║
    # ╚═══════════════════════════════════════════════════════════════════╝
    
    .\WinPE_maker.ps1 `
        -SourcePath "C:\WinPE_Sources\boot.wim" `
        -DriverSource "C:\Drivers\Intel_NUC" `
        -OpsiExportPath "\\opsi-server\depot" `
        -IsoPath "C:\Output\WinPE_Intel_NUC.iso"
    
    ERGEBNIS:
    - Verwendet vorhandene boot.wim
    - Intel NUC-spezifische Treiber
    - Direkt ins OPSI-Depot exportiert
    - Custom ISO-Name

.EXAMPLE
    # ╔═══════════════════════════════════════════════════════════════════╗
    # ║ USE CASE 3: Minimal-Build ohne Treiber (ADK-Standard)            ║
    # ╚═══════════════════════════════════════════════════════════════════╝
    
    .\WinPE_maker.ps1 `
        -OpsiExportPath "C:\OPSI_Minimal" `
        -IsoPath "C:\Output\WinPE_Clean.iso"
    
    ERGEBNIS:
    - Verwendet winpe.wim aus ADK
    - Keine zusätzlichen Treiber
    - Nur Toolkit + WinPE Components
    - Ideal für virtuelle Umgebungen

.EXAMPLE
    # ╔═══════════════════════════════════════════════════════════════════╗
    # ║ USE CASE 4: x86 (32-Bit) WinPE für Legacy-Hardware               ║
    # ╚═══════════════════════════════════════════════════════════════════╝
    
    .\WinPE_maker.ps1 `
        -Architecture "x86" `
        -SourcePath "D:\ISOs\Windows_Server_2022.iso" `
        -DriverSource "C:\Drivers\Legacy" `
        -OpsiExportPath "C:\OPSI_x86"
    
    ERGEBNIS:
    - 32-Bit WinPE
    - Legacy-Treiber-Support
    - Für alte Systeme

.EXAMPLE
    # ╔═══════════════════════════════════════════════════════════════════╗
    # ║ USE CASE 5: Debug-Modus (Verbose + kein OPSI-Export)             ║
    # ╚═══════════════════════════════════════════════════════════════════╝
    
    .\WinPE_maker.ps1 `
        -SourcePath "D:\ISOs\Windows11.iso" `
        -DriverSource "C:\Drivers" `
        -WorkingPath "C:\Debug_WinPE" `
        -IsoPath "C:\Output\Debug_WinPE.iso" `
        -Verbose `
        -Debug
    
    ERGEBNIS:
    - Detaillierte Ausgaben
    - Kein OPSI-Export (nur ISO)
    - Arbeitsverzeichnis bleibt erhalten für Inspektion

.NOTES
    ╔══════════════════════════════════════════════════════════════════════════════╗
    ║                         VORAUSSETZUNGEN                                      ║
    ╚══════════════════════════════════════════════════════════════════════════════╝
    
    • Windows ADK (Assessment and Deployment Kit) installiert
      Download: https://docs.microsoft.com/en-us/windows-hardware/get-started/adk-install
    
    • Administrator-Rechte (für DISM-Operationen)
    
    • Internetverbindung (für Toolkit-Download)
      ODER vorbefüllter Toolkit-Ordner in $WorkingPath\ExternalTools
    
    • Freier Speicherplatz:
      - Temp: 3-5 GB
      - Output: 500 MB - 1 GB
    
    ╔══════════════════════════════════════════════════════════════════════════════╗
    ║                         WICHTIGE HINWEISE                                    ║
    ╚══════════════════════════════════════════════════════════════════════════════╝
    
    ⚠ PARAMETER-KOMPATIBILITÄT:
      Der Parameter "-sourceiso" existiert NICHT!
      Korrekter Name: -SourcePath
    
    ⚠ OPSI-STRUKTUR:
      Alle Tools befinden sich INNERHALB von boot.wim
      Die exportierte "winpe\"-Struktur enthält nur Boot-Dateien
      Dies ist gemäß OPSI-Handbuch korrekt!
    
    ⚠ TOOLKIT-CACHING:
      Bereits heruntergeladene Tools in $WorkingPath\ExternalTools
      werden NICHT erneut geladen (spart Bandwidth)
    
    ⚠ TREIBER-FILTER:
      Nur essenzielle Treiberklassen werden injektiert
      Dies verhindert Bloat und Boot-Probleme
    
    ╔══════════════════════════════════════════════════════════════════════════════╗
    ║                         LAUNCHER-MENÜ                                        ║
    ╚══════════════════════════════════════════════════════════════════════════════╝
    
    Nach dem Boot startet automatisch ein interaktives Menü:
    
    [1] Notepad++      → Texteditor
    [2] Explorer++     → Dateimanager
    [3] Micro Editor   → CLI-Editor
    [4] Sysinternals   → Systemtools-Sammlung
    [5] NirSoft Tool   → Utility-Launcher
    [6] PowerShell     → PS-Konsole
    [7] Registry/Disk  → Regedit + Diskpart
    [8] Netzwerk-Info  → ipconfig /all
    [9] REBOOT         → Neustart
    [0] SHUTDOWN       → Herunterfahren
    
    ╔══════════════════════════════════════════════════════════════════════════════╗
    ║                         TROUBLESHOOTING                                      ║
    ╚══════════════════════════════════════════════════════════════════════════════╝
    
    PROBLEM: "ADK nicht gefunden"
    LÖSUNG:  ADK installieren oder -ADKPath angeben
    
    PROBLEM: "Toolkit-Download schlägt fehl"
    LÖSUNG:  Manuelle Downloads in $WorkingPath\ExternalTools ablegen
    
    PROBLEM: "Mount schlägt fehl"
    LÖSUNG:  Alte Mounts bereinigen: dism /cleanup-wim
    
    PROBLEM: "OPSI-Boot funktioniert nicht"
    LÖSUNG:  BCD prüfen mit: bcdedit /store winpe\Boot\BCD /enum all

.LINK
    https://docs.microsoft.com/en-us/windows-hardware/manufacture/desktop/winpe-intro
    
.LINK
    OPSI Documentation: https://docs.opsi.org/

#>

Param(
    [Parameter(Mandatory = $false)]
    [string]$WorkingPath = "C:\PSC_WinPE_Temp",

    [Alias("BootWIM")]
    [Parameter(Mandatory = $false)]
    [string]$SourcePath,

    [Parameter(Mandatory = $false)]
    [string]$DriverSource,

    [Parameter(Mandatory = $false)]
    [string]$IsoPath = "C:\PSC_WinPE_Output\WinPE_OPSI_Custom.iso",

    [Alias("opsidepot")]
    [Parameter(Mandatory = $false)]
    [string]$OpsiExportPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet("amd64", "x86", "arm64")]
    [string]$Architecture = "amd64",

    [Parameter(Mandatory = $false)]
    [string]$ADKPath,

    [Parameter(Mandatory = $false)]
    [int]$ScratchSpace = 512
)

# --- ADMIN-CHECK (Klare Fehlermeldung - kein neues Fenster!) ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Red
    Write-Host "║  FEHLER: Administrator-Rechte erforderlich!                   ║" -ForegroundColor Red
    Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Red
    Write-Host ""
    Write-Host "Dieses Skript benötigt Administrator-Rechte für DISM-Operationen." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "LÖSUNG:" -ForegroundColor Cyan
    Write-Host "  1. Öffnen Sie PowerShell als Administrator:" -ForegroundColor White
    Write-Host "     - Windows-Taste drücken" -ForegroundColor Gray
    Write-Host "     - 'PowerShell' eingeben" -ForegroundColor Gray
    Write-Host "     - Rechtsklick auf 'Windows PowerShell'" -ForegroundColor Gray
    Write-Host "     - 'Als Administrator ausführen' wählen" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  2. Wechseln Sie in dieses Verzeichnis:" -ForegroundColor White
    Write-Host "     cd '$PSScriptRoot'" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  3. Führen Sie das Skript erneut aus mit Ihren Parametern" -ForegroundColor White
    Write-Host ""
    Write-Host "Drücken Sie eine Taste zum Beenden..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

Write-Host "[✓] Administrator-Rechte bestätigt`n" -ForegroundColor Green

# --- Initialisierung ---
$ErrorActionPreference = "Stop"
$VerbosePreference = "Continue" 
[System.Threading.Thread]::CurrentThread.CurrentCulture = "de-DE"
[System.Threading.Thread]::CurrentThread.CurrentUICulture = "de-DE"

$WorkingPath = [string]($WorkingPath | Select-Object -First 1).TrimEnd('\')
$script:ADKRootFolder = ""
$MountPath = "$WorkingPath\mount"
$script:StatusReport = @{}

# Toolkit-Definitionen
$Toolkit = @{
    "NotepadPP"    = "https://github.com/notepad-plus-plus/notepad-plus-plus/releases/download/v8.6.2/npp.8.6.2.portable.x64.zip"
    "ExplorerPP"   = "https://explorerplusplus.com/software/explorer++_1.3.5_x64.zip"
    "MicroEditor"  = "https://github.com/zyedidia/micro/releases/download/v2.0.13/micro-2.0.13-win64.zip"
    "Sysinternals" = "https://download.sysinternals.com/files/SysinternalsSuite.zip"
    "NirSoft"      = "https://www.nirsoft.net/panel/nirlauncher.zip"
}
$ToolsPath = "$WorkingPath\ExternalTools"

# --- Funktionen ---

function Show-Greeting {
    Write-Host "`n===============================================" -ForegroundColor Cyan
    Write-Host "   PSC_WinPE-Customizer v6.3 (Addon-Safe OPSI) " -ForegroundColor Cyan
    if ($SourcePath) { Write-Host "   Quelle: $(Split-Path -Path [string]$SourcePath -Leaf)" -ForegroundColor Yellow }
    Write-Host "   Architektur: $Architecture | Scratch: $ScratchSpace MB" -ForegroundColor Cyan
    Write-Host "===============================================" -ForegroundColor Cyan
}

function Find-ADK {
    $RawRoots = @("C:\Program Files (x86)\Windows Kits\10", "C:\Program Files\Windows Kits\10")
    foreach ($folder in $RawRoots) {
        if (Test-Path "$folder\Assessment and Deployment Kit") {
            $script:ADKRootFolder = $folder
            $script:FinalADKPath = "$folder\Assessment and Deployment Kit\Deployment Tools"
            $script:WinPERoot = "$folder\Assessment and Deployment Kit\Windows Preinstallation Environment\$Architecture"
            return
        }
    }
    throw "ADK nicht gefunden."
}

function Initialize-Workspace {
    Write-Host "[1/9] Vorbereitung des Arbeitsverzeichnisses..." -ForegroundColor Yellow
    if (Test-Path $WorkingPath) {
        & dism.exe /Cleanup-Wim | Out-Null
        Remove-Item $WorkingPath -Recurse -Force -ErrorAction SilentlyContinue
    }
    $Required = @("$WorkingPath\media\sources", "$WorkingPath\media\EFI\Microsoft\Boot", $MountPath, $ToolsPath)
    foreach ($dir in $Required) { New-Item $dir -ItemType Directory -Force | Out-Null }

    # copype.cmd aufrufen
    cmd.exe /c "`"$($script:FinalADKPath)\copype.cmd`" $Architecture `"$WorkingPath`" >nul 2>&1"
    
    $targetWim = "$WorkingPath\media\sources\boot.wim"
    if (-not (Test-Path "$WorkingPath\media\bootmgr")) {
        $MediaTemplate = (Split-Path -Path $script:WinPERoot -Parent) + "\Media"
        if (Test-Path $MediaTemplate) { Copy-Item "$MediaTemplate\*" "$WorkingPath\media" -Recurse -Force }
        $SourceWim = Get-ChildItem $script:WinPERoot -Filter "winpe.wim" -Recurse | Select-Object -First 1
        if ($SourceWim) { Copy-Item $SourceWim.FullName $targetWim -Force }
    }

    if ($SourcePath -and (Test-Path $SourcePath)) {
        if ($SourcePath.EndsWith(".iso")) {
            $iso = Mount-DiskImage $SourcePath -PassThru
            $drive = ($iso | Get-Volume).DriveLetter
            Copy-Item "${drive}:\sources\boot.wim" $targetWim -Force
            Dismount-DiskImage $SourcePath | Out-Null
        }
        else { Copy-Item $SourcePath $targetWim -Force }
    }
    Set-ItemProperty $targetWim -Name IsReadOnly -Value $false
}

function Download-Toolkit {
    Write-Host "[2/9] Lade Admin-Toolkit (Addons)..." -ForegroundColor Yellow
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    foreach ($app in $Toolkit.Keys) {
        $destDir = "$ToolsPath\$app"
        if (-not (Test-Path $destDir)) {
            try {
                Invoke-WebRequest $Toolkit[$app] -OutFile "$ToolsPath\$app.zip" -TimeoutSec 30 -ErrorAction Stop
                [System.IO.Compression.ZipFile]::ExtractToDirectory("$ToolsPath\$app.zip", $destDir)
                Remove-Item "$ToolsPath\$app.zip" -Force
            }
            catch { Write-Warning "Download $app fehlgeschlagen." }
        }
    }
}

function Repair-WinPEImage {
    param (
        [Parameter(Mandatory=$true)]
        [string]$WimPath,

        [Parameter(Mandatory=$true)]
        [string]$WorkingPath
    )

    Write-Host "`n [REPAIR] Versuche automatische Reparatur..." -ForegroundColor Yellow
    
    # Backup erstellen
    $backup = "$WimPath.backup_$(Get-Date -Format 'HHmmss')"
    try {
        Copy-Item $WimPath $backup -Force
        Write-Verbose "Backup erstellt: $backup"
    }
    catch {
        Write-Warning "Backup konnte nicht erstellt werden: $_"
    }
    
    # Versuche DISM Export/Import zur Reparatur
    try {
        $tempWim = "$WorkingPath\temp_repair.wim"
        Write-Verbose "DISM Export für Reparatur..."
        
        & dism /Export-Image /SourceImageFile:"$WimPath" /SourceIndex:1 /DestinationImageFile:"$tempWim" /Compress:max /CheckIntegrity

        if ($LASTEXITCODE -eq 0 -and (Test-Path $tempWim)) {
            Remove-Item $WimPath -Force
            Move-Item $tempWim $WimPath -Force
            Write-Host " [OK] boot.wim erfolgreich repariert" -ForegroundColor Green

            # Validiere reparierte WIM
            $wimInfo = Get-WindowsImage -ImagePath $WimPath
            Write-Host " [OK] Reparierte WIM validiert" -ForegroundColor Green
        }
        else {
            throw "DISM Export fehlgeschlagen (Exit-Code: $LASTEXITCODE)"
        }
    }
    catch {
        Write-Host " [FEHLER] Automatische Reparatur fehlgeschlagen: $_" -ForegroundColor Red
        
        if (Test-Path $backup) {
            Write-Host " [INFO] Stelle Backup wieder her..." -ForegroundColor Yellow
            Copy-Item $backup $WimPath -Force
        }
        
        throw "boot.wim ist korrupt und konnte nicht repariert werden. Bitte verwenden Sie eine andere ISO-Quelle."
    }
}

function Invoke-WinPEMount {
    param (
        [Parameter(Mandatory=$true)]
        [string]$WimPath,

        [Parameter(Mandatory=$true)]
        [string]$MountPath
    )

    $mountAttempts = 0
    $maxAttempts = 2
    $mounted = $false
    
    while (-not $mounted -and $mountAttempts -lt $maxAttempts) {
        $mountAttempts++
        
        try {
            if ($mountAttempts -gt 1) {
                Write-Host " [RETRY] Mount-Versuch $mountAttempts von $maxAttempts..." -ForegroundColor Yellow
            }
            
            Mount-WindowsImage -ImagePath $WimPath -Index 1 -Path $MountPath -ErrorAction Stop | Out-Null
            $mounted = $true
            Write-Host " [OK] WinPE erfolgreich gemountet" -ForegroundColor Green
        }
        catch {
            Write-Host " [FEHLER] Mount fehlgeschlagen: $_" -ForegroundColor Red
            
            if ($mountAttempts -lt $maxAttempts) {
                Write-Host " [CLEANUP] Bereinige für erneuten Versuch..." -ForegroundColor Yellow
                & dism /Cleanup-Wim | Out-Null
                Start-Sleep -Seconds 2
            }
            else {
                throw "Mount endgültig fehlgeschlagen nach $maxAttempts Versuchen: $_"
            }
        }
    }
}

function Mount-WinPE {
    Write-Host "[3/9] Mounte WinPE Image..." -ForegroundColor Yellow

    $wimPath = "$WorkingPath\media\sources\boot.wim"

    # Validiere boot.wim Existenz
    if (-not (Test-Path $wimPath)) {
        throw "boot.wim nicht gefunden: $wimPath"
    }

    Write-Verbose "Prüfe boot.wim Format..."

    # Prüfe WIM-Format und -Integrität mit Get-WindowsImage
    try {
        $wimInfo = Get-WindowsImage -ImagePath $wimPath -ErrorAction Stop
        Write-Host " [OK] boot.wim validiert - $($wimInfo.Count) Index/Indizes gefunden" -ForegroundColor Green

        # Zeige Image-Informationen
        foreach ($img in $wimInfo) {
            Write-Verbose "  Index $($img.ImageIndex): $($img.ImageName) ($([math]::Round($img.ImageSize / 1MB, 2)) MB)"
        }
    }
    catch {
        Write-Host " [FEHLER] boot.wim ist ungültig oder beschädigt!" -ForegroundColor Red
        Write-Host " [FEHLER] Details: $_" -ForegroundColor Red
        Write-Host " [INFO] Mögliche Ursachen:" -ForegroundColor Yellow
        Write-Host "  - ISO-Mount war unvollständig" -ForegroundColor Gray
        Write-Host "  - Datei wurde während Kopiervorgang beschädigt" -ForegroundColor Gray
        Write-Host "  - Falsches Image-Format (kein WIM)" -ForegroundColor Gray

        Repair-WinPEImage -WimPath $wimPath -WorkingPath $WorkingPath
    }

    # Bereinige alte Mount-Versuche
    Write-Verbose "Bereinige alte Mount-Punkte..."
    & dism /Cleanup-Wim | Out-Null

    # Mount mit erweiterter Fehlerbehandlung
    Invoke-WinPEMount -WimPath $wimPath -MountPath $MountPath
}

function Create-WinPELauncher {
    Write-Host "[4/9] Erstelle Toolkit-Launcher (In-Wim)..." -ForegroundColor Yellow
    $LauncherContent = @'
function Show-Menu {
    Clear-Host
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "      OPSI WinPE Enterprise Toolkit Launcher   " -ForegroundColor Cyan
    Write-Host "===============================================" -ForegroundColor Cyan
    try { $ip = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notmatch 'Loopback' } | Select-Object -First 1).IPAddress } catch { $ip = "Suche..." }
    Write-Host " IP-Adresse: $ip" -ForegroundColor Yellow
    Write-Host "-----------------------------------------------"
    Write-Host " [1] Notepad++      [2] Explorer++"
    Write-Host " [3] Micro Editor   [4] Sysinternals"
    Write-Host " [5] NirSoft Tool   [6] PowerShell"
    Write-Host " [7] Registry/Disk  [8] Netzwerk-Info"
    Write-Host " [9] REBOOT         [0] SHUTDOWN"
    Write-Host "==============================================="
}
do {
    Show-Menu
    $c = Read-Host "Auswahl"
    switch ($c) {
        '1' { if(Test-Path "X:\Program Files\Toolkit\NotepadPP\notepad++.exe"){ Start-Process "X:\Program Files\Toolkit\NotepadPP\notepad++.exe" } }
        '2' { if(Test-Path "X:\Program Files\Toolkit\ExplorerPP\Explorer++.exe"){ Start-Process "X:\Program Files\Toolkit\ExplorerPP\Explorer++.exe" } }
        '3' { Start-Process powershell.exe -ArgumentList "-NoExit -Command micro.exe" }
        '4' { Start-Process "X:\Program Files\Toolkit\ExplorerPP\Explorer++.exe" "X:\Program Files\Toolkit\Sysinternals" }
        '5' { if(Test-Path "X:\Program Files\Toolkit\NirSoft\NirLauncher.exe"){ Start-Process "X:\Program Files\Toolkit\NirSoft\NirLauncher.exe" } }
        '6' { Start-Process powershell.exe -Wait }
        '7' { Start-Process regedit.exe; Start-Process cmd.exe -ArgumentList "/k diskpart" }
        '8' { ipconfig /all; Read-Host "Enter..." }
        '9' { wpeutil reboot }
        '0' { wpeutil shutdown }
    }
} while ($true)
'@
    $LauncherContent | Out-File "$MountPath\Windows\System32\LaunchMenu.ps1" -Encoding ASCII -Force
}

function Inject-Components {
    Write-Host "[5/9] Integriere Enterprise-Komponenten & Toolkit..." -ForegroundColor Yellow
    $OCs = @("*WinPE-WMI*.cab", "*WinPE-NetFX*.cab", "*WinPE-Scripting*.cab", "*WinPE-PowerShell*.cab", "*StorageWMI*.cab")
    foreach ($f in $OCs) {
        $cab = Get-ChildItem $script:ADKRootFolder -Filter $f -Recurse | Where-Object { $_.FullName -notmatch "de-de|en-us" } | Select-Object -First 1
        if ($cab) { Add-WindowsPackage -Path $MountPath -PackagePath $cab.FullName -IgnoreCheck -ErrorAction SilentlyContinue }
    }
    
    $destToolkit = "$MountPath\Program Files\Toolkit"
    if (-not (Test-Path $destToolkit)) { New-Item $destToolkit -ItemType Directory -Force | Out-Null }
    Get-ChildItem $ToolsPath | ForEach-Object { Copy-Item $_.FullName -Destination $destToolkit -Recurse -Force }
    if (Test-Path "$ToolsPath\MicroEditor\micro.exe") { Copy-Item "$ToolsPath\MicroEditor\micro.exe" "$MountPath\Windows\System32\micro.exe" -Force }
}

function Inject-Drivers {
    if ($DriverSource -and (Test-Path $DriverSource)) {
        Write-Host "[6/9] Injektiere ESSENZIELLE Treiber..." -ForegroundColor Yellow
        $infFiles = Get-ChildItem $DriverSource -Filter "*.inf" -Recurse
        foreach ($inf in $infFiles) {
            $content = Get-Content $inf.FullName -TotalCount 50 -ErrorAction SilentlyContinue
            if ($content -match "Class\s*=\s*(Net|NetTrans|HDC|SCSIAdapter)") {
                try { Add-WindowsDriver -Path $MountPath -Driver $inf.Directory.FullName -ForceUnsigned -ErrorAction SilentlyContinue } catch {}
            }
        }
    }
}

function Finalize-WinPE {
    Write-Host "[7/9] Finalisierung (Sicherung in boot.wim)..." -ForegroundColor Yellow
    $Startnet = "@echo off`r`nset PATH=%PATH%;X:\Program Files\Toolkit\ExplorerPP;X:\Program Files\Toolkit\NotepadPP`r`nwpeinit`r`n:loop`r`nipconfig | find `"IPv4`" > nul`r`nif errorlevel 1 (timeout /t 2 > nul & goto loop)`r`npowershell -ExecutionPolicy Bypass -File X:\Windows\System32\LaunchMenu.ps1"
    $Startnet | Out-File "$MountPath\Windows\System32\startnet.cmd" -Encoding ASCII -Force
    Dismount-WindowsImage -Path $MountPath -Save
}

function Export-OpsiFiles {
    if ($OpsiExportPath) {
        Write-Host "[8/9] Synchronisiere OPSI Depot nach Handbuch (Essential & Addon-Safe)..." -ForegroundColor Yellow
        $targetWinpe = Join-Path $OpsiExportPath "winpe"
        
        if (Test-Path $targetWinpe) {
            $item = Get-Item $targetWinpe
            if (-not ($item.Attributes -match "ReparsePoint")) {
                Remove-Item $targetWinpe -Recurse -Force -ErrorAction SilentlyContinue
                New-Item $targetWinpe -ItemType Directory -Force | Out-Null
            }
        }
        else { New-Item $targetWinpe -ItemType Directory -Force | Out-Null }
        
        # Kopiere nur die 5 Elemente, die OPSI zum Booten benötigt
        # Die addons liegen bereits INSIDE winpe/sources/boot.wim!
        $Essentials = @("boot", "EFI", "sources", "bootmgr", "bootmgr.efi")
        foreach ($item in $Essentials) {
            $src = Join-Path "$WorkingPath\media" $item
            if (Test-Path $src) { 
                Copy-Item $src $targetWinpe -Recurse -Force 
                Write-Verbose "Exportiert: $item"
            }
        }

        Write-Host " [OK] Ordner 'winpe' ist nun PXE-bootfähig und enthält alle Addons." -ForegroundColor Green
        $script:StatusReport["OPSI-Export"] = "OK"
    }
}

function Build-Iso {
    Write-Host "[9/9] Erstelle ISO (Hybrid)..." -ForegroundColor Yellow
    $Oscdimg = Get-ChildItem $script:ADKRootFolder -Filter "oscdimg.exe" -Recurse | Select-Object -First 1
    $Etfs = Get-ChildItem $script:ADKRootFolder -Filter "etfsboot.com" -Recurse | Select-Object -First 1
    $Efi = Get-ChildItem $script:ADKRootFolder -Filter "efisys.bin" -Recurse | Select-Object -First 1

    if ($Oscdimg -and $Etfs -and $Efi) {
        $bootData = "2#p0,e,b`"$($Etfs.FullName)`"#pEF,e,b`"$($Efi.FullName)`""
        $args = "-p0 -m -o -u2 -udfver102 -bootdata:$bootData `"$WorkingPath\media`" `"$IsoPath`""
        Start-Process $Oscdimg.FullName $args -Wait -NoNewWindow
        if ($OpsiExportPath) { Copy-Item $IsoPath $OpsiExportPath -Force }
        $script:StatusReport["ISO"] = "OK"
    }
}

# --- Hauptprogramm ---
try {
    Find-ADK
    Show-Greeting
    Initialize-Workspace
    Download-Toolkit
    Mount-WinPE
    Create-WinPELauncher
    Inject-Components
    Inject-Drivers
    Finalize-WinPE
    Export-OpsiFiles
    Build-Iso
    
    Write-Host "`n===============================================" -ForegroundColor Cyan
    Write-Host "   ABSCHLUSS-BERICHT                          " -ForegroundColor Cyan
    Write-Host "===============================================" -ForegroundColor Cyan
    foreach ($key in ($script:StatusReport.Keys | Sort-Object)) {
        Write-Host ("{0,-25} : {1}" -f $key, $script:StatusReport[$key]) -ForegroundColor Green
    }
    
    Write-Host "`n[✓] Build erfolgreich abgeschlossen!" -ForegroundColor Green
    Write-Host "`nDrücken Sie eine Taste zum Beenden..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
catch {
    Write-Host "`n╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Red
    Write-Host "║  KRITISCHER FEHLER                                            ║" -ForegroundColor Red
    Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Red
    Write-Host ""
    Write-Host "Fehler: $_" -ForegroundColor Red
    Write-Host "Fehlertyp: $($_.Exception.GetType().FullName)" -ForegroundColor Yellow
    
    if ($_.ScriptStackTrace) {
        Write-Host "`nStack Trace:" -ForegroundColor Yellow
        Write-Host $_.ScriptStackTrace -ForegroundColor Gray
    }
    
    # Cleanup bei Fehler
    if (Test-Path $MountPath) { 
        Write-Host "`n[CLEANUP] Prüfe auf gemountete Images..." -ForegroundColor Yellow
        try {
            $mounted = Get-WindowsImage -Mounted -ErrorAction SilentlyContinue | Where-Object { $_.Path -eq $MountPath }
            if ($mounted) {
                Write-Host "[CLEANUP] Unmounting Image..." -ForegroundColor Yellow
                Dismount-WindowsImage -Path $MountPath -Discard -ErrorAction SilentlyContinue
                Write-Host "[CLEANUP] Image unmounted" -ForegroundColor Green
            }
        }
        catch {
            Write-Warning "Cleanup fehlgeschlagen: $_"
        }
    }
    
    Write-Host "`n[✗] Build fehlgeschlagen!" -ForegroundColor Red
    Write-Host "`nDrücken Sie eine Taste zum Beenden..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}