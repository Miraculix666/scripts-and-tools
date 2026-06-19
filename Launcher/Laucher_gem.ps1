#--------------------------------------------------------------------------------
# SCRIPT NAME: Launcher.ps1
# DESCRIPTION: A portable, folder-structure-based Application Launcher using
#              PowerShell and Windows Forms (WinForms).
# REQUIREMENTS: Runs on any Windows system with PowerShell 5.1+ and .NET Framework.
#              No external binaries or modules required.
#--------------------------------------------------------------------------------

# --- Global Configuration ---
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
$AppsPath = Join-Path -Path $ScriptRoot -ChildPath "Apps"
$CredFilePath = Join-Path -Path $ScriptRoot -ChildPath "SecureAdmin.cred"
$AdminFlagFile = "exeasadmin.txt"
$InfoFile = "BalloonInfo.txt"
$IconFile = "AppIcon.ico" # Search order: AppIcon.ico, AppIcon.png, AppIcon.jpg

# --- 1. Load Assemblies ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- 2. Credential Management Functions ---

Function Get-AdminCredential {
    <#
    .SYNOPSIS
    Loads saved admin credentials or prompts the user to create them.
    #>
    Write-Host "Checking for saved credentials..." -ForegroundColor Yellow

    if (Test-Path $CredFilePath) {
        try {
            # Import-Clixml securely decrypts the credentials (tied to the current user)
            $Credential = Import-Clixml -Path $CredFilePath
            Write-Host "Credentials loaded successfully." -ForegroundColor Green
            return $Credential
        }
        catch {
            Write-Host "Error loading credentials. File may be corrupted or encrypted for a different user." -ForegroundColor Red
            Remove-Item $CredFilePath -Force -ErrorAction SilentlyContinue
        }
    }

    # If file doesn't exist or failed to load, prompt the user
    Write-Host "Saved credentials not found. Please enter admin credentials for elevated tasks." -ForegroundColor Cyan
    $Credential = Get-Credential -Message "Admin-Konto für erweiterte Ausführung eingeben:"

    try {
        # Securely save the new credentials
        $Credential | Export-Clixml -Path $CredFilePath -Force
        Write-Host "Credentials saved securely for future use." -ForegroundColor Green
        return $Credential
    }
    catch {
        Write-Host "Warning: Could not save credentials. Elevation will fail next time unless re-entered." -ForegroundColor Yellow
        return $Credential
    }
}

# --- 3. Data Structure and Scanning ---

Function Scan-Applications {
    <#
    .SYNOPSIS
    Scans the Apps directory and builds a hierarchical list of applications.
    #>
    Write-Host "Scanning applications in: $AppsPath" -ForegroundColor Cyan
    $AppStructure = @()

    if (-not (Test-Path $AppsPath)) {
        Write-Host "Error: The 'Apps' directory was not found. Please create it next to Launcher.ps1." -ForegroundColor Red
        return $AppStructure
    }

    $CategoryFolders = Get-ChildItem -Path $AppsPath -Directory -ErrorAction SilentlyContinue

    foreach ($CategoryDir in $CategoryFolders) {
        $Category = @{
            Name = $CategoryDir.Name
            Nodes = @()
        }

        # Find potential applications (scripts, executables, shortcuts) in the category folder
        $AppItems = Get-ChildItem -Path $CategoryDir.FullName -Exclude $InfoFile, $AdminFlagFile -ErrorAction SilentlyContinue | Where-Object {
            -not $_.PSIsContainer -and ($_.Extension -in @(".ps1", ".exe", ".cmd", ".bat", ".lnk"))
        }

        foreach ($Item in $AppItems) {
            $AppDir = $Item.Directory.FullName
            
            # 1. Get Description (Balloon Text)
            $DescPath = Join-Path $AppDir $InfoFile
            $Description = if (Test-Path $DescPath) {
                (Get-Content $DescPath -Raw -Encoding UTF8) -join "`r`n"
            } else {
                "Keine Beschreibung gefunden."
            }

            # 2. Check for Admin Flag
            $NeedsAdmin = Test-Path (Join-Path $AppDir $AdminFlagFile)
            
            # 3. Find Icon
            $IconPath = ""
            if (Test-Path (Join-Path $AppDir $IconFile)) { $IconPath = Join-Path $AppDir $IconFile }
            else { 
                # Check common alternative image formats
                if (Test-Path (Join-Path $AppDir "AppIcon.png")) { $IconPath = Join-Path $AppDir "AppIcon.png" }
                elseif (Test-Path (Join-Path $AppDir "AppIcon.jpg")) { $IconPath = Join-Path $AppDir "AppIcon.jpg" }
            }

            $App = [PSCustomObject]@{
                Name = $Item.Name
                Path = $Item.FullName
                Description = $Description
                NeedsAdmin = $NeedsAdmin
                IconPath = $IconPath
            }
            $Category.Nodes += $App
        }
        $AppStructure += $Category
    }
    return $AppStructure
}

# --- 4. Execution Logic ---

Function Launch-Application {
    param(
        [Parameter(Mandatory=$true)]$AppObject,
        [Parameter(Mandatory=$true)]$AdminCredential
    )

    $AppPath = $AppObject.Path
    $Arguments = ""
    $WorkingDir = Split-Path -Parent $AppPath

    Write-Host "Starte Anwendung: $($AppObject.Name)" -ForegroundColor Green

    # Handle PowerShell scripts - requires explicit powershell.exe call
    if ($AppPath.ToLower().EndsWith(".ps1")) {
        $Executable = "powershell.exe"
        # Use -File for security; -ExecutionPolicy Bypass is needed for local scripts
        $Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$AppPath`""
        $ProcessParams = @{
            FilePath = $Executable
            ArgumentList = $Arguments
            WorkingDirectory = $WorkingDir
            WindowStyle = "Normal"
        }
    }
    else {
        # Handle Executables, Batch files, Links, etc.
        $ProcessParams = @{
            FilePath = $AppPath
            WorkingDirectory = $WorkingDir
            WindowStyle = "Normal"
        }
    }
    
    # Check for Admin Elevation requirement
    if ($AppObject.NeedsAdmin) {
        Write-Host "-> Wird mit Admin-Konto gestartet..." -ForegroundColor Yellow
        $ProcessParams.Add('Credential', $AdminCredential)
    }

    try {
        # Start the process with or without credentials
        Start-Process @ProcessParams -NoNewWindow
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Starten von '$($AppObject.Name)': $($_.Exception.Message)", "Startfehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# --- 5. GUI Definition and Setup ---

Function New-LauncherForm {
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Portable Script Launcher - PS-Coding"
    $Form.Size = New-Object System.Drawing.Size(600, 700)
    $Form.StartPosition = "CenterScreen"
    $Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $Form.MaximizeBox = $false
    $Form.MinimizeBox = $false
    return $Form
}

Function New-LauncherControls {
    param($Form)
    
    $TreeView = New-Object System.Windows.Forms.TreeView
    $TreeView.Location = New-Object System.Drawing.Point(10, 10)
    $TreeView.Size = New-Object System.Drawing.Size(565, 400)
    $TreeView.Dock = [System.Windows.Forms.DockStyle]::Top
    $TreeView.ImageList = New-Object System.Windows.Forms.ImageList
    [void]$Form.Controls.Add($TreeView)

    [void]$TreeView.ImageList.Images.Add("default", [System.Drawing.SystemIcons]::Application)

    $DescBox = New-Object System.Windows.Forms.RichTextBox
    $DescBox.Text = "Wählen Sie ein Element zum Anzeigen der Beschreibung aus."
    $DescBox.Location = New-Object System.Drawing.Point(10, 420)
    $DescBox.Size = New-Object System.Drawing.Size(565, 120)
    $DescBox.ReadOnly = $true
    [void]$Form.Controls.Add($DescBox)

    $LaunchButton = New-Object System.Windows.Forms.Button
    $LaunchButton.Text = "Starten"
    $LaunchButton.Location = New-Object System.Drawing.Point(455, 555)
    $LaunchButton.Size = New-Object System.Drawing.Size(120, 40)
    $LaunchButton.Enabled = $false
    [void]$Form.Controls.Add($LaunchButton)

    $StatusLabel = New-Object System.Windows.Forms.Label
    $StatusLabel.Text = "Bereit."
    $StatusLabel.Location = New-Object System.Drawing.Point(10, 555)
    $StatusLabel.Size = New-Object System.Drawing.Size(400, 40)
    $StatusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    [void]$Form.Controls.Add($StatusLabel)

    return @{
        TreeView = $TreeView
        DescBox = $DescBox
        LaunchButton = $LaunchButton
        StatusLabel = $StatusLabel
    }
}

Function Populate-LauncherTreeView {
    param($TreeView, $AppStructure)
    $IconIndex = 0
    foreach ($Category in $AppStructure) {
        $CatNode = New-Object System.Windows.Forms.TreeNode ($Category.Name)
        
        foreach ($App in $Category.Nodes) {
            
            $AppNode = New-Object System.Windows.Forms.TreeNode ($App.Name)
            $AppNode.Tag = $App # Store the entire App object in the Tag property

            $NodeText = $App.Name
            if ($App.NeedsAdmin) {
                $NodeText += " (Admin)"
                $AppNode.ForeColor = [System.Drawing.Color]::Firebrick
            }
            $AppNode.Text = $NodeText

            # Handle Icon Loading
            if (Test-Path $App.IconPath) {
                try {
                    # Load the icon and add it to the ImageList
                    $Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($App.Path)
                    $IconKey = "Icon_$IconIndex"
                    [void]$TreeView.ImageList.Images.Add($IconKey, $Icon.ToBitmap())
                    
                    # Set the node's image and selected image index
                    $AppNode.ImageKey = $IconKey
                    $AppNode.SelectedImageKey = $IconKey
                    $IconIndex++
                }
                catch {
                    # Fallback to default icon
                    $AppNode.ImageKey = "default"
                    $AppNode.SelectedImageKey = "default"
                }
            } else {
                $AppNode.ImageKey = "default"
                $AppNode.SelectedImageKey = "default"
            }
            
            [void]$CatNode.Nodes.Add($AppNode)
        }
        [void]$TreeView.Nodes.Add($CatNode)
    }

    $TreeView.ExpandAll()
}

Function Register-LauncherEvents {
    param($TreeView, $DescBox, $LaunchButton, $StatusLabel, $AdminCreds)

    $TreeView.Add_AfterSelect({
        $SelectedNode = $TreeView.SelectedNode
        
        # Only enable button and show info for leaf nodes (applications)
        if ($SelectedNode.Tag -is [PSCustomObject]) {
            $AppObject = $SelectedNode.Tag
            $DescBox.Text = $AppObject.Description
            $LaunchButton.Enabled = $true
            
            $StatusLabel.Text = "Startbereit: $($AppObject.Name)"
            if ($AppObject.NeedsAdmin) {
                $StatusLabel.Text += " (Admin-Rechte erforderlich)"
            }
        } else {
            # Selected a Category Node
            $DescBox.Text = "Wählen Sie eine Anwendung aus der Kategorie aus."
            $LaunchButton.Enabled = $false
            $StatusLabel.Text = "Bereit."
        }
    }.GetNewClosure())

    $LaunchButton.Add_Click({
        $AppObject = $TreeView.SelectedNode.Tag
        
        if ($null -ne $AppObject) {
            # Pass the credentials only if needed
            $CredsToUse = if ($AppObject.NeedsAdmin) { $AdminCreds } else { $null }
            
            Launch-Application -AppObject $AppObject -AdminCredential $CredsToUse
        }
    }.GetNewClosure())
}

Function Build-GUI {

    # Prepare data and credentials
    $AppStructure = Scan-Applications
    $AdminCreds = $null
    if ($AppStructure.Where({$_.Nodes.NeedsAdmin}).Count -gt 0) {
        $AdminCreds = Get-AdminCredential
    }

    # --- 5.1 Main Form Setup ---
    $Form = New-LauncherForm

    # --- 5.2 to 5.5 Control Setup ---
    $Controls = New-LauncherControls -Form $Form
    $TreeView = $Controls.TreeView
    $DescBox = $Controls.DescBox
    $LaunchButton = $Controls.LaunchButton
    $StatusLabel = $Controls.StatusLabel

    # --- 5.6 Populate TreeView ---
    Populate-LauncherTreeView -TreeView $TreeView -AppStructure $AppStructure

    # --- 5.7 Event Handlers ---
    Register-LauncherEvents -TreeView $TreeView -DescBox $DescBox -LaunchButton $LaunchButton -StatusLabel $StatusLabel -AdminCreds $AdminCreds

    # --- 5.8 Show the GUI ---
    [void]$Form.ShowDialog()
}

# --- 6. Main Script Execution ---
Build-GUI
