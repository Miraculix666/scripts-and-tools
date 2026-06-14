# File Analyzer Script
# Description: Analyzes files by age and size, finds duplicates, and generates HTML reports

param(
    [Parameter(HelpMessage="Path to analyze (default: current directory)")]
    [string]$TargetPath = ".",
    
    [Parameter(HelpMessage="Output directory for results")]
    [string]$OutputDirectory = ".\FileAnalysis",
    
    [Parameter(HelpMessage="Log level (1=Error, 2=Warning, 3=Info, 4=Verbose)")]
    [ValidateRange(1,4)]
    [int]$LogLevel = 3,
    
    [Parameter(HelpMessage="Age thresholds in years")]
    [hashtable]$AgeThresholds = @{
        Thirty = 30
        Ten = 10
        Six = 6
        Five = 5
    },
    
    [Parameter(HelpMessage="Size thresholds in bytes")]
    [hashtable]$SizeThresholds = @{
        TenMB = 10MB
        HundredMB = 100MB
        OneGB = 1GB
    }
)

# Script initialization
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
$script:LogFile = Join-Path $OutputDirectory "FileAnalysis.log"

# Ensure output directory exists
if (-not (Test-Path -Path $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}

# Logging function with levels
function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter()]
        [ValidateSet('Error', 'Warning', 'Info', 'Verbose')]
        [string]$Level = 'Info',
        
        [Parameter()]
        [switch]$NoConsole
    )
    
    $levelNum = switch ($Level) {
        'Error' { 1 }
        'Warning' { 2 }
        'Info' { 3 }
        'Verbose' { 4 }
    }
    
    if ($levelNum -le $LogLevel) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logMessage = "[$timestamp] [$Level] $Message"
        
        # Always write to log file
        Add-Content -Path $script:LogFile -Value $logMessage
        
        # Write to console with color if not suppressed
        if (-not $NoConsole) {
            $color = switch ($Level) {
                'Error' { 'Red' }
                'Warning' { 'Yellow' }
                'Info' { 'White' }
                'Verbose' { 'Gray' }
            }
            Write-Host $logMessage -ForegroundColor $color
        }
    }
}

# Function to format file size
function Format-FileSize {
    param([long]$Size)
    
    if ($Size -gt 1TB) { return "{0:N2} TB" -f ($Size / 1TB) }
    if ($Size -gt 1GB) { return "{0:N2} GB" -f ($Size / 1GB) }
    if ($Size -gt 1MB) { return "{0:N2} MB" -f ($Size / 1MB) }
    if ($Size -gt 1KB) { return "{0:N2} KB" -f ($Size / 1KB) }
    return "$Size B"
}

# Function to analyze files by age
function Get-FilesByAge {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [Parameter(Mandatory=$true)]
        [datetime]$Threshold,
        [Parameter(Mandatory=$true)]
        [string]$OutputFileName
    )
    
    try {
        Write-Log "Analyzing files older than $(($Threshold).ToString('yyyy-MM-dd'))" -Level Verbose
        
        $outputPath = Join-Path $OutputDirectory "$OutputFileName.txt"
        $oldFiles = Get-ChildItem -Path $Path -Recurse -File |
            Where-Object { $_.LastWriteTime -lt $Threshold }
        
        if ($oldFiles) {
            $oldFiles | Select-Object -ExpandProperty FullName | Set-Content -Path $outputPath
            Write-Log "Found $($oldFiles.Count) files older than threshold" -Level Info
        } else {
            Write-Log "No files found older than threshold" -Level Info
            "" | Set-Content -Path $outputPath
        }
        
        return $outputPath
    }
    catch {
        Write-Log "Error analyzing files by age: $_" -Level Error
        throw
    }
}

# Function to analyze files by size
function Get-FilesBySize {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [Parameter(Mandatory=$true)]
        [long]$Threshold,
        [Parameter(Mandatory=$true)]
        [string]$OutputFileName
    )
    
    try {
        Write-Log "Analyzing files larger than $(Format-FileSize $Threshold)" -Level Verbose
        
        $outputPath = Join-Path $OutputDirectory "$OutputFileName.txt"
        $largeFiles = Get-ChildItem -Path $Path -Recurse -File |
            Where-Object { $_.Length -gt $Threshold }
        
        if ($largeFiles) {
            $largeFiles | ForEach-Object {
                "{0} ({1})" -f $_.FullName, (Format-FileSize $_.Length)
            } | Set-Content -Path $outputPath
            Write-Log "Found $($largeFiles.Count) large files" -Level Info
        } else {
            Write-Log "No files found larger than threshold" -Level Info
            "" | Set-Content -Path $outputPath
        }
        
        return $outputPath
    }
    catch {
        Write-Log "Error analyzing files by size: $_" -Level Error
        throw
    }
}

# Function to find duplicate files
function Find-Duplicates {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    
    try {
        Write-Log "Searching for duplicate files" -Level Info
        
        $outputPath = Join-Path $OutputDirectory "Duplicates.txt"
        $files = Get-ChildItem -Path $Path -Recurse -File
        $duplicates = $files | Group-Object -Property Name | Where-Object { $_.Count -gt 1 }
        
        if ($duplicates) {
            $duplicates | ForEach-Object {
                "Duplicate set: $($_.Name)" | Set-Content -Path $outputPath -Append
                $_.Group | ForEach-Object {
                    "  $($_.FullName) ($(Format-FileSize $_.Length), Modified: $($_.LastWriteTime))"
                } | Add-Content -Path $outputPath
                "" | Add-Content -Path $outputPath
            }
            Write-Log "Found $($duplicates.Count) sets of duplicate files" -Level Info
        } else {
            Write-Log "No duplicate files found" -Level Info
            "No duplicate files found." | Set-Content -Path $outputPath
        }
        
        return $outputPath
    }
    catch {
        Write-Log "Error finding duplicates: $_" -Level Error
        throw
    }
}

# Function to generate HTML report header
function Get-HTMLTemplateHeader {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Title
    )
    
    return @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>File Analysis - $Title</title>
    <style>
        :root {
            --bg-primary: #121212;
            --bg-secondary: #1E1E1E;
            --text-primary: #FFFFFF;
            --text-secondary: #BB86FC;
            --accent: #03DAC6;
        }
        
        body {
            background-color: var(--bg-primary);
            color: var(--text-primary);
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            line-height: 1.6;
            margin: 0;
            padding: 2rem;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
        }
        
        h1, h2 {
            color: var(--text-secondary);
            margin-bottom: 1rem;
        }
        
        .summary {
            background-color: var(--bg-secondary);
            border-radius: 8px;
            padding: 1rem;
            margin-bottom: 2rem;
        }
        
        .file-list {
            background-color: var(--bg-secondary);
            border-radius: 8px;
            padding: 1rem;
            margin-bottom: 1rem;
        }
        
        .file-item {
            display: flex;
            align-items: center;
            padding: 0.5rem;
            border-bottom: 1px solid #333;
        }
        
        .file-item:last-child {
            border-bottom: none;
        }
        
        .file-icon {
            margin-right: 1rem;
            color: var(--accent);
        }
        
        .controls {
            margin-bottom: 2rem;
        }
        
        button {
            background-color: var(--accent);
            color: var(--bg-primary);
            border: none;
            padding: 0.5rem 1rem;
            border-radius: 4px;
            cursor: pointer;
            font-weight: bold;
            margin-right: 1rem;
        }
        
        button:hover {
            opacity: 0.9;
        }
        
        .archive {
            display: none;
        }
        
        .stats {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 1rem;
            margin-bottom: 2rem;
        }
        
        .stat-card {
            background-color: var(--bg-secondary);
            padding: 1rem;
            border-radius: 8px;
            text-align: center;
        }
        
        .stat-value {
            font-size: 1.5rem;
            color: var(--accent);
            margin: 0.5rem 0;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>$Title</h1>
        
        <div class="controls">
            <button onclick="toggleArchive()">Toggle Archived Files</button>
            <button onclick="expandAll()">Expand All</button>
            <button onclick="collapseAll()">Collapse All</button>
        </div>
        
        <div class="file-list">
"@
}

# Function to generate HTML report footer
function Get-HTMLTemplateFooter {
    return @"
        </div>
    </div>
    <script>
        function toggleArchive() {
            const archives = document.getElementsByClassName('archive');
            for (let item of archives) {
                item.style.display = item.style.display === 'none' ? 'flex' : 'none';
            }
        }
        
        function expandAll() {
            const items = document.getElementsByClassName('file-item');
            for (let item of items) {
                item.style.display = 'flex';
            }
        }
        
        function collapseAll() {
            const items = document.getElementsByClassName('file-item');
            for (let item of items) {
                if (!item.classList.contains('archive')) {
                    item.style.display = 'flex';
                } else {
                    item.style.display = 'none';
                }
            }
        }
    </script>
</body>
</html>
"@
}

# Function to generate HTML report
function New-HTMLReport {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Title,
        [Parameter(Mandatory=$true)]
        [string[]]$FilePaths
    )

    try {
        $outputPath = Join-Path $OutputDirectory "$Title.html"

        $htmlHeader = Get-HTMLTemplateHeader -Title $Title
        $htmlFooter = Get-HTMLTemplateFooter

        $fileItemsHtml = foreach ($filePath in $FilePaths) {
            if (Test-Path $filePath) {
                $files = Get-Content $filePath
                foreach ($file in $files) {
                    $isArchived = $file -match "\\---ARCHIVE---\\"
                    $class = if ($isArchived) { 'archive' } else { '' }
                    @"
                    <div class="file-item $class">
                        <span class="file-icon">📄</span>
                        <span class="file-path">$([System.Web.HttpUtility]::HtmlEncode($file))</span>
                    </div>
"@
                }
            }
        }

        $htmlContent = $htmlHeader + ($fileItemsHtml -join "`n") + $htmlFooter
        
        $htmlContent | Out-File -FilePath $outputPath -Encoding UTF8
        Write-Log "Generated HTML report: $outputPath" -Level Info
        return $outputPath
    }
    catch {
        Write-Log "Error generating HTML report: $_" -Level Error
        throw
    }
}

# Main execution block
try {
    Write-Log "Starting file analysis on path: $TargetPath" -Level Info
    Write-Log "Results will be saved to: $OutputDirectory" -Level Info
    
    # Create date thresholds
    $dateThresholds = @{}
    foreach ($key in $AgeThresholds.Keys) {
        $dateThresholds[$key] = (Get-Date).AddYears(-$AgeThresholds[$key])
    }
    
    # Analyze files by age
    $ageResults = @()
    foreach ($key in $dateThresholds.Keys) {
        $ageResults += Get-FilesByAge -Path $TargetPath -Threshold $dateThresholds[$key] -OutputFileName "OlderThan$($AgeThresholds[$key])Years"
    }
    
    # Analyze files by size
    $sizeResults = @()
    foreach ($key in $SizeThresholds.Keys) {
        $sizeResults += Get-FilesBySize -Path $TargetPath -Threshold $SizeThresholds[$key] -OutputFileName "LargerThan$key"
    }
    
    # Find duplicates
    $duplicatesPath = Find-Duplicates -Path $TargetPath
    
    # Generate HTML reports
    New-HTMLReport -Title "Files by Age" -FilePaths $ageResults
    New-HTMLReport -Title "Files by Size" -FilePaths $sizeResults
    New-HTMLReport -Title "Duplicate Files" -FilePaths @($duplicatesPath)
    
    Write-Log "Analysis completed successfully" -Level Info
    Write-Log "Reports are available in: $OutputDirectory" -Level Info
}
catch {
    Write-Log "Critical error during analysis: $_" -Level Error
    throw
}
