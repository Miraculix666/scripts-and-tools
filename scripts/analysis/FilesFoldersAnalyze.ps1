<#
.SYNOPSIS
    Analyzes storage usage by examining file age, size, duplicates, and directory sizes with caching and export capabilities.
.DESCRIPTION
    File-Analyzer-Pro.ps1 is a comprehensive PowerShell script designed to analyze a specified directory path.
    It performs several checks:
    - Files by Age: Identifies files older than specified thresholds (in years).
    - Files by Size: Identifies files larger than specified thresholds.
    - Duplicate Files: Finds potential duplicates based on name/size (default) or file hash (optional).
    - Directory Size Tree: Calculates the total size of each directory and its subdirectories.

    The script features a caching system to significantly speed up subsequent runs.
    Results are compiled into an interactive HTML report and can be exported to CSV and TXT formats.
.PARAMETER TargetPath
    The directory path to be analyzed. If not provided, the script will prompt for it.
.PARAMETER OutputDirectory
    The directory where the analysis report, cache, and exports will be saved.
.PARAMETER ExcludePath
    An array of full directory paths to exclude from the analysis.
.PARAMETER AgeThresholds
    A hashtable defining the age categories in years.
.PARAMETER SizeThresholds
    A hashtable defining the size categories.
.PARAMETER EnableHashCheck
    Enables a deep, content-based duplicate analysis using SHA256 hashes. Slow and resource-intensive.
.PARAMETER ForceScan
    Forces a new file scan, ignoring any existing cache file.
.PARAMETER ExportFormats
    Exports the analysis data to specified formats. Valid options are 'CSV', 'TXT'.
.PARAMETER Silent
    Suppresses the final confirmation prompt before starting the analysis.
.EXAMPLE
    .\File-Analyzer-Pro.ps1 -TargetPath "D:\Files" -ExcludePath "D:\Files\Archive" -ExportFormats 'CSV','TXT'
    Analyzes "D:\Files" (excluding the Archive subfolder), uses the cache if available, and exports the results to CSV and TXT files.
.EXAMPLE
    .\File-Analyzer-Pro.ps1 -TargetPath "\\server\share" -ForceScan -EnableHashCheck
    Forces a new, deep analysis of a network share, rebuilding the cache and using hash checking for duplicates.
.NOTES
    Author: Your Name / PS-Coding Assistant
    Version: 3.2
    Last Modified: 2025-08-28
    Requires PowerShell 5.1 or higher. Works with PowerShell 7+.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false, HelpMessage = "Path to analyze. The script will prompt if not provided.")]
    [string]$TargetPath,

    [Parameter(Mandatory = $false, HelpMessage = "Output directory for results.")]
    [string]$OutputDirectory = ".\FileAnalysis",

    [Parameter(Mandatory = $false, HelpMessage = "Array of full paths to exclude from the analysis.")]
    [string[]]$ExcludePath,

    [Parameter(Mandatory = $false, HelpMessage = "Age thresholds in years.")]
    [hashtable]$AgeThresholds = @{
        'Over10Years' = 10
        'Over5Years'  = 5
        'Over1Year'   = 1
    },

    [Parameter(Mandatory = $false, HelpMessage = "Size thresholds in bytes.")]
    [hashtable]$SizeThresholds = @{
        'Over1GB'   = 1GB
        'Over100MB' = 100MB
        'Over10MB'  = 10MB
    },
    
    [Parameter(Mandatory = $false, HelpMessage = "Enables deep duplicate analysis by file content hash.")]
    [switch]$EnableHashCheck,

    [Parameter(Mandatory = $false, HelpMessage = "Forces a new file scan, ignoring the cache.")]
    [switch]$ForceScan,

    [Parameter(Mandatory = $false, HelpMessage = "Exports data to specified formats. Valid: 'CSV', 'TXT'.")]
    [ValidateSet('CSV', 'TXT')]
    [string[]]$ExportFormats,

    [Parameter(Mandatory = $false, HelpMessage = "Suppresses the final confirmation prompt.")]
    [switch]$Silent
)

#region SCRIPT INITIALIZATION
$ErrorActionPreference = "Stop"
$script:LogFile = Join-Path $OutputDirectory "FileAnalysis_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$script:CacheFile = Join-Path $OutputDirectory "FileScan.cache.xml"

# --- Ensure output directory exists ---
if (-not (Test-Path -Path $OutputDirectory)) {
    Write-Verbose "Creating output directory: $OutputDirectory"
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}

# --- Interactive Mode: Prompt for TargetPath if not provided ---
if (-not $TargetPath) {
    while (-not ($TargetPath) -or -not (Test-Path -Path $TargetPath -PathType Container)) {
        Write-Host "Please provide the full path to the directory you want to analyze." -ForegroundColor Yellow
        $TargetPath = Read-Host "Target Path"
        if (-not (Test-Path -Path $TargetPath -PathType Container)) {
            Write-Host "Error: The path '$TargetPath' does not exist or is not a directory. Please try again." -ForegroundColor Red
        }
    }
}
else {
    if (-not (Test-Path -Path $TargetPath -PathType Container)) {
        Write-Error "The specified TargetPath '$TargetPath' does not exist or is not a directory."
        return
    }
}
$TargetPath = Resolve-Path -Path $TargetPath

# --- Resolve and validate Exclude Paths ---
$resolvedExcludePaths = @()
if ($ExcludePath) {
    foreach ($path in $ExcludePath) {
        if (Test-Path $path -PathType Container) {
            $resolvedExcludePaths += (Resolve-Path $path).Path
        }
        else {
            Write-Warning "Exclude path '$path' not found or is not a directory. It will be ignored."
        }
    }
}

# --- Pre-analysis confirmation ---
if (-not $Silent) {
    Write-Host "`n--- ANALYSIS CONFIGURATION ---" -ForegroundColor Cyan
    Write-Host "Target Path:          $TargetPath"
    if ($resolvedExcludePaths) {
        Write-Host "Exclude Paths:        $($resolvedExcludePaths -join ', ')"
    }
    Write-Host "Output Directory:     $OutputDirectory"
    $duplicateMethod = if ($EnableHashCheck) { "Content Hash (Slow, Accurate)" } else { "Name & Size (Fast, Default)" }
    Write-Host "Duplicate Check Method: $duplicateMethod"
    $cacheStatus = if ($ForceScan) { "Forcing new scan" } elseif (Test-Path $script:CacheFile) { "Will use existing cache" } else { "No cache found, will perform full scan" }
    Write-Host "Cache Status:         $cacheStatus"
    if ($ExportFormats) {
        Write-Host "Export Formats:       $($ExportFormats -join ', ')"
    }
    Write-Host "Log File:             $script:LogFile"
    Write-Host "-----------------------------"
    $confirmation = Read-Host "Do you want to start the analysis with these settings? (Y/N)"
    if ($confirmation -ne 'Y') {
        Write-Host "Analysis cancelled by user." -ForegroundColor Yellow
        return
    }
}
#endregion

#region HELPER FUNCTIONS
function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $true)]
        [ValidateSet('Error', 'Warning', 'Info', 'Verbose')]
        [string]$Level
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Add-Content -Path $script:LogFile -Value $logMessage
    $color = switch ($Level) {
        'Error'   { 'Red' }
        'Warning' { 'Yellow' }
        'Info'    { 'Green' }
        'Verbose' { 'Gray' }
    }
    Write-Host $logMessage -ForegroundColor $color
}

function Format-FileSize {
    param([long]$Size)
    $units = @("B", "KB", "MB", "GB", "TB", "PB")
    $index = 0
    $formattedSize = $Size
    while ($formattedSize -ge 1024 -and $index -lt ($units.Count - 1)) {
        $formattedSize /= 1024
        $index++
    }
    return "{0:N2} {1}" -f $formattedSize, $units[$index]
}
#endregion

#region CORE ANALYSIS FUNCTIONS

function Get-AllFiles {
    param(
        [string]$Path,
        [string[]]$ExcludePathsArray
    )
    Write-Log "Gathering all files from '$Path'. This may take a while..." "Info"
    $allFiles = @()
    $i = 0
    
    $allDirs = Get-ChildItem -Path $Path -Recurse -Directory -Force -ErrorAction SilentlyContinue
    # Filter out excluded directories
    $filteredDirs = $allDirs | ForEach-Object {
        $currentPath = $_.FullName
        $isExcluded = $false
        foreach ($exclude in $ExcludePathsArray) {
            if ($currentPath.StartsWith($exclude)) {
                $isExcluded = $true
                break
            }
        }
        if (-not $isExcluded) {
            $_
        }
    }
    
    $totalDirs = $filteredDirs.Count + 1
    
    # Process root directory (if not excluded)
    $isRootExcluded = $false
    foreach ($exclude in $ExcludePathsArray) {
        if ($Path -eq $exclude) {
            $isRootExcluded = $true
            break
        }
    }
    if (-not $isRootExcluded) {
        if ($totalDirs -gt 0) { Write-Progress -Activity "Scanning directories" -Status "Processing $Path" -PercentComplete ($i++ * 100 / $totalDirs) }
        $allFiles += Get-ChildItem -Path $Path -File -Force -ErrorAction SilentlyContinue
    }

    # Process subdirectories
    foreach ($dir in $filteredDirs) {
        if ($totalDirs -gt 0) { Write-Progress -Activity "Scanning directories" -Status "Processing $($dir.FullName)" -PercentComplete ($i++ * 100 / $totalDirs) }
        $allFiles += Get-ChildItem -Path $dir.FullName -File -Force -ErrorAction SilentlyContinue
    }
    
    Write-Progress -Activity "Scanning directories" -Completed
    Write-Log "Found a total of $($allFiles.Count) files." "Info"
    return $allFiles
}

function Find-DuplicateFiles {
    param(
        $Files,
        [switch]$UseHashCheck
    )
    
    if ($UseHashCheck) {
        Write-Log "Searching for duplicate files by content (SHA256 Hash). This may be slow." "Info"
        $duplicateSets = @()
        $i = 0
        
        $filesGroupedBySize = $Files | Where-Object { $_.Length -gt 0 } | Group-Object -Property Length | Where-Object { $_.Count -gt 1 }
        
        $measureTotal = $filesGroupedBySize | ForEach-Object { $_.Count } | Measure-Object -Sum
        $totalFilesToHash = if ($measureTotal) { $measureTotal.Sum } else { 0 }

        foreach ($group in $filesGroupedBySize) {
            $filesWithHash = $group.Group | ForEach-Object {
                if ($totalFilesToHash -gt 0) {
                    Write-Progress -Activity "Hashing files" -Status "Processing $($_.Name)" -PercentComplete (++$i * 100 / $totalFilesToHash)
                }
                [pscustomobject]@{
                    File = $_
                    Hash = (Get-FileHash -Algorithm SHA256 -Path $_.FullName).Hash
                }
            }
            
            $hashGroups = $filesWithHash | Group-Object -Property Hash | Where-Object { $_.Count -gt 1 }
            
            if ($hashGroups) {
                $hashGroups | ForEach-Object {
                    $duplicateSets += [pscustomobject]@{
                        Count = $_.Count
                        Group = $_.Group.File
                    }
                }
            }
        }
        Write-Progress -Activity "Hashing files" -Completed
    }
    else {
        Write-Log "Searching for duplicate files by name and size (fast check)." "Info"
        # Ensure result is always an array
        $duplicateSets = @($Files | Group-Object -Property Name, Length | Where-Object { $_.Count -gt 1 })
    }

    Write-Log "Found $($duplicateSets.Count) sets of potential duplicate files." "Info"
    return $duplicateSets
}

function Get-DirectorySizeTree {
    param(
        [string]$Path,
        [string[]]$ExcludePathsArray,
        [int]$Depth = 0
    )
    
    $result = [pscustomobject]@{ Name = (Get-Item $Path).Name; Path = $Path; Depth = $Depth; Files = 0; Folders = 0; SizeOfFiles = 0; TotalSize = 0; Children = [System.Collections.Generic.List[object]]::new() }

    try {
        $items = Get-ChildItem -Path $Path -Force -ErrorAction SilentlyContinue
        $files = $items | Where-Object { -not $_.PSIsContainer }
        $directories = $items | Where-Object { $_.PSIsContainer }

        $result.Files = $files.Count
        
        $measureSize = $files | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue
        $currentSize = if ($measureSize) { $measureSize.Sum } else { 0 }
        
        $result.SizeOfFiles = $currentSize
        $result.TotalSize = $currentSize

        $filteredDirs = $directories | ForEach-Object {
            $isExcluded = $false
            foreach ($exclude in $ExcludePathsArray) {
                if ($_.FullName.StartsWith($exclude)) {
                    $isExcluded = $true
                    break
                }
            }
            if (-not $isExcluded) { $_ }
        }
        $result.Folders = $filteredDirs.Count

        foreach ($dir in $filteredDirs) {
            $childResult = Get-DirectorySizeTree -Path $dir.FullName -ExcludePathsArray $ExcludePathsArray -Depth ($Depth + 1)
            if ($null -ne $childResult) {
                $result.TotalSize += $childResult.TotalSize
                [void]$result.Children.Add($childResult)
            }
        }
    }
    catch {
        Write-Log "Could not access path '$Path'. Reason: $($_.Exception.Message)" "Warning"
        return $null
    }
    
    return $result
}
#endregion

#region EXPORT FUNCTIONS
function Export-AnalysisData {
    param(
        $ReportData,
        [string[]]$Formats,
        [string]$OutDir
    )
    Write-Log "Exporting analysis data..." "Info"

    if ('CSV' -in $Formats) {
        $dupExport = $ReportData.Duplicates | ForEach-Object {
            $set = $_
            $set.Group | ForEach-Object {
                [pscustomobject]@{
                    DuplicateSet = $set.Group[0].Name
                    FilePath = $_.FullName
                    Size_Bytes = $_.Length
                    LastModified = $_.LastWriteTime
                }
            }
        }
        $dupExport | Export-Csv -Path (Join-Path $OutDir "Duplicates.csv") -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        
        $sizeExport = $ReportData.FilesBySize.Values | ForEach-Object { $_ }
        $ageExport = $ReportData.FilesByAge.Values | ForEach-Object { $_ }

        $sizeExport | Select-Object FullName, @{N='Size_MB';E={[math]::Round($_.Length/1MB, 2)}}, LastWriteTime | Export-Csv -Path (Join-Path $OutDir "FilesBySize.csv") -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        $ageExport | Select-Object FullName, @{N='Size_MB';E={[math]::Round($_.Length/1MB, 2)}}, LastWriteTime | Export-Csv -Path (Join-Path $OutDir "FilesByAge.csv") -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        
        function Flatten-Tree($node) {
            if ($null -ne $node) {
                $node | Select-Object Path, Depth, Files, Folders, TotalSize
                $node.Children | ForEach-Object { Flatten-Tree $_ }
            }
        }
        Flatten-Tree $ReportData.DirectoryTree | Export-Csv -Path (Join-Path $OutDir "DirectoryTree.csv") -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        Write-Log "CSV export complete." "Verbose"
    }

    if ('TXT' -in $Formats) {
        $topFolders = if ($null -ne $ReportData.DirectoryTree) {
            $ReportData.DirectoryTree.Children | Sort-Object TotalSize -Descending | Select-Object -First 5 | Format-Table -AutoSize | Out-String
        } else { "Could not be determined." }

        $txtReport = @"
File Analysis Summary
=====================
Analyzed Path: $($ReportData.TargetPath)
Report Date:   $($ReportData.ReportDate)

--- Statistics ---
Total Size Analyzed: $(if ($null -ne $ReportData.DirectoryTree) { Format-FileSize $ReportData.DirectoryTree.TotalSize } else { 'N/A' })
Total Files Found:   $($ReportData.TotalFiles)
Duplicate Sets Found: $($ReportData.Duplicates.Count)
Wasted Space (Duplicates): $(Format-FileSize $ReportData.WastedSpace)

--- Top 5 Largest Folders (Root) ---
$topFolders
--- Files By Age ---
$($ReportData.FilesByAge.Keys | ForEach-Object { "$_ : $($ReportData.FilesByAge[$_].Count) files" } | Out-String)
--- Files By Size ---
$($ReportData.FilesBySize.Keys | ForEach-Object { "$_ : $($ReportData.FilesBySize[$_].Count) files" } | Out-String)
"@
        $txtReport | Set-Content -Path (Join-Path $OutDir "Summary.txt") -Encoding UTF8
        Write-Log "TXT summary export complete." "Verbose"
    }
}
#endregion

#region HTML REPORT GENERATION
function ConvertTo-HtmlTree {
    param($TreeData, [long]$TotalAnalysisSize)
    if ($null -eq $TreeData) { return "" }
    $html = ""
    $percentage = if ($TotalAnalysisSize -gt 0) { ($TreeData.TotalSize / $TotalAnalysisSize) * 100 } else { 0 }
    $indent = $TreeData.Depth * 20
    $html += "<li class='tree-item'><div class='tree-row' style='padding-left: ${indent}px;'>"
    $html += "<span class='toggle'>$(&{if($TreeData.Children.Count -gt 0){'&#9658;'}})</span>"
    # COMPATIBILITY FIX: Use modern [System.Net.WebUtility] for cross-platform HTML encoding
    $html += "<span class='icon'>&#128193;</span><span class='name'>$([System.Net.WebUtility]::HtmlEncode($TreeData.Name))</span>"
    $html += "<span class='size'>$(Format-FileSize $TreeData.TotalSize)</span>"
    $html += "<div class='percentage-bar-container'><div class='percentage-bar' style='width: $($percentage)%'></div></div>"
    $html += "<span class='percentage'>$("{0:N2}" -f $percentage)%</span></div>"
    if ($TreeData.Children.Count -gt 0) {
        $html += "<ul class='nested'>"
        $sortedChildren = $TreeData.Children | Sort-Object -Property TotalSize -Descending
        foreach ($child in $sortedChildren) { $html += ConvertTo-HtmlTree -TreeData $child -TotalAnalysisSize $TotalAnalysisSize }
        $html += "</ul>"
    }
    $html += "</li>"
    return $html
}

function New-HTMLReport {
    param($ReportData)
    Write-Log "Generating HTML report..." "Info"
    $outputPath = Join-Path $OutputDirectory "FileAnalysis_Report.html"
    $ageContent = ""
    foreach ($key in $ReportData.FilesByAge.Keys) {
        $ageContent += "<h4>$([System.Net.WebUtility]::HtmlEncode($key)) ($($ReportData.FilesByAge[$key].Count) files)</h4><ul>"
        $ReportData.FilesByAge[$key] | ForEach-Object { $ageContent += "<li>$([System.Net.WebUtility]::HtmlEncode($_.FullName)) ($($_.LastWriteTime.ToString('dd.MM.yyyy')))</li>" }
        $ageContent += "</ul>"
    }
    if (-not $ageContent) { $ageContent = "<p>No files found matching the age criteria.</p>" }
    $sizeContent = ""
    foreach ($key in $ReportData.FilesBySize.Keys) {
        $sizeContent += "<h4>$([System.Net.WebUtility]::HtmlEncode($key)) ($($ReportData.FilesBySize[$key].Count) files)</h4><ul>"
        $ReportData.FilesBySize[$key] | ForEach-Object { $sizeContent += "<li>$([System.Net.WebUtility]::HtmlEncode($_.FullName)) ($(Format-FileSize $_.Length))</li>" }
        $sizeContent += "</ul>"
    }
    if (-not $sizeContent) { $sizeContent = "<p>No files found matching the size criteria.</p>" }
    $duplicatesContent = ""
    if ($ReportData.Duplicates.Count -gt 0) {
        foreach ($set in $ReportData.Duplicates) {
            $fileSize = Format-FileSize $set.Group[0].Length
            $duplicatesContent += "<h4>Set of $($set.Count) duplicates (Size: $fileSize)</h4><ul>"
            $set.Group | ForEach-Object { $duplicatesContent += "<li>$([System.Net.WebUtility]::HtmlEncode($_.FullName))</li>" }
            $duplicatesContent += "</ul>"
        }
    } else { $duplicatesContent = "<p>No duplicate files found.</p>" }
    
    $treeHtml = ""
    if ($null -ne $ReportData.DirectoryTree) {
        $treeHtml = "<ul class='tree'>" + (ConvertTo-HtmlTree -TreeData $ReportData.DirectoryTree -TotalAnalysisSize $ReportData.DirectoryTree.TotalSize) + "</ul>"
    } else {
        $treeHtml = "<p>Could not generate directory tree. Check log for access errors.</p>"
    }

    $duplicateMethodNote = if ($ReportData.HashCheckEnabled) { "Based on file content (SHA256 Hash)" } else { "Based on file name and size" }
    $htmlTemplate = @"
<!DOCTYPE html><html lang="de"><head><meta charset="UTF-8"><title>File Analysis Report</title><style>:root{--bg-color:#f4f7f6;--header-color:#fff;--text-color:#333;--primary-color:#007bff;--border-color:#dee2e6;--shadow-color:rgba(0,0,0,.1)}body{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background-color:var(--bg-color);color:var(--text-color);margin:0}.container{max-width:1400px;margin:20px auto;padding:20px;background-color:var(--header-color);border-radius:8px;box-shadow:0 2px 10px var(--shadow-color)}h1,h2,h3{color:var(--primary-color);border-bottom:2px solid var(--border-color);padding-bottom:10px}.summary{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:20px;margin-bottom:20px}.card{background-color:var(--bg-color);padding:20px;border-radius:8px;text-align:center}.card h3{margin-top:0;font-size:1.1em}.card .value{font-size:2em;font-weight:700;color:var(--primary-color)}.tabs{display:flex;border-bottom:1px solid var(--border-color);margin-bottom:20px}.tab-button{background:0 0;border:none;padding:15px 20px;cursor:pointer;font-size:1em}.tab-button.active{border-bottom:3px solid var(--primary-color);font-weight:700}.tab-content{display:none}.tab-content.active{display:block}.note{font-style:italic;color:#666;font-size:.9em;margin-top:-10px;margin-bottom:15px}ul{list-style-type:none;padding-left:20px}li{padding:5px 0;word-break:break-all}.tree{list-style-type:none;padding-left:0}.tree-item .nested{display:none}.tree-item .toggle{cursor:pointer;-webkit-user-select:none;-moz-user-select:none;user-select:none;width:20px;display:inline-block}.tree-row{display:flex;align-items:center;padding:4px 0}.tree-row .icon{margin:0 5px}.tree-row .name{flex-grow:1}.tree-row .size{font-weight:700;min-width:120px;text-align:right;margin-right:10px}.percentage-bar-container{width:150px;height:16px;background-color:#e0e0e0;border-radius:4px;margin-right:10px}.percentage-bar{height:100%;background-color:var(--primary-color);border-radius:4px}.percentage{min-width:60px;text-align:right}</style></head><body><div class="container"><h1>File Analysis Report</h1><p><strong>Analyzed Path:</strong> $([System.Net.WebUtility]::HtmlEncode($ReportData.TargetPath))</p><p><strong>Report Date:</strong> $($ReportData.ReportDate.ToString('dd.MM.yyyy HH:mm:ss'))</p><h2>Summary</h2><div class="summary"><div class="card"><h3>Total Size</h3><p class="value">$(if ($null -ne $ReportData.DirectoryTree) { Format-FileSize $ReportData.DirectoryTree.TotalSize } else { 'N/A' })</p></div><div class="card"><h3>Total Files</h3><p class="value">$($ReportData.TotalFiles)</p></div><div class="card"><h3>Duplicate Sets</h3><p class="value">$($ReportData.Duplicates.Count)</p></div><div class="card"><h3>Wasted Space</h3><p class="value">$(Format-FileSize $ReportData.WastedSpace)</p></div></div><div class="tabs"><button class="tab-button active" onclick="openTab(event, 'Tree')">Directory Tree</button><button class="tab-button" onclick="openTab(event, 'Duplicates')">Duplicates</button><button class="tab-button" onclick="openTab(event, 'Size')">Files by Size</button><button class="tab-button" onclick="openTab(event, 'Age')">Files by Age</button></div><div id="Tree" class="tab-content active"><h3>Directory Size Tree</h3>$treeHtml</div><div id="Duplicates" class="tab-content"><h3>Duplicate Files</h3><p class="note">Method: $duplicateMethodNote</p>$duplicatesContent</div><div id="Size" class="tab-content"><h3>Files by Size</h3>$sizeContent</div><div id="Age" class="tab-content"><h3>Files by Age</h3>$ageContent</div></div><script>function openTab(e,t){var n,a,l;for(a=(n=document.getElementsByClassName("tab-content")).length,l=0;l<a;l++)n[l].style.display="none";for(a=(n=document.getElementsByClassName("tab-button")).length,l=0;l<a;l++)n[l].className=n[l].className.replace(" active","");document.getElementById(t).style.display="block",e.currentTarget.className+=" active"}document.querySelectorAll(".tree .toggle").forEach(e=>{e.addEventListener("click",function(){this.parentElement.parentElement.querySelector(".nested").classList.toggle("active"),this.innerHTML="&#9658;"===this.innerHTML?"&#9660;":"&#9658;"})});var style=document.createElement("style");style.innerHTML=".nested.active{display:block}",document.head.appendChild(style)</script></body></html>
"@
    $htmlTemplate | Out-File -FilePath $outputPath -Encoding UTF8
    Write-Log "Successfully generated HTML report: $outputPath" "Info"
    if ($pscmdlet.ShouldProcess($outputPath, "Opening HTML report")) { Invoke-Item -Path $outputPath }
}
#endregion

#region MAIN EXECUTION BLOCK
try {
    Write-Log "--- Starting File Analysis ---" "Info"
    
    $allFiles = $null
    if (-not $ForceScan -and (Test-Path $script:CacheFile)) {
        Write-Log "Found cache file. Importing data from $($script:CacheFile)." "Info"
        $allFiles = Import-Clixml -Path $script:CacheFile
    }
    else {
        $allFiles = Get-AllFiles -Path $TargetPath -ExcludePathsArray $resolvedExcludePaths
        Write-Log "Saving scan results to cache file: $($script:CacheFile)." "Verbose"
        $allFiles | Export-Clixml -Path $script:CacheFile
    }
    
    $filesByAge = [ordered]@{}
    foreach ($key in $AgeThresholds.Keys | Sort-Object { $AgeThresholds[$_] } -Descending) {
        $filesByAge[$key] = @($allFiles | Where-Object { $_.LastWriteTime -lt (Get-Date).AddYears(-$AgeThresholds[$key]) })
    }
    
    $filesBySize = [ordered]@{}
    foreach ($key in $SizeThresholds.Keys | Sort-Object { $SizeThresholds[$_] } -Descending) {
        $filesBySize[$key] = @($allFiles | Where-Object { $_.Length -gt $SizeThresholds[$key] })
    }
    
    $duplicates = Find-DuplicateFiles -Files $allFiles -UseHashCheck:$EnableHashCheck
    
    $wastedSpace = 0
    if ($duplicates.Count -gt 0) {
        $measureWasted = $duplicates | ForEach-Object { ($_.Group[0].Length * ($_.Count - 1)) } | Measure-Object -Sum
        if ($measureWasted) { $wastedSpace = $measureWasted.Sum }
    }
    
    Write-Log "Calculating directory sizes. This might take a while..." "Info"
    $directoryTree = Get-DirectorySizeTree -Path $TargetPath -ExcludePathsArray $resolvedExcludePaths
    
    if ($null -eq $directoryTree) {
        throw "Failed to generate the directory tree. This can be caused by permissions issues on the root folder '$TargetPath'."
    }

    $reportData = [pscustomobject]@{
        TargetPath      = $TargetPath
        ReportDate      = Get-Date
        TotalFiles      = $allFiles.Count
        FilesByAge      = $filesByAge
        FilesBySize     = $filesBySize
        Duplicates      = $duplicates
        WastedSpace     = $wastedSpace
        DirectoryTree   = $directoryTree
        HashCheckEnabled = $EnableHashCheck.IsPresent
    }
    
    New-HTMLReport -ReportData $reportData
    
    if ($ExportFormats) {
        Export-AnalysisData -ReportData $reportData -Formats $ExportFormats -OutDir $OutputDirectory
    }
    
    Write-Log "--- Analysis Completed Successfully ---" "Info"
}
catch {
    Write-Log "A critical error occurred during analysis: $($_.Exception.Message)" "Error"
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" "Verbose"
}
finally {
    Write-Host "`nAnalysis finished. Check the log file for details: $script:LogFile"
}
#endregion
