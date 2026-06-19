}
catch {
    Write-Error "An error occurred: $_"
    Read-Host "Press Enter to exit"
    exit 1
}
Starting enhanced AD user report generation...
Processing 53420 user-group combinations...
CSV export completed: C:\daten\AD_Benutzer_Gruppen_L.csv
Excel export completed: C:\daten\AD_Benutzer_Gruppen_L.xlsx
HTML visualization exported: C:\daten\AD_Benutzer_Gruppen_L.html
WARNUNG: Visio export failed: 

Objektname nicht gefunden.
WARNUNG: MindManager export failed: Fehler beim Aufrufen der Methode, da [System.Management.Automation.PSParameter
izedProperty] keine Methode mit dem Namen "Add" enthält.
An error occurred: Fehler beim Aufrufen der Methode, da [System.__ComObject] keine Methode mit dem Namen "Quit" 
enthält.
    + CategoryInfo          : NotSpecified: (:) [Write-Error], WriteErrorException
    + FullyQualifiedErrorId : Microsoft.PowerShell.Commands.WriteErrorException
 
Press Enter to exit: 


[CmdletBinding()]
param (
    [Parameter()]
    [string]$OutputPath = "C:\daten\AD_Benutzer_Gruppen_L.csv",
    [Parameter()]
    [string]$ExcelPath = "C:\daten\AD_Benutzer_Gruppen_L.xlsx",
    [Parameter()]
    [string]$HtmlPath = "C:\daten\AD_Benutzer_Gruppen_L.html",
    [Parameter()]
    [string]$VisioPath = "C:\daten\AD_Benutzer_Gruppen_L.vsdx",
    [Parameter()]
    [string]$MindManagerPath = "C:\daten\AD_Benutzer_Gruppen_L.mmap"
)

function Close-ExcelProcesses {
    Get-Process -Name "excel" -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            $_.CloseMainWindow() | Out-Null
            Start-Sleep -Milliseconds 500
            if (!$_.HasExited) { $_.Kill() }
        } catch {
            Write-Warning "Could not close Excel process: $_"
        }
    }
    Start-Sleep -Seconds 2
}

function Get-ADUserGroupData {
    param ([string]$SearchPattern = "L*")
    
    $users = Get-ADUser -Filter "SamAccountName -like '$SearchPattern'" -Properties SamAccountName, Name, MemberOf, DistinguishedName, Comment
    
    if (-not $users) {
        Write-Warning "No users found with SamAccountName starting with 'L'"
        return $null
    }
    
    $userData = @()
    $groupColors = @{}
    $colorIndex = 35
    
    foreach ($user in $users) {
        Write-Verbose "Processing user: $($user.SamAccountName)"
        
        $ouMatch = $user.DistinguishedName -match 'OU=([^,]+)'
        $ou = if ($ouMatch) { $Matches[1] } else { "No OU" }
        $numericPrefix = if ($ou -match '^\d{2}$') { $Matches[0] } else { "999" }
        
        $groups = $user.MemberOf | ForEach-Object {
            try {
                (Get-ADGroup $_).Name
            } catch {
                Write-Warning "Could not resolve group for user $($user.SamAccountName): $_"
                return "Unknown Group"
            }
        } | Sort-Object
        
        foreach ($group in $groups) {
            if (-not $groupColors.ContainsKey($group)) {
                $groupColors[$group] = $colorIndex
                $colorIndex++
                if ($colorIndex -gt 46) { $colorIndex = 35 }
            }
            
            $userData += [PSCustomObject]@{
                SortPrefix = $numericPrefix
                OU = $ou
                UserName = $user.Name
                SamAccountName = $user.SamAccountName
                Group = $group
                Comment = $user.Comment
                ColorIndex = $groupColors[$group]
            }
        }
    }
    
    return @{
        Data = $userData
        Colors = $groupColors
    }
}

function Export-ToCSV {
    param (
        [Parameter(Mandatory)][object[]]$Data,
        [Parameter(Mandatory)][string]$Path
    )
    
    $directory = Split-Path -Parent $Path
    if (-not (Test-Path $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
    
    if (Test-Path $Path) {
        Remove-Item $Path -Force -ErrorAction Stop
    }
    
    $Data | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Force
    Write-Host "CSV export completed: $Path" -ForegroundColor Green
}

function Export-ToExcel {
    param (
        [Parameter(Mandatory)][object[]]$Data,
        [Parameter(Mandatory)][hashtable]$GroupColors,
        [Parameter(Mandatory)][string]$Path
    )
    
    $excel = $null
    
    try {
        $directory = Split-Path -Parent $Path
        if (-not (Test-Path $directory)) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
        }
        
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        
        $headers = @("OU", "Benutzer", "SamAccountName", "Gruppe", "Kommentar")
        1..5 | ForEach-Object { 
            $worksheet.Cells.Item(1, $_) = $headers[$_ - 1]
        }
        
        $row = 2
        foreach ($item in $Data) {
            $worksheet.Cells.Item($row, 1) = $item.OU
            $worksheet.Cells.Item($row, 2) = $item.UserName
            $worksheet.Cells.Item($row, 3) = $item.SamAccountName
            $worksheet.Cells.Item($row, 4) = $item.Group
            $worksheet.Cells.Item($row, 5) = $item.Comment
            
            $groupCell = $worksheet.Cells.Item($row, 4)
            $groupCell.Interior.ColorIndex = $GroupColors[$item.Group]
            
            $row++
        }
        
        $headerRange = $worksheet.Range($worksheet.Cells(1, 1), $worksheet.Cells(1, 5))
        $headerRange.Font.Bold = $true
        $headerRange.Interior.ColorIndex = 15
        
        $worksheet.Range($worksheet.Cells(1, 1), $worksheet.Cells($row - 1, 5)).AutoFilter() | Out-Null
        $worksheet.Columns.Item(1).ColumnWidth = 20
        $worksheet.Columns.Item(2).ColumnWidth = 30
        $worksheet.Columns.Item(3).ColumnWidth = 20
        $worksheet.Columns.Item(4).ColumnWidth = 50
        $worksheet.Columns.Item(5).ColumnWidth = 40
        
        if (Test-Path $Path) {
            Remove-Item $Path -Force
        }
        
        $workbook.SaveAs($Path)
        $workbook.Close($true)
        
        Write-Host "Excel export completed: $Path" -ForegroundColor Green
    }
    finally {
        if ($excel) {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
}

function Get-HtmlTemplate {
    return @"
<!DOCTYPE html>
<html>
<head>
    <title>AD User Groups Visualization</title>
    <script src="https://d3js.org/d3.v7.min.js"></script>
    <style>
        body { font-family: Arial, sans-serif; margin: 0; padding: 20px; }
        .node circle { fill: #fff; stroke: steelblue; stroke-width: 1.5px; }
        .node text { font: 12px sans-serif; }
        .link { fill: none; stroke: #ccc; stroke-width: 1.5px; }
        #visualization { width: 100%; height: 800px; border: 1px solid #ccc; }
    </style>
</head>
<body>
    <h1>AD User Groups Visualization</h1>
    <div id="visualization"></div>
    <script>
__JS_CONTENT__
    </script>
</body>
</html>
"@
}

function Get-VisualizationScript {
    return @'
const data = {
            nodes: [],
            links: []
        };
        
        // Process data for visualization
        const users = DATA_PLACEHOLDER;
        const processedUsers = new Set();
        const processedGroups = new Set();
        
        users.forEach(user => {
            if (!processedUsers.has(user.SamAccountName)) {
                data.nodes.push({
                    id: user.SamAccountName,
                    type: 'user',
                    name: user.UserName,
                    ou: user.OU
                });
                processedUsers.add(user.SamAccountName);
            }
            
            if (!processedGroups.has(user.Group)) {
                data.nodes.push({
                    id: user.Group,
                    type: 'group',
                    name: user.Group
                });
                processedGroups.add(user.Group);
            }
            
            data.links.push({
                source: user.SamAccountName,
                target: user.Group
            });
        });
        
        // Create force-directed graph
        const width = window.innerWidth - 40;
        const height = 800;
        
        const simulation = d3.forceSimulation(data.nodes)
            .force("link", d3.forceLink(data.links).id(d => d.id))
            .force("charge", d3.forceManyBody().strength(-300))
            .force("center", d3.forceCenter(width / 2, height / 2));
        
        const svg = d3.select("#visualization")
            .append("svg")
            .attr("width", width)
            .attr("height", height);
        
        const link = svg.append("g")
            .selectAll("line")
            .data(data.links)
            .join("line")
            .attr("class", "link");
        
        const node = svg.append("g")
            .selectAll("g")
            .data(data.nodes)
            .join("g")
            .attr("class", "node")
            .call(d3.drag()
                .on("start", dragstarted)
                .on("drag", dragged)
                .on("end", dragended));
        
        node.append("circle")
            .attr("r", d => d.type === 'user' ? 5 : 8)
            .style("fill", d => d.type === 'user' ? "#69b3a2" : "#ff7f50");
        
        node.append("text")
            .text(d => d.name)
            .attr("x", 8)
            .attr("y", 3);
        
        simulation.on("tick", () => {
            link
                .attr("x1", d => d.source.x)
                .attr("y1", d => d.source.y)
                .attr("x2", d => d.target.x)
                .attr("y2", d => d.target.y);
            
            node
                .attr("transform", d => `translate(${d.x},${d.y})`);
        });
        
        function dragstarted(event, d) {
            if (!event.active) simulation.alphaTarget(0.3).restart();
            d.fx = d.x;
            d.fy = d.y;
        }
        
        function dragged(event, d) {
            d.fx = event.x;
            d.fy = event.y;
        }
        
        function dragended(event, d) {
            if (!event.active) simulation.alphaTarget(0);
            d.fx = null;
            d.fy = null;
        }
'@
}

function Export-ToHtml {
    param (
        [Parameter(Mandatory)][object[]]$Data,
        [Parameter(Mandatory)][string]$Path
    )

    $jsonData = $($Data | ConvertTo-Json)

    $jsScript = Get-VisualizationScript
    $jsScript = $jsScript.Replace('DATA_PLACEHOLDER', $jsonData)

    $htmlTemplate = Get-HtmlTemplate
    $html = $htmlTemplate.Replace('__JS_CONTENT__', $jsScript)
    
    $html | Out-File -FilePath $Path -Encoding UTF8
    Write-Host "HTML visualization exported: $Path" -ForegroundColor Green
}

function Export-ToVisio {
    param (
        [Parameter(Mandatory)][object[]]$Data,
        [Parameter(Mandatory)][string]$Path
    )
    
    try {
        $visio = New-Object -ComObject Visio.Application
        $visio.Visible = $false
        
        $doc = $visio.Documents.Add("")
        $page = $doc.Pages.Item(1)
        
        # Create stencil for shapes
        $stencil = $visio.Documents.OpenEx("ORGCH_M.VSS", 4)
        $userShape = $stencil.Masters.Item("Process")
        $groupShape = $stencil.Masters.Item("Decision")
        
        $shapes = @{}
        $yPos = 1
        $xPos = 1
        
        # Create user shapes
        $Data | Sort-Object OU, UserName -Unique | ForEach-Object {
            if (-not $shapes.ContainsKey($_.SamAccountName)) {
                $shape = $page.Drop($userShape, $xPos * 2, 10 - $yPos)
                $shape.Text = "$($_.UserName)`n($($_.SamAccountName))"
                $shapes[$_.SamAccountName] = $shape
                $yPos++
                if ($yPos -gt 8) {
                    $yPos = 1
                    $xPos++
                }
            }
        }
        
        # Create group shapes and connections
        $yPos = 1
        $xPos += 2
        $Data | Sort-Object Group -Unique | ForEach-Object {
            if (-not $shapes.ContainsKey($_.Group)) {
                $shape = $page.Drop($groupShape, $xPos * 2, 10 - $yPos)
                $shape.Text = $_.Group
                $shapes[$_.Group] = $shape
                $yPos++
                if ($yPos -gt 8) {
                    $yPos = 1
                    $xPos++
                }
            }
        }
        
        # Add connections
        $Data | ForEach-Object {
            $page.Shapes.AddConnector(1, $shapes[$_.SamAccountName], $shapes[$_.Group]) | Out-Null
        }
        
        $doc.SaveAs($Path)
        $visio.Quit()
        
        Write-Host "Visio visualization exported: $Path" -ForegroundColor Green
    }
    catch {
        Write-Warning "Visio export failed: $_"
    }
    finally {
        if ($visio) {
            $visio.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($visio) | Out-Null
        }
    }
}

function Export-ToMindManager {
    param (
        [Parameter(Mandatory)][object[]]$Data,
        [Parameter(Mandatory)][string]$Path
    )
    
    try {
        $mm = New-Object -ComObject MindManager.Application
        $mm.Visible = $false
        
        $doc = $mm.Documents.Add()
        $root = $doc.CentralTopic
        $root.Text = "AD User Groups"
        
        # Group by OU
        $Data | Group-Object OU | Sort-Object {
            if ($_.Name -match '^\d{2}') { 
                [int]($_.Name -replace '^(\d{2}).*$','$1')
            } else { 
                999 
            }
        } | ForEach-Object {
            $ouTopic = $root.AddSubTopic()
            $ouTopic.Text = $_.Name
            
            # Add users under OU
            $_.Group | Group-Object SamAccountName | ForEach-Object {
                $userTopic = $ouTopic.AddSubTopic()
                $userTopic.Text = "$($_.Group[0].UserName)`n($($_.Name))"
                
                # Add groups under user
                $_.Group | ForEach-Object {
                    $groupTopic = $userTopic.AddSubTopic()
                    $groupTopic.Text = $_.Group
                }
            }
        }
        
        $doc.SaveAs($Path)
        $mm.Quit()
        
        Write-Host "MindManager visualization exported: $Path" -ForegroundColor Green
    }
    catch {
        Write-Warning "MindManager export failed: $_"
    }
    finally {
        if ($mm) {
            $mm.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($mm) | Out-Null
        }
    }
}

try {
    Write-Host "Starting enhanced AD user report generation..." -ForegroundColor Green
    
    Close-ExcelProcesses
    
    $result = Get-ADUserGroupData
    if (-not $result) { exit 0 }
    
    $sortedData = $result.Data | Sort-Object {
        if ($_.SortPrefix -match '^\d{2}$') { 
            [int]$_.SortPrefix 
        } else { 
            999 
        }
    }, UserName, Group | Select-Object OU, UserName, SamAccountName, Group, Comment
    
    Write-Host "Processing $($sortedData.Count) user-group combinations..." -ForegroundColor Green
    
    # Export all formats
    Export-ToCSV -Data $sortedData -Path $OutputPath
    Export-ToExcel -Data $sortedData -GroupColors $result.Colors -Path $ExcelPath
    Export-ToHtml -Data $sortedData -Path $HtmlPath
    Export-ToVisio -Data $sortedData -Path $VisioPath
    Export-ToMindManager -Data $sortedData -Path $MindManagerPath
    
    Write-Host "Script completed successfully." -ForegroundColor Green
    Read-Host "Press Enter to exit"
}
catch {
    Write-Error "An error occurred: $_"
    Read-Host "Press Enter to exit"
    exit 1
}
