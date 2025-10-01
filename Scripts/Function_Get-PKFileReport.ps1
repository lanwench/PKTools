#requires -Version 4

Function Get-PKFileReport {
<#
.SYNOPSIS
    Generates an HTML report of files in a specified directory, including summary statistics and detailed file information.

.DESCRIPTION
    Scans a given folder (recursively, unless specified otherwise) and produces a styled HTML report containing:
        - A summary of file types and their counts
        - Detailed information for each file, such as name, type, full path, creation and modification dates, and optionally, the owner
        - Interactive, sortable tables with custom branding and color schemes
    Launches the report in the default web browser unless specified otherwise

.NOTES
    Name    : Function_Get-PKFileReport.ps1
    Author  : Paula Kingsley
    Created : 2025-09-30
    Version : 01.00.0000 
    History :

        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v01.00.0000 - 2025-09-30 - Created script

.PARAMETER Path
    The path to the folder or file to report on. Accepts pipeline input. Defaults to the current directory. Validates that the path exists.

.PARAMETER GetOwner
    If specified, includes file owner and editor information in the report.

.PARAMETER NoRecurse
    If specified, does not search subdirectories (non-recursive).

.PARAMETER NoLaunch
    If specified, will not prompt the user to launch the generated HTML report in the default browser after creation.

.PARAMETER ReportPath
    The directory where the HTML report will be saved. Defaults to user's temp directory

.INPUTS
    System.String, System.IO.FileSystemInfo

.OUTPUTS
    HTML file, with optional launch in default browser

.EXAMPLE
    PS C:\Projects> Get-PKFileReport -GetOwner -Verbose

        Generates a file report for the "C:\Projects" directory, including owner information, and outputs verbose progress.

.EXAMPLE
    PS C:\> Get-PKFileReport -Path "G:\Shared Drives\TeamX\Data" -NoRecurse -NoLaunch -ReportPath "C:\Temp\Reports"

        Generates a non-recursive file report for "G:\Shared Drives\TeamX\Data" and saves it to "C:\Temp\Reports", without launching the report after creation.

#>

[CmdletBinding(ConfirmImpact='Low',SupportsShouldProcess=$true)]
Param (
    [Parameter(
        HelpMessage = "Path for scan",
        Position = 0,    
        ValueFromPipeline,
        ValueFromPipelineByPropertyName
    )]
    [Alias("Name","Drive","Folder")]
    [ValidateScript( {
        If (Test-Path $_) { $true } Else { Throw "Invalid path '$_'" }
    })]
    [object]$Path = $PWD,

    [Parameter(
        HelpMessage = "Get the file owner, where available"
    )]
    [switch]$GetOwner,
    
    [Parameter(
        HelpMessage = "Don't recurse into subdirectories"
    )]
    [switch]$NoRecurse,
    
    [Parameter(
        HelpMessage = "Don't prompt to launch the output report in the default browser"
    )]
    [switch]$NoLaunch,
    
    [Parameter(
        HelpMessage = "Path for the output report HTML file (defaults to user's temp directo)"
    )]
    [string]$ReportPath = $Env:Temp
)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    # Show our settings
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object {$CurrentParams.keys -notContains $_} | 
        Where-Object {Test-Path variable:$_}| Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"
    
    #region Validate input; if it's already a file object, do nothing
    If (-not $PipelineInput -and ($Path -is [string])) {
        $Path = Get-Item -Path $Path -ErrorAction Stop
    }
    #endregion Validate input

    #region Set variables
    $ComputerName = $Env:COMPUTERNAME
    $Username = $Env:USERNAME
    $OutFile = "$ReportPath\FolderReport_$(Get-Date -f yyyy-MM-dd_HH-mm-ss).html"
    $reportTitle = "$ComputerName - File report for $Path"
    $ReportDate = (Get-Date -Format F)

    # Create an arraylist for the contents instead of using the icky += 
    $Output = [System.Collections.ArrayList]::new()
    
    #endregion Set variables

    #region Nielsen colors

    # Nielsen brand colors - avoiding #6E37FA
    $CompanyColors = @(
        "#32BBB9", "#FF9408", "#050404ff", "#FA32A0", "#B30095",
        "#FFD500", "#AAF564", "#50E6AA", "#2765F0", "#005F81",
        "#C35000", "#A00032", "#B40073", "#64005A", "#D29100"
    )

    #endregion Nielsen colors

    #region here strings for CSS and JavaScript - do not indent!!
$preContent = @"
<div class="header-section">
    <h1>$reportTitle</h1>
    <h2 class="section-title">Report details</h2>
    <div class="report-meta">Report <strong>$OutFile</strong> generated on <strong>$ReportDate</strong> by <strong>$Username</strong></div>
</div>
"@

$css = @"
<style>
    body { 
        font-family: Segoe UI, Arial, sans-serif; 
        background: #f8f9fa;
        margin: 20px; 
    }
    
    /* Common section styling */
    .header-section, .summary-section {
        margin-bottom: 20px;
        padding: 15px;
        background: white;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        width: calc(100% - 30px); /* Full width minus padding */
        max-width: 1200px; /* Optional: prevent extremely wide displays */
    }
    
    /* Typography */
    h1 {
        margin-top: 0;
        margin-bottom: 15px;
        color: #6E37FA;
        font-size: 30px;
        font-weight: normal;
    }
    h2.section-title {
        margin-top: 0;
        margin-bottom: 15px;
        color: #FA32A0;
        font-size: 22px;
    }
    .report-meta {
        margin-top: 8px;
        margin-bottom: 5px;
        color: #333;
        line-height: 1.5;
    }
    .file-count {
        font-size: 1.1em;
        margin-bottom: 15px;
    }
    
    /* Tables */
    table { 
        border-collapse: collapse; 
        width: auto; 
        margin: 20px 0; 
        table-layout: auto; 
    }
    th, td { 
        border: 1px solid #ddd; 
        padding: 8px;  
    }
    th { 
        background: #6E37FA; 
        color: #fff; 
        cursor: pointer; 
        user-select: none; 
        position: relative; 
        font-size: 16px;
        padding: 10px 8px;
        font-weight: bold;
    }
    tr:nth-child(even) { background: #f2f2f2; }
    tr:hover { background: #e6f7ff; }
    
    /* Summary table specific */
    .summary-table {
        width: auto;
        border-collapse: collapse;
        margin-bottom: 10px;
    }
    .summary-table th {
        text-align: left;
        padding: 8px;
        background-color: #f2f2f2;
        color: #333;
    }
    .summary-table td {
        padding: 6px 8px;
        border-bottom: 1px solid #eee;
    }
    .color-box {
        display: inline-block;
        width: 12px;
        height: 12px;
        margin-right: 5px;
        border-radius: 2px;
    }
    
    /* Sort indicators */
    th::after { 
        content: ' ▼'; 
        color: #999999;
        opacity: 0.7;
        font-size: 12px;
        margin-left: 5px;
    }
    th.sort-asc::after { 
        content: ' ▲'; 
        color: #ffff99;
        opacity: 1;
        font-weight: bold;
        margin-left: 5px;
    }
    th.sort-desc::after { 
        content: ' ▼'; 
        color: #ffff99;
        opacity: 1;
        font-weight: bold;
        margin-left: 5px;
    }
    th.sort-asc, th.sort-desc { 
        background: #5526dc; 
        font-weight: bold;
    }
</style>
"@

$script = @"
<script>
    document.addEventListener('DOMContentLoaded', () => {
        const sortStates = {};
        
        document.querySelectorAll('th').forEach((th, idx) => {
            sortStates[idx] = null; // No initial sort
            
            th.addEventListener('click', () => {
                const table = th.closest('table');
                const tbody = table.querySelector('tbody');
                const rows = Array.from(tbody.rows);
                
                // Clear all other column indicators
                document.querySelectorAll('th').forEach(header => {
                    header.classList.remove('sort-asc', 'sort-desc');
                });
                
                // Determine new sort direction
                let newDirection;
                if (sortStates[idx] === null || sortStates[idx] === 'desc') {
                    newDirection = 'asc';
                } else {
                    newDirection = 'desc';
                }
                
                // Apply visual indicator to clicked column
                th.classList.add('sort-' + newDirection);
                
                // Sort the rows
                rows.sort((a, b) => {
                    const va = a.cells[idx].textContent.trim();
                    const vb = b.cells[idx].textContent.trim();
                    
                    const result = va.localeCompare(vb, undefined, {numeric: true, sensitivity: 'base'});
                    return newDirection === 'asc' ? result : -result;
                });
                
                // Re-append sorted rows
                tbody.innerHTML = '';
                rows.forEach(row => tbody.appendChild(row));
                
                // Update sort state
                sortStates[idx] = newDirection;
                
                // Reset other columns' sort states
                Object.keys(sortStates).forEach(key => {
                    if (key != idx) sortStates[key] = null;
                });
            });
        });
    });
</script>
"@

    #endregion here strings for CSS and JavaScript - do not indent!!

    #region Select properties based on output
    If ($GetOwner) {
        $Select = "Name","FileType","FullName","Owner","Editor","Created","Changed"
    } 
    Else {
        $Select = "Name","FileType","FullName","Created","Changed"
    }
    #endregion Select properties based on output

    #region Inner function to get file type description
    Function _GetType {
    [CmdletBinding()]
    Param([Parameter(Position=0,ValueFromPipeline,Mandatory)][object]$InputFile)
    Begin {}
    Process {
        Try {
            $shell = New-Object -ComObject Shell.Application
            $folder = $shell.Namespace($InputFile.DirectoryName)
            if (-not $folder) {
                Write-Verbose "[$($InputFile.FullName)] Folder object is null"
                "ERROR"
                return
            }
            $file = $folder.ParseName($InputFile.Name)
            if (-not $file) {
                "ERROR"
                return
            }
            $FileType = $folder.GetDetailsOf($file, 2)
            Write-Output $FileType
        }
        Catch {
            "ERROR"
        }
    }
    End {}
}

    #endregion Inner function to get file type description

    Write-Verbose "[BEGIN: $ScriptName] Starting file report scan in $Path"
}
Process {
    Try {
        Get-ChildItem -Path $Path -Recurse:(-not $NoRecurse) -File -ErrorAction Stop -Verbose:$False | ForEach-Object {

            $CurrentFile = $_
            Write-Verbose "[Processing] $($CurrentFile.FullName)"
            
            If ($GetOwner) {
                Try {
                    $Owner = ((Get-Acl $CurrentFile.FullName -ErrorAction Stop).Owner) ?? "Unknown"
                }
                Catch {
                    Write-Warning $_.Exception.Message
                    $Owner = "ERROR"
                }
            } # end if getting owner
            Try {
                If ($CurrentFile.Extension) {
                    $Type = (_GetType -InputFile $CurrentFile) ?? ($CurrentFile.Extension.TrimStart('.'))
                }
                Else {$Type = "(no extension)"}
                
                $FileInfo = [PSCustomObject]@{
                    Parent    = $CurrentFile.DirectoryName
                    Name      = $CurrentFile.Name
                    FileType  = $Type
                    FullName  = $CurrentFile.FullName
                    Owner     = If ($GetOwner) { $owner } else { $null }
                    Created   = Get-Date $CurrentFile.CreationTime -format u
                    Changed   = Get-Date $CurrentFile.LastWriteTime -format u 
                }
                $Output.Add($FileInfo) | Out-Null
            } # end try for individual file details
            Catch {
                $FileInfo = [PSCustomObject]@{
                    Parent    = $CurrentFile.DirectoryName
                    Name      = $CurrentFile.Name
                    FileType  = "ERROR"
                    FullName  = $CurrentFile.FullName
                    Owner     = "ERROR"
                    Created   = "ERROR"
                    Changed   = "ERROR"
                }
                $Output.Add($FileInfo) | Out-Null

            } # end catch for individual file

        }  # end for each object

    } # end try for GCI
    Catch {
        Write-Warning $_.Exception.Message
        $FileInfo = [PSCustomObject]@{
            Parent    = $Path
            Name      = "ERROR"
            FileType  = "ERROR"
            FullName  = "ERROR"
            Owner     = "ERROR"
            Created   = "ERROR"
            Changed   = "ERROR"
        }
    }
} # end process
End {

    If ($Output.Count -eq 0) {
        Write-Verbose "No results found in $Path"
    }
    Else {    

        Write-Verbose "Creating HTML content"

        # Calculate summary data
        $FileCount = $Output.Count
        $FileTypeGroups = $Output | Group-Object -Property FileType | Sort-Object -Property Count -Descending
        
        # Build summary table rows (limit to top 15 file types)
        $TopFileTypes = $FileTypeGroups | Select-Object -First 15
        $OtherFilesCount = ($FileTypeGroups | Select-Object -Skip 15 | Measure-Object -Property Count -Sum).Sum

        $SummaryTableRows = ""
        For ($i = 0; $i -lt $TopFileTypes.Count; $i++) {
            $ColorIndex = $i % $CompanyColors.Count
            $Color = $CompanyColors[$ColorIndex]
            $percentage = [math]::Round(($TopFileTypes[$i].Count / $FileCount) * 100, 1)
            $SummaryTableRows += "<tr><td><span class='color-box' style='background-color:$Color'></span> $($TopFileTypes[$i].Name)</td><td>$($TopFileTypes[$i].Count)</td><td>$percentage%</td></tr>`n"
        }
        
        # Add "Other" row if needed
        if ($OtherFilesCount -gt 0) {
            $OtherPercentage = [math]::Round(($OtherFilesCount / $FileCount) * 100, 1)
            $SummaryTableRows += "<tr><td><span class='color-box' style='background-color:#777F9E'></span> Other file types</td><td>$OtherFilesCount</td><td>$OtherPercentage%</td></tr>`n"
        }
        
        # Convert content to HTML fragment (just the table)
        $TableFragment = $Output | Select-Object $Select | ConvertTo-Html -Fragment -Property $Select -ErrorAction Stop

        # Explicitly structure the table with thead and tbody
        $TableFragment = $TableFragment -replace '<table>', '<table>'
        $TableFragment = $TableFragment -replace '<tr><th(.*?)</tr>', '<thead><tr><th$1</tr></thead><tbody>'
        $TableFragment = $TableFragment -replace '</table>', '</tbody></table>'

        # Create summary section with simple table
        $SummarySection = @"
<div class="summary-section">
    <h2 class="section-title">Summary of files</h2>
    <div class="file-count">Total Files: $FileCount</div>
    <table class="summary-table">
        <thead>
            <tr>
                <th>File Type</th>
                <th>Count</th>
                <th>Percentage</th>
            </tr>
        </thead>
        <tbody>
            $summaryTableRows
        </tbody>
    </table>
</div>
<h2 class="section-title">File details</h2>
"@

        # Build the complete HTML document
        $HTMLDocument = @"
<!-- saved from url=(0014)about:internet -->
<!DOCTYPE html>
<html>
<head>
    <title>$reportTitle</title>
    $css
    $script
</head>
<body>
$preContent
$SummarySection
$TableFragment
</body>
</html>
"@

        # Write to the file
        Write-Verbose "Creating output file"
        $HTMLDocument | Set-Content "$OutFile" -Confirm:$False -Force -ErrorAction Stop

        # Launch it
        If (-not $NoLaunch) {
            Write-Verbose "Prompting to launch output file in default browser"
            If ($PSCmdlet.ShouldContinue("Launch output file in default browser?",$OutFile)) {
                Invoke-Item "$OutFile" -Confirm:$False -ErrorAction SilentlyContinue
            }
        }
        Else {
            Write-Verbose "Skipping prompt to launch output file '$OutFile'"
            Write-Output "Output file successfully created as '$OutFile'"
        }   
    }

    # We're done here
    $Stopwatch.Stop()
    Write-Verbose "[END: $ScriptName] Processing completed in $($Stopwatch.Elapsed.TotalSeconds) seconds."
}
} # end function

