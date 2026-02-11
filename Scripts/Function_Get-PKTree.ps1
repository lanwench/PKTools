#Requires -version 4
Function Get-PKTree {
<#
.SYNOPSIS 
    Invokes Get-ChildItem -Recurse to return files, folders (or both) on the current or other directory, with options for depth, string matching, and output type

.DESCRIPTION
    Invokes Get-ChildItem -Recurse to return files, folders (or both) on the current or other directory, with options for depth, string matching, and output type
    Returns a System.IO.FileSystemInfo object or a string

.NOTES
    Name    : Function_Get-PKTree.ps1
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :
    
        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v01.00.0000 - 2023-03-02 - Created script

.PARAMETER Path
    Starting path for search (default is current directory)

.PARAMETER Type
    Return only files, only folders, or both (default is both)

.PARAMETER Match
    HelpMessage = "One or more strings to search for (e.g., read*.md or '*.csv','*.png' ... default is *)"    

.PARAMETER Depth
    Depth to search (default is unlimited)

.PARAMETER Force
    Use -Force switch

.PARAMETER ReturnNamesOnly
    Return full names (strings) instead of file or folder objects

.EXAMPLE
    PS C:\Users\jbloggs\repos\> Get-PKTree -Match *.psd1,*.md -Verbose
        VERBOSE: PSBoundParameters: 
	
        Key             Value                            
        ---             -----                            
        Path            C:\Users\jbloggs\repos
        Match           {*.psd1, *.md}                   
        Verbose         True                             
        Type            All                              
        Depth                                            
        Force           False
        PipelineInput   False                            
        ReturnNamesOnly False                            
        ScriptName      Get-PKTree                       
        ScriptVersion   1.0.0                            

        VERBOSE: [BEGIN: Get-PKTree] Return file and directory object(s) matching '*.psd1', '*.md' (unlimited depth) 
        VERBOSE: [C:\Users\jbloggs\repos] Searching...
        VERBOSE: [C:\Users\jbloggs\repos] 5 match(es) found

            Directory: C:\Users\jbloggs\repos\mytools

        Mode                 LastWriteTime         Length Name                                                                                                                                         
        ----                 -------------         ------ ----                                                                                                                                         
        -a----        11/17/2022   9:51 AM           4049 mytools.psd1                                                                                                                    
        -a----        11/17/2022   9:51 AM          10288 README.md                                                                                                                                    

            Directory: C:\Users\jbloggs\repos\vmstuff

        Mode                 LastWriteTime         Length Name                                                                                                                                         
        ----                 -------------         ------ ----                                                                                                                                         
        -a----        11/17/2022   9:51 AM           3971 vmstuff.psd1                                                                                                                             
        -a----        11/17/2022   9:51 AM           2810 README.md                                                                                                                                    

            Directory: C:\Users\jbloggs\repos\sandbox

        Mode                 LastWriteTime         Length Name                                                                                                                                         
        ----                 -------------         ------ ----                                                                                                                                         
        -a----        11/17/2022   9:53 AM            118 README.md                                                                                                                                    
                                                                                                        
        VERBOSE: [END: Get-PKTree]

.EXAMPLE
    PS C:\> Get-PKTree -Path c:\temp -Type Files -ReturnNamesOnly
        
        C:\temp\Reports\ADUsers\allusers.csv
        C:\temp\Reports\ADUsers\adlist.txt
        C:\temp\Reports\testing-testing.xml
        C:\temp\Groups.csv
        C:\temp\resume.docx

.EXAMPLE
    PS C:\> "c:\temp","c:\windows\temp" | Get-PKTree -Depth 0 -Type Folders -ReturnNamesOnly

        C:\temp\Music-temp
        C:\temp\Reports
        C:\temp\Reports\ADUsers
        C:\windows\temp\Crashpad
        C:\windows\temp\Crashpad\attachments
        C:\windows\temp\Crashpad\reports
        C:\windows\temp\DA22FD9F-AB8C-4985-A32E-A2BA4E8F5E9A-Sigs
        C:\windows\temp\data
        C:\windows\temp\Deployment
        C:\windows\temp\gen_py
        C:\windows\temp\gen_py\3.8
        C:\windows\temp\PPP
        C:\windows\temp\PPP\Log
        ...


#>
[cmdletbinding()]
Param (
[Parameter(
    Position = 0,
    ValueFromPipeline,
    ValueFromPipelineByPropertyName,
    HelpMessage = "Starting path for search (default is current directory)"
)]
[Alias("Name","FullName")]
[string]$Path = $PWD,

[Parameter(
    HelpMessage = "Return only files, only folders, or both (default is both)"
)]
[ValidateSet("Files","Folders","All")]
[String]$Type = "All",

[Parameter(
    HelpMessage = "One or more strings to search for (e.g., read*.md or '*.csv','*.png' ... default is *)"    
)]
[string[]]$Match ='*',

[Parameter(
    HelpMessage = "Depth to search (default is unlimited)"    
)]
[System.Nullable[int]]$Depth,

[Parameter(
    HelpMessage = "Use -Force switch"
)]
[switch]$Force,

[Parameter(
    HelpMessage = "Return full names (strings) instead of file or folder objects"
)]
[switch]$ReturnNamesOnly
) 
Begin {
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here?
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    $CurrentParams = $PSBoundParameters
    $ScriptName = $MyInvocation.MyCommand.Name
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path Variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    #region Parameters and activity 
    $Param = @{
        Recurse     = $True
        Include     = $Match
        ErrorAction = "Stop"
        Verbose     = $False
    }
    If ($CurrentParams.Depth -ne $Null) {
        $DepthStr = "to level $($Depth.ToString())"
        $Param.Add("Depth",$Depth)
    }
    Else {
        $DepthStr = "(unlimited depth)"
    }
    
    Switch ($Type) {
        All {$TypeStr = "file and directory"}
        Folders {
            $TypeStr = "directory"
            $Param.Add("Directory",$True)        
        }
        Files {
            $TypeStr = "file"
            $Param.Add("File",$True)
        }
    }
    If ($ReturnNamesOnly.IsPresent) {
        $ReturnStr = "Return full name of $TypeStr object(s)"
    }
    Else {
        $ReturnStr = "Return $TypeStr object(s)"
    }
    $SearchStr = "$($Match -join("', '"))"
    $Activity = "$ReturnStr matching '$SearchStr' $DepthStr"

    
    If ($Force.IsPresent) {
        $Activity += "; -Force specified"
        $Param.Add("Force",$True)
    }
    
    #endregion Parameters and activity 

    Write-Verbose "[BEGIN: $ScriptName] $Activity "
}
Process {

    Foreach ($P in $Path) {

        Write-Verbose "[$P] Searching..."
        Try {
        
            If ($Output = Get-ChildItem @Param -Path $P) {
                Write-Verbose "[$P] $($Output.Count) match(es) found"
                If ($ReturnNamesOnly.IsPresent) {
                    Write-Output $Output | Select-Object -ExpandProperty FullName
                }
                Else {
                    Write-Output $Output
                }
            }
            Else {
                Write-Warning "[$Path] No match found for filter and/or type at current level"
            }
        } 
        Catch {
            Write-Warning "[$Path] $($_.Exception.Message)"
        }
    }
}

End {
    Write-Verbose "[END: $ScriptName]"
}
}