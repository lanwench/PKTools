#Requires -version 3
Function Show-PKJsonContent {
<# 
.SYNOPSIS
    Does something cool, interactively or as a PSJob

.DESCRIPTION
    Does something cool, interactively or as a PSJob
    Accepts pipeline input
    Returns a PSobject or PSJob

.NOTES        
    Name    : Function_Show-PKJsonContent.ps1
    Created : 2018-01-19
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2018-01-19 - Created script

.PARAMETER ComputerName
    Name of computer to do cool thing on; separate multiple names with commas

.PARAMETER Credential
    Valid credentials on target computer (default is current user credentials)



.PARAMETER SuppressConsoleOutput
    Hide all non-verbose/non-error console output

.EXAMPLE
    PS C:\> Do-SomethingCool -ComputerName foo

        
#> 

[CmdletBinding(
    DefaultParameterSetName = "__DefaultParameterSet",
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
[OutputType([Array])]
Param (
    [Parameter(
        Position = 0,
        Mandatory=$True,
        ValueFromPipeline=$True,
        HelpMessage="JSON object"
    )]
    [ValidateNotNullOrEmpty()]
    [object] $InputObject,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Depth for property expansion"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateRange(1,9)]
    [int]$Depth = 4,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Hide all non-verbose/non-error console output"
    )]
    [Switch] $SuppressConsoleOutput

)


Begin { 
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptPath",$PSCommandPath)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"


    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference    = "Continue"
   
    
    $JsonViewFile = "c:\users\pkingsley\git\personal\pktools\Files\jsonview\jsonview.exe"
    If (-not ($JSONView = (Get-Item $JsonViewFile -ErrorAction SilentlyContinue).FullName)) {
        $Msg = "JSONView.exe not found`nhttps://jsonviewer.codeplex.com/"
        $Host.UI.WriteErrorLine($Msg)
        Break
    }
    Else {
        $Msg = "Found $JSONView"
        Write-Verbose $Msg
    }
    


    <#

    #region Splats

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    #
    
    # Splat for Write-Progress
    $Activity = "Do a cool thing"
    If ($AsJob.IsPresent) {
        $Activity = "$Activity as remote PSJob"
        If ($WaitForJob.IsPresent) {$Activity = "$Activity (wait $JobWaitTimeout second(s) for job output)"}
    }
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = "Working"
        PercentComplete  = $Null
    }

    # Parameters for Test-WSMan
    $Param_WSMAN = @{}
    $Param_WSMAN = @{
        ComputerName   = ""
        Credential     = $Credential
        Authentication = "Kerberos"
        ErrorAction    = "Silentlycontinue"
        Verbose        = $False
    }
   

    # Parameters for Invoke-Command
    $ConfirmMsg = $Activity
    $Param_IC = @{}
    $Param_IC = @{
        ComputerName   = ""
        Authentication = "Kerberos"
        ScriptBlock    = $ScriptBlock
        Credential     = $Credential
        ErrorAction    = "Stop"
        Verbose        = $False
    }
    If ($AsJob.IsPresent) {
        $Jobs = @()
        $Param_IC.AsJob = $True
        $Param_IC.JobName = $Null
        $JobPrefix = "Thing"
    }
    
    #endregion Splats

    #>

 
    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "Action: $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg}


} #end begin

Process {

# Need to use reserved PowerShell variable $input to bypass the "one object at a time" default processing of pipelined input
    $InputObject = $input
    $Date = Get-Date
 
    # Set the file path to the temporary file we'll save the exported JSON to so it can be loaded into jsonview.exe
    # We will constantly overwrite/reuse this file since we have no way of safely deleting it without knowing for sure it's not being used by the jsonview.exe
    $TempFilePath = $env:TEMP+"\PSjsonviewtemp.json"
 
    # Create a status bar so if the conversion takes a long time to run, the user has some kind of visual feedback it's still running  
    $Params = @{
            "Activity" = "[$Date] Converting object to JSON.  This process may take several minutes to complete." 
            "Status" = "(Note: Increasing the -depth paramater beyond 4 may result in substantially longer processing times)"
            "Id" = 1
        }
    
    Write-Progress @Params
 
    # Convert the object to JSON.  Need to use the Compress option to get around a bug when processing some Powershell Objects
    try { 
        $JSON = $InputObject | ConvertTo-Json -Compress -Depth $Depth 
    }
    catch { 
        Write-Warning "This object cannot be converted to JSON. Try selecting a sub-property and trying again.`n$_"; break 
    }
    Write-Progress "Completed converting the object to JSON..." -id 1 -Completed 
    
    # Write the JSON to the temporary file and then open it with jsonview.exe
    $JSON | Out-File $TempFilePath -Force
 
    # Call the external JSON view application and pass it the file to display
    Start-Process -FilePath $JSONView -ArgumentList $TempFilePath


}
End {
}  
  
} # end Do-SomethingCool



<#


# Displays all properties and sub properties of an object in a GUI
# JSON; Home; Properties; Object; GUI
# Requires -version 3
# JSONViewer Download Available At: https://jsonviewer.codeplex.com/
 
Function Show-AllProperties
{
    param
    (
        [Parameter(ValueFromPipeline=$true)] $InputObject,
        [Parameter(ValueFromPipeline=$false)]
        [ValidateRange(1,9)]
        [int]$Depth=4
    )
 
    # This specifies how many layers deep the JSON will go.  It seems that anything more than 5 can result in a signifigant performance penalty
    $JSONViewPath = "c:\bin\jsonview\jsonview.exe"
 
    if(!(Test-Path $JSONViewPath)) { Write-Error "$JSONViewPath is not found.  This is required."; break}
 
    # Need to use reserved PowerShell variable $input to bypass the "one object at a time" default processing of pipelined input
    $InputObject = $input
    $Date = Get-Date
 
    # Set the file path to the temporary file we'll save the exported JSON to so it can be loaded into jsonview.exe
    # We will constantly overwrite/reuse this file since we have no way of safely deleting it without knowing for sure it's not being used by the jsonview.exe
    $TempFilePath = $env:TEMP+"\PSjsonviewtemp.json"
 
    # Create a status bar so if the conversion takes a long time to run, the user has some kind of visual feedback it's still running  
    $Params = @{
            "Activity" = "[$Date] Converting object to JSON.  This process may take several minutes to complete." 
            "Status" = "(Note: Increasing the -depth paramater beyond 4 may result in substantially longer processing times)"
            "Id" = 1
        }
    
    Write-Progress @Params
 
    # Convert the object to JSON.  Need to use the Compress option to get around a bug when processing some Powershell Objects
    try { $JSON = $InputObject | ConvertTo-Json -Compress -Depth $Depth }
    catch { Write-Warning "This object cannot be converted to JSON. Try selecting a sub property and trying again.`n$_"; break }
    Write-Progress "Completed converting the object to JSON..." -id 1 -Completed 
    
    # Write the JSON to the temporary file and then open it with jsonview.exe
    $JSON | Out-File $TempFilePath -Force
 
    # Call the external JSON view application and pass it the file to display
    Start-Process -FilePath $JSONViewPath -ArgumentList $TempFilePath
}
 
# Set a three letter alias to make it faster to run
Set-Alias sap Show-AllProperties

#>