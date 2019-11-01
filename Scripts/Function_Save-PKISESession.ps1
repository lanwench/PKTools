#requires -version 3
Function Save-PKISESession{
<#
.SYNOPSIS 
    Saves open tabs in current ISE session to a file

.DESCRIPTION
    Saves open tabs in current ISE session to a file
    
.NOTES
    Name    : Function_Save-PKISESession.ps1
    Created : 2016-05-29
    Author  : Paula Kingsley
    Version : 02.01.000
    History :
    
        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v1.0.0      - 2016-05-29 - Created script based on links
        v02.00.0000 - 2018-02-14 - Updated with default path, standardization, changed force to noclobber
        v02.01.0000 - 2019-10-08 - Minor cosmetic updates
        
.LINK
    https://itfordummies.net/2014/10/27/save-restore-powershell-ise-opened-scripts/

.LINK
    https://stackoverflow.com/questions/3710374/get-encoding-of-a-file-in-windows   

.EXAMPLE
    PS C:\> Save-PKISESession -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value                              
        ---           -----                              
        Verbose       True                               
        SessionFile   C:\Users\jbloggs\PSISESession.txt
        NoClobber     False                              
        ScriptName    Save-PKISESession                  
        ScriptVersion 2.0.0                              

        WARNING: File 'C:\Users\jbloggs\PSISESession.txt' already exists and will be overwritten
        VERBOSE: Save 13 current tab(s) to 'C:\Users\jbloggs\PSISESession.txt'
        VERBOSE: Saved 13 tab(s) to file 'C:\Users\jbloggs\PSISESession.txt'

.EXAMPLE
    PS :\> Save-PKISESession -SessionFile c:\temp\newsessionfile.txt -NoClobber

        ERROR: File 'c:\temp\newsessionfile.txt' already exists and will not be overwritten (-NoClobber detected)

.EXAMPLE
    PS C:\> Save-PKISESession -SavePath c:\temp\newsessionfile.txt

        WARNING: File 'c:\temp\newsessionfile.txt' already exists and will be overwritten
        Operation cancelled by user


#>

[CmdletBinding(
    SupportsShouldProcess=$True,
    ConfirmImpact = "High"
)]
Param(
    [Parameter(
        Position=0,
        HelpMessage = "Output file for session storage (will be created if nonexistent)"
    )]
    [ValidateNotNullOrEmpty()]
    [String]$SessionFile = "$Home\PSISESession.txt",

    [Parameter(
        HelpMessage = "Don't overwrite existing entries in file"
    )]
    [ValidateNotNullOrEmpty()]
    [switch]$NoClobber

)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "02.01.0000"

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences 
    $ErrorActionPreference = "Stop"
    
    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose = $False
    }

    If ($Host.Name -ne "Windows PowerShell ISE Host") {
        $Msg = "This script requires the PowerShell ISE; current  host is '$($Host.Name)'"
        $Host.UI.WriteErrorLine($Msg)
        Break
    }
}
Process{

    If ($Null = Test-Path $SessionFile -ErrorAction SilentlyContinue) {
        $Msg = "File '$SessionFile' already exists"
        If ($NoClobber.IsPresent) {
            $Msg = "$Msg and will not be overwritten (-NoClobber detected)"
            $Host.UI.WriteErrorLine("ERROR: $Msg")
            Break
        }
        Else {
            $Msg = "$Msg and will be overwritten"
            Write-Warning $Msg    
        }        
    }

    $Msg = "Save $($psISE.PowerShellTabs.Files.Count) current tab(s) to '$SessionFile'"
    Write-Verbose $Msg
    $ConfirmMsg = "`n`n`t$Msg`n`n"
    If ($PSCmdlet.ShouldProcess($Env:ComputerName,$ConfirmMsg)) {
        Try {
            $psISE.CurrentPowerShellTab.Files | Foreach-Object {$_.SaveAs($_.FullPath)}
            
            # Using forward spash for delimiter as we may have commas or semicolons in the filename
            $($psISE.PowerShellTabs.Files.FullPath -join'/') | Out-File -Encoding UTF8 -FilePath $SessionFile -Confirm:$False -EA Stop -Verbose:$False

            $Msg = "Saved $($psISE.CurrentPowerShellTab.Files.count) tab(s) to file '$SessionFile'"
            Write-Verbose $Msg
        }
        Catch {
            $Msg = "Operation failed"
            If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
            $Host.UI.WriteErrorLine($Msg)
        }
    }
    Else {
        $Msg = "Operation cancelled by user"
        $Host.UI.WriteErrorLine($Msg)
    }
}
} #end Save-PKISESession






    
