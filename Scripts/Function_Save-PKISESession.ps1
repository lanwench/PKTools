#requires -version 3
Function Save-PKISESession{
<#
.SYNOPSIS 
    Saves open tabs in current ISE session to a file

.DESCRIPTION
    Saves open tabs in current ISE session to a file
    
.NOTES
    Name    : Function_Save-PKISESession.ps1
    Author  : Paula Kingsley
    Version : 1.0.0
    History :
    
        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v1.0.0 - 2016-05-29 - Created script based on links
        
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
        Force         False                                   
        SavePath      C:\Users\PKINGS~1\AppData\Local\Temp\ISESession.txt
        ScriptName    Save-PKISESession                                  
        ScriptVersion 1.0.0           
                                           
        VERBOSE: Save 10 current tab(s) to C:\Users\PKINGS~1\AppData\Local\Temp\ISESession.txt
        VERBOSE: Saved tabs to file

.EXAMPLE
    PS C:\> Save-PKISESession -SavePath c:\temp\foo.log -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value            
        ---           -----            
        SavePath      c:\temp\foo.log  
        Verbose       True             
        Force         False            
        ScriptName    Save-PKISESession
        ScriptVersion 1.0.0            

        VERBOSE: Save 10 current tab(s) to c:\temp\foo.log
        Operation cancelled by user

.EXAMPLE
    PS C:\> Save-PKISESession -SavePath c:\temp\foo.log -Verbose
 
        VERBOSE: PSBoundParameters: 
	
        Key           Value            
        ---           -----            
        SavePath      c:\temp\foo.log  
        Verbose       True             
        Force         False            
        ScriptName    Save-PKISESession
        ScriptVersion 1.0.0            

        VERBOSE: Save 10 current tab(s) to c:\temp\foo.log
        VERBOSE: Saved tabs to file

.EXAMPLE
    PS C:\> Save-PKISESession -SavePath c:\temp\foo.log -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value            
        ---           -----            
        SavePath      c:\temp\foo.log  
        Verbose       True             
        Force         False            
        ScriptName    Save-PKISESession
        ScriptVersion 1.0.0            

        File 'c:\temp\foo.log' already exists; please use -Force to overwrite

.EXAMPLE
    PS C:\> Save-PKISESession -SavePath c:\temp\foo.log -Force -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value            
        ---           -----            
        SavePath      c:\temp\foo.log  
        Force         True             
        Verbose       True             
        ScriptName    Save-PKISESession
        ScriptVersion 1.0.0            

        WARNING: File 'c:\temp\foo.log' already exists and will be overwritten
        VERBOSE: Save 10 current tab(s) to c:\temp\foo.log
        VERBOSE: Saved tabs to file

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
    [String]$SavePath = "$env:temp\ISESession.txt",

    [Parameter(
        HelpMessage = "Overwrite existing entries in file"
    )]
    [ValidateNotNullOrEmpty()]
    [switch]$Force

)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "1.0.0"

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

    If ($Null = Test-Path $SavePath -ErrorAction SilentlyContinue) {
        $Msg = "File '$SavePath' already exists"
        If ($Force.IsPresent) {
            $Msg = "$Msg and will be overwritten"
            Write-Warning $Msg    
        }
        Else {
            $Msg = "$Msg; please use -Force to overwrite"
            $Host.UI.WriteErrorLine($Msg)
            Break
        }
    }

    $Msg = "Save $($psISE.PowerShellTabs.Files.Count) current tab(s) to '$SavePath'"
    Write-Verbose $Msg
    If ($PSCmdlet.ShouldProcess($Env:ComputerName,$Msg)) {
        Try {
            $psISE.CurrentPowerShellTab.Files | Foreach-Object {$_.SaveAs($_.FullPath)}
            #"ise ""$($psISE.PowerShellTabs.Files.FullPath -join'/')""" | Out-File -Encoding UTF8 -FilePath $SavePath -Confirm:$False -EA Stop -Verbose:$False
            
            # Using forward spash for delimiter as we may have commas or semicolons in the filename

            #"""$($psISE.PowerShellTabs.Files.FullPath -join'/')""" | Out-File -Encoding UTF8 -FilePath $SavePath -Confirm:$False -EA Stop -Verbose:$False
            $($psISE.PowerShellTabs.Files.FullPath -join'/') | Out-File -Encoding UTF8 -FilePath $SavePath -Confirm:$False -EA Stop -Verbose:$False

            #"ise ""$($psISE.PowerShellTabs.Files.FullPath -join',')""" | Out-File -Encoding UTF8 -FilePath $SavePath -Confirm:$False -EA Stop -Verbose:$False
            #"ise ""$($psISE.PowerShellTabs.Files.FullPath -join'";"')""" | Out-File -Encoding UTF8 -FilePath $SavePath -Confirm:$False -EA Stop -Verbose:$False
            
            $Msg = "Saved $($psISE.CurrentPowerShellTab.Files.count) tab(s) to file '$SavePath'"
            Write-Verbose $Msg
        }
        Catch {
            $Msg = $_.Exception.Message
            $Host.UI.WriteErrorLine($Msg)
        }
    }
    Else {
        $Msg = "Operation cancelled by user"
        $Host.UI.WriteErrorLine($Msg)
    }
}
} #end Save-PKISESession






    
