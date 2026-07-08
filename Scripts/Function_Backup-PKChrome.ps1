#requires -Version 4
Function Backup-PKChromeProfile {
<#
.SYNOPSIS
    Backs up Google Chrome user profile data to a date-stamped folder

.DESCRIPTION
    Backs up Google Chrome user profile data to a date-stamped folder in the target path
    Backs up all user data or profile folders only (-SourceType ProfilesOnly)
    Optionally includes cache files (-IncludeCache); excluded by default
    Stops Chrome if running before copying
    Supports -WhatIf and -Confirm (ShouldProcess)
    Windows only
    Requires Chrome installed at the default LocalAppData path
    Returns a PSObject

.NOTES
    Name    : Function_Backup-PKChromeProfile
    Created : 2018-04-16
    Author  : Paula Kingsley
    Version : 01.02
    History :

        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v01.00 - 2018-04-16 - Created script
        v01.01 - 2018-08-29 - Updated, added comment block
        v01.02 - 2026-06-16 - Cosmetic updates; converted version format to major.minor

.PARAMETER ProfilePath
    Absolute path to source profile files; default is '$env:LOCALAPPDATA\Google\Chrome\User Data'

.PARAMETER TargetPath
    Absolute path to target directory; default is '$env:USERPROFILE'

.PARAMETER SourceType
    Copy all folder contents (All), or copy only folders matching 'Default' and 'Profile' (ProfilesOnly)

.PARAMETER IncludeCache
    Include cache files (this will slow down processing)

.PARAMETER SuppressConsoleOutput
    Suppress all non-verbose/non-error console output

.EXAMPLE
    PS C:\> Backup-PKChromeProfile -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                                                   
        ---                   -----                                                   
        Verbose               True                                                    
        ProfilePath           C:\Users\jbloggs\AppData\Local\Google\Chrome\User Data
        TargetPath            C:\Users\jbloggs                                      
        SourceType            All                                                     
        IncludeCache          False                                                   
        SuppressConsoleOutput False                                                   
        ScriptName            Backup-PKChromeProfile                                  
        ScriptVersion         01.02

        Action: Copy all Chrome user data to C:\Users\jbloggs\2018-08-29_04-30-54_Chrome_jbloggs-05122

        VERBOSE: Look for Google profile files
        VERBOSE: Found Google profile(s) in 'c:\users\jbloggs\appdata\local\google\chrome\user data'
        VERBOSE: 19,625 file(s) copied to 'C:\Users\jbloggs\2018-08-29_04-30-54_Chrome_WORKSTATION14'

.EXAMPLE
    PS C:\> Backup-PKChromeProfile -TargetPath c:\temp -IncludeCache -SourceType ProfilesOnly -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                                                   
        ---                   -----                                                   
        TargetPath            c:\temp                                                 
        IncludeCache          True
        SourceType            ProfilesOnly                                            
        Verbose               True                                                    
        ProfilePath           C:\Users\jbloggs\AppData\Local\Google\Chrome\User Data
        SuppressConsoleOutput False                                                   
        ScriptName            Backup-PKChromeProfile                                  
        ScriptVersion         01.02

        Action: Copy all Chrome user profile data folders to c:\temp\2018-08-29_04-11-49_Chrome_WORKSTATION22 (excluding cache files)

        VERBOSE: Look for Google profile files
        VERBOSE: Found Google profile(s) in 'c:\users\jbloggs\appdata\local\google\chrome\user data'
        VERBOSE: Stop running Chrome process(es)
        VERBOSE: Operation cancelled by user
    
#>
[CmdletBinding(SupportsShouldProcess,ConfirmImpact="High")]
Param(
    [Parameter(
        HelpMessage = "Absolute path to source file(s)"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({Test-Path $_})]
    [string]$ProfilePath = "$env:LOCALAPPDATA\Google\Chrome\User Data",

    [Parameter(
        HelpMessage = "Absolute path for backup files"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({Test-Path $_})]
    [string]$TargetPath = $env:USERPROFILE,

    [Parameter(
        HelpMessage = "Copy all folder contents, or copy only folders matching 'Default' and 'Profile'"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("All","ProfilesOnly")]
    [string]$SourceType = "All",

    [Parameter(
        HelpMessage = "Include cache files (this will slow down up processing)"
    )]
    [switch]$IncludeCache
    
)
Begin{

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.02"

    # How did we get here?
    $ScriptName = $MyInvocation.MyCommand.Name
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object {$CurrentParams.keys -notContains $_} |
        Where-Object {Test-Path variable:$_} | ForEach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    If ($IsWindows -eq $False) {
        $Msg = "This function requires Windows"
        Write-Error $Msg
        Return
    }

    # General-purpose splat
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose = $False
    }

    # Dates
    $FileDate = Get-Date -f yyyy-MM-dd_hh-mm-ss @StdParams
    $Target = "$TargetPath\$FileDate`_Chrome_$Env:ComputerName"

    # Message for verbose / progress output
    Switch ($SourceType) {
        All {$Msg = "Copy all Chrome user data to $Target"}
        ProfilesOnly {$Msg = "Copy all Chrome user profile data folders to $Target"}
    }
    If ($ExcludeCache.IsPresent) {$Msg += " (excluding cache files)"}
    $Activity = "Copy Chrome profile"

    # Make it look pretty
    $SourcePath = $ProfilePath.ToLower()

    Write-Verbose "[BEGIN: $ScriptName] $Msg"
    

}
Process {

    # Set flag
    [switch]$Continue = $False
        
    Try {
        # Test path
        $Msg = "Look for Google profile files"
        Write-Verbose $Msg
        Write-Progress -Activity $Activity -CurrentOperation $Msg -Status Working
        
        If ($Null = Test-Path $SourcePath -ErrorAction SilentlyContinue -Verbose:$False) {
            $FileList = @()
            $Msg = "Found Google profile(s) in '$SourcePath'"
            Write-Verbose $Msg

            # Get the source files using Get-ChildItem
            Switch ($SourceType) {
                All {
                    $Filelist = Get-ChildItem $SourcePath -Recurse @StdParams
                }
                ProfilesOnly {
                    $FileList += (Get-ChildItem -Path $SourcePath -Filter "default" -Recurse -Force @StdParams)
                    $FileList += (Get-ChildItem -Path $SourcePath -Filter "profile*" -Recurse -Force @StdParams)
                }
            }

            # Filter
            If (-not $IncludeCache.IsPresent) {
                $FileList = $FileList | Where-Object {$_.fullname -notmatch "cache"}
            }
            $FileList = $FileList | Where-Object {$_ -notmatch "lockfile"}

            # Reset flag
            $Continue = $True
        }
        Else {
            $Msg = "Invalid path '$SourcePath'"
            Write-Error $Msg
        }
        
        # Continue if files found
        If ($Continue.IsPresent) {
            
            # Reset flag
            $Continue = $False
            
            # Stop Chrome if it's running
            If ($ChromeProc = Get-Process chrome -ErrorAction SilentlyContinue) {
                $Msg = "Stop running Chrome process(es)"
                Write-Verbose $Msg
                Write-Progress -Activity $Activity -CurrentOperation $Msg -Status Working

                $ConfirmMsg = "`n$Msg`n`n"
                If ($PSCmdlet.ShouldProcess($Env:ComputerName,$Msg)) {
                    $Null = $ChromeProc | Stop-Process -Force -PassThru -Confirm:$False @StdParams
                    $Msg = "Stopped Chrome"
                    Write-Verbose $Msg
                    $Continue = $True
                }
                Else {
                    $Msg = "Operation cancelled by user"
                    Write-Verbose $Msg
                }
            }
            Else {
                $Continue = $True
            }
        }

        # If Chrome isn't running, prompt to copy files
        If ($Continue.IsPresent) {
            Try {
                $ConfirmMsg = "`n`n`tCopy $SourcePath files to`n`t$Target`n`n"
                If ($PSCmdlet.ShouldProcess($Env:ComputerName,$ConfirmMsg)) {    
                    # Set counters
                    $Current = 0
                    $Total = $FileList.Count
                    $Msg = "Copy profile files to $Target"
                    Write-Verbose

                    Foreach ($File in $Filelist) {                        
                        $Current ++
                        $TargetFile = $File.Fullname.tolower().replace($SourcePath,'')
                        $DestinationFile = "$Target$TargetFile"
                        Write-Progress -Activity $Activity -CurrentOperation $File.FullName -Status "Copying file" -PercentComplete ($Current/$Total*100)
                        $Null = Copy-Item $File.FullName -Destination $DestinationFile -Force -PassThru -Confirm:$False @StdParams
                    }

                    $GetFiles = Get-ChildItem $Target -Recurse -File @StdParams
                    $Msg = "$('{0:N0}' -f $($GetFiles.Count)) file(s) copied to '$Target'"
                    Write-Verbose $Msg
                }
                Else {
                    $Msg = "Operation cancelled by user"
                    Write-Verbose $Msg
                }
            }
            Catch {
                $Msg = "File copy failed"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
                Write-Error $Msg
            }
        }
    }
    Catch {
        $Msg = "Operation failed"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
        Write-Error $Msg
    }
}
End {
    Write-Verbose "[END: $ScriptName] Script ran successfully"
}
} #end Backup-PKChromeProfile
