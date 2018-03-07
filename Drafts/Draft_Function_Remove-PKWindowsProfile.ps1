#requires -Version 3
Function Get-PKWindowsProfile {
[Cmdletbinding(
    DefaultParameterSetName = "All",
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
Param(
    [parameter(
        Mandatory = $False,
        Position = 0,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "Target hostname or FQDN"
    )]
    [Alias("Hostname","Name","FQDN","DNSHostName")]
    [ValidateNotNullOrEmpty()]
    [String[]]$ComputerName = $env:COMPUTERNAME,

    [Parameter(
        ParameterSetName = "Inactive",
        Mandatory = $False,
        Position = 1,
	    HelpMessage = "Display only user profiles inactive for X days"
    )]
    [switch]$InactiveOnly,

    [Parameter(
        ParameterSetName = "Inactive",
        Mandatory = $False,
        Position = 2,
	    HelpMessage = "Age in days of last profile use (default is 90; ignored if displaying all profiles)"
    )]
    [ValidateNotNullOrEmpty()]
    [Int32]$DaysInactive = 90,
	
    [Parameter(
        Mandatory = $False,
        Position = 3,
	    HelpMessage = "Include system/built-in users (not included by default)"
    )]
    [switch]$IncludeSystemUsers,

	[Parameter(
        Mandatory = $False,
        Position = 4,
	    HelpMessage = "Names of user accounts to exclude (ignored if displaying all profiles)"
    )]
   
    [ValidateNotNullOrEmpty()]
    [String[]]$ExcludedUsers,

    [Parameter(
        Mandatory = $False,
        Position = 5,
	    HelpMessage = "Get total folder size for profile (may delay script processing)"
    )]
    [switch]$CalculateFolderSize,

    [Parameter(
        Mandatory = $False,
        Position = 6,
	    HelpMessage = "Run as PSJob"
    )]
    [switch]$AsJob,

    [Parameter(
        Mandatory = $False,
        Position = 7,
        HelpMessage = "Valid credentials on target computer(s)"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential] $Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(
        Mandatory = $False,
        Position = 8,
        HelpMessage = "Don't test WinRM connection to computer"
    )]
    [ValidateNotNullOrEmpty()]
    [switch]$SkipConnectionTest

)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Detect pipeline input
    [switch]$PipelineInput = (-not $PSBoundParameters.ContainsKey("ComputerName")) -and (-not $ComputerName)

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference = "Continue"

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }
    
    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # For write-progress
    Switch ($Source) {
        All {$Activity = "List all user profiles"}
        Inactive {$Activity = "List user profiles inactive for $DaysInactive day(s) or greater"}
    }
    If ($IncludeSystemUsers.IsPresent) {
        $Activity = "$Activity, including built-in/system users"
    }

    # Splat for test-wsman
    $Param_WSMan = @{}
    $Param_WSMan = @{
        ComputerName   = ""
        Authentication = "Negotiate"
        Credential     = $Credential
        ErrorAction    = "SilentlyContinue"
        Verbose        = $False
    }

    # Splat for write-progress
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        PercentComplete  = ""
        CurrentOperation = $Null
        Status           = "Working"
    }

    #Splat for invoke-command
    $Param_IC = @{}
    $Param_IC = @{
        ComputerName = ""
        Credential   = $Credential
        Scriptblock  = $ScriptBlock
        ArgumentList = $ServiceFilter,$IncludeNoMatch
        AsJob        = $True
        JobName      = ""
        ErrorAction  = "Stop"
        Verbose      = $False
    }



    # Splat for get-wmiobject
    If ($IncludeSystemUsers.IsPresent) {
        $Query = "Select * from win32_userprofile"
    }
    Else {
        $Query = "Select * from win32_userprofile where special=False AND loaded=False AND NOT localpath like '%c:\\windows%'"
    }
    $Param_GWMI = @{
        ComputerName = $Null
        Query        = $Query
        ErrorAction  = "SilentlyContinue"
        Verbose      = $False
    }

    # Create regex for partial matches in array of excluded usernames
    If ($ExcludedUsers) {
        [regex]$NotMatch = '(' + (($ExcludedUsers | Foreach {[regex]::escape($_)}) –join "|") + ')'
        If ($ExcludedUsers.Count -gt 1) {
            $ExcludedStr = "'$ExcludedUsers'"
        }
        Else {
            $ExcludedStr = "'$($ExcludedUsers -join(''', '''))'"
        }
    }

   
    Write-Verbose $Activity

    #Splat for invoke-command
    $Param_IC = @{}
    $Param_IC = @{
        ComputerName = ""
        Credential   = $Credential
        Scriptblock  = $ScriptBlock
        ArgumentList = $Action,$ExcludedUsers,$IncludeSystemUsers,$Age
        AsJob        = $True
        JobName      = ""
        ErrorAction  = "Stop"
        Verbose      = $False
    }

    # Arguments for Invoke-Command
    $ArgList = $Action,$ExcludedUsers,$IncludeSystemUsers,$Age

    # Create regex for partial matches in array of excluded usernames
    If ($ExcludedUsers) {
        [regex]$NotMatch = '(' + (($ExcludedUsers | Foreach {[regex]::escape($_)}) –join "|") + ')'
        If ($ExcludedUsers.Count -gt 1) {
            $ExcludedStr = "'$ExcludedUsers'"
        }
        Else {
            $ExcludedStr = "'$($ExcludedUsers -join(''', '''))'"
        }
        $ArgList += $NotMatch,$ExcludedStr
    }


    $ScriptBlock = {
        Param($ArgList)

        [string]$Action = $Using:Action
        [switch]$IncludeSystemUsers = $Using:IncludeSystemUsers
        
        If ($Using:Age) {
            [int32]$Age = $Using:Age
        }

        If ($Using:NoMatch) {
            [regex]$NoMatch = $Using:NoMatch
        }

        # Splat for get-wmiobject
        If ($IncludeSystemUsers.IsPresent) {
            $Query = "Select * from win32_userprofile"
        }
        Else {
            $Query = "Select * from win32_userprofile where special=False AND loaded=False AND NOT localpath like '%c:\\windows%'"
        }
        $Param_GWMI = @{
            Query        = $Query
            ErrorAction  = "SilentlyContinue"
            Verbose      = $False
        }

        Try {
            $Profiles = @()
            $Profiles = Get-WmiObject @Param_GWMI
        
            If ($Profiles.Count -gt 0) {
                    
                # We're just showing everything we found
                If ($Action -eq "DisplayAll") {
                        
                    If ($ExcludedUsers) {
                        $Output = $Profiles  | Where-Object {$_.LocalPath -notmatch $NotMatch.ToString()} 
                        $Msg = "Found $($Output.Count) user profile(s) excluding $ExcludedStr"
                    }
                    Else {
                        $Output = $Profiles
                        $Msg = "Found $($Output.Count) user profile(s)"
                    }
                    Write-Verbose $Msg
                        
                    $Output = $Output |  Select-Object `
                        @{Label = "ComputerName";Expression={$_.__SERVER}},
	                    @{Label = "UserName";Expression = {$_.LocalPath | Split-Path -Leaf }},
                        @{Label = "LastUseTime";Expression={$_.ConvertToDateTime($_.LastUseTime)}},
                        @{Label = "SizeGB";Expression={
                            
                            $All = Get-ChildItem $Output.LocalPath -Recurse -Force -File


                            }
                        },
                        LocalPath,
                        SID 
                    
                    If ($CalculateFolderSize.IsPresent) {
                        $Param_Reg = @{
                            
                        }
                        
                        $Output | Foreach-Object {
                            

                        }
                        
                        
                    }
                    
                    
                    
Get-WmiObject -Query "ASSOCIATORS OF {Win32_Directory.Name='C:\'} Where ResultClass = CIM_DataFile"
                    
                    
                    
                    
                    Write-Output $Output
                }
                # Or we are filtering on date and possibly name, to either display or delete
                    Else {
                        
                        $Output = ($Profiles | Where-Object { (-not ($_.LastUseTime) -or (($_.ConvertToDateTime($_.LastUseTime)) -lt ((Get-Date).AddDays(-$DaysInactive)))) })

                        If ($ExcludedUsers) {
                            $Output = $Output  | Where-Object {$_.LocalPath -notmatch $NotMatch.ToString()} 
                            $Msg = "Found $($Output.Count) user profile(s) last used more than $DaysInactive day(s) ago, excluding $ExcludedStr"
                        }
                        Else {$Msg = "Found $($Output.Count) user profile(s) last used more than $DaysInactive day(s) ago"}
                        
                        If ($Output.Count -gt 0) {

                            $Display = $Output |  Select-Object `
                            @{Label="ComputerName";Expression={$_.__SERVER}},
	                        @{Label = "UserName";Expression = {$_.LocalPath | Split-Path -Leaf }},
                            @{Label="LastUseTime";Expression={$_.ConvertToDateTime($_.LastUseTime)}},
                            LocalPath,
                            SID
                            
                            Write-Verbose $Msg
                            Write-Output $Display | Sort LastUseTime -Descending
                            
                        } #end if any old profiles found

                        Else {
                            $Msg = "No matching profile(s) found on $Computer"
                            $Host.UI.WriteErrorLine($Msg)
                        }
                    } # end if age
                } #end if any profiles found

                Else {
                    $Msg = "No user profiles found on $Env:ComputerName"
                    $Host.UI.WriteErrorLine("ERROR: $Msg")
                }
            
            }
            Catch {
                $Msg = "WMI query failed on $Env:ComputerName"
                $ErrorDetails = $_.Exception.Message
                $Host.UI.WriteErrorLine("ERROR: $Msg`n$ErrorDetails")
            }
        
        } #end scriptblock
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    }

}
Process {

    $Total = $ComputerName.Count
    $Current = 0

    Foreach ($Computer in $ComputerName) {
        
        $Current ++
        $Param_WP.PercentComplete = ($Current / $Total * 100)
        $Param_WP.CurrentOperation = $Param_WSMAN.ComputerName = $Param_GWMI.ComputerName = $Computer
        $Msg = $computer
        Write-Verbose $Msg
        Write-Progress @Param_WP

        [switch]$Continue = $False
        
        #If ($ComputerName) {
        If (-not $SkipConnectionTest.IsPresent) {
            $Msg = "Test WinRM connection"
            If ($PSCmdlet.ShouldProcess($Computer,$Msg)) {
                Try {
                    $Null = Test-WSMan @Param_WSMAN
                    $Continue = $True
                }
                Catch {
                    $Msg = "WinRM connection failed on $Computer"
                    $ErrorDetails = $_.Exception.Message
                    $Host.UI.WriteErrorLine("ERROR: $Msg`n$ErrorDetails")
                }
            } #end if testing WinRM
            Else {
                $Msg = "WinRM connection test skipped"
                Write-Warning $Msg
                $continue = $True
            }
        }
        Else {
            $Continue = $True
        }

        # Either we could connect or we're going to try without testing connectivity
        If ($Continue -eq $True) {
            Try {
                $Profiles = @()
                $Profiles = Get-WmiObject @Param_GWMI
                
	            If ($Profiles.Count -gt 0) {
                    
                    # We're just showing everything we found
                    If ($Action -eq "DisplayAll") {
                        
                        If ($ExcludedUsers) {
                            $Output = $Profiles  | Where-Object {$_.LocalPath -notmatch $NotMatch.ToString()} 
                            $Msg = "Found $($Output.Count) user profile(s) excluding $ExcludedStr"
                        }
                        Else {
                            $Output = $Profiles
                            $Msg = "Found $($Output.Count) user profile(s)"
                        }
                        Write-Verbose $Msg
                        
                        $Output = $Output |  Select-Object `
                            @{Label = "ComputerName";Expression={$_.__SERVER}},
	                        @{Label = "UserName";Expression = {$_.LocalPath | Split-Path -Leaf }},
                            @{Label = "LastUseTime";Expression={$_.ConvertToDateTime($_.LastUseTime)}},
                            LocalPath,
                            SID 
                        Write-Output $Output
                    }
                    # Or we are filtering on date and possibly name, to either display or delete
                    Else {
                        
                        $Output = ($Profiles | Where-Object { (-not ($_.LastUseTime) -or (($_.ConvertToDateTime($_.LastUseTime)) -lt ((Get-Date).AddDays(-$DaysInactive)))) })

                        If ($ExcludedUsers) {
                            $Output = $Output  | Where-Object {$_.LocalPath -notmatch $NotMatch.ToString()} 
                            $Msg = "Found $($Output.Count) user profile(s) last used more than $DaysInactive day(s) ago, excluding $ExcludedStr"
                        }
                        Else {$Msg = "Found $($Output.Count) user profile(s) last used more than $DaysInactive day(s) ago"}
                        
                        If ($Output.Count -gt 0) {

                            $Display = $Output |  Select-Object `
                            @{Label="ComputerName";Expression={$_.__SERVER}},
	                        @{Label = "UserName";Expression = {$_.LocalPath | Split-Path -Leaf }},
                            @{Label="LastUseTime";Expression={$_.ConvertToDateTime($_.LastUseTime)}},
                            LocalPath,
                            SID
                    
                            If ($Action -eq "DeleteInactive") {
                                Write-Verbose $Msg
                                $Msg = "Delete $($Output.Count) profile(s)"
                                Write-Verbose $Msg
                                Write-Verbose ($Display | Format-Table | Out-String)
                        
                                $Deleted = 0
                                Foreach ($Prof in $Output) {
                                    $Msg = "`nDelete $($Prof.LocalPath | Split-Path -Leaf), last used $($Prof.ConvertToDateTime($Prof.LastUseTime))`n"
                                    If ($PSCmdlet.ShouldProcess($Computer,$Msg)) {
                                        If ($Prof.Delete()) {
                                            $Deleted ++
                                        }
                                    }
                                    Else {
                                        $Msg = "Profile $($Prof.LocalPath | Split-Path -Leaf) deletion cancelled by user"
                                        Write-Verbose $Msg
                                    }
                                } #end for each
                                If ($Deleted.Count -gt 0) {
                                    $Msg = "$($Deleted.Count) profile(s) deleted"
                                    Write-Verbose $Msg
                                }
                                Else {
                                    $Msg = "No profile deleted"
                                    $Host.UI.WriteErrorLine($Msg)
                                }
                            } #end if delete
                            Else {
                                Write-Verbose $Msg
                                Write-Output $Display | Sort LastUseTime -Descending
                            }
                        } #end if any old profiles found

                        Else {
                            $Msg = "No matching profile(s) found on $Computer"
                            $Host.UI.WriteErrorLine($Msg)
                        }
                    } # end if age
                } #end if any profiles found

                Else {
                    $Msg = "No user profiles found on $Computer"
                    $Host.UI.WriteErrorLine("ERROR: $Msg")
                }
            
            }
            Catch {
                $Msg = "WMI query failed on $Computer"
                $ErrorDetails = $_.Exception.Message
                $Host.UI.WriteErrorLine("ERROR: $Msg`n$ErrorDetails")
                Continue
            }
        } 
        
    } #end for each computer

}
End {

    Write-Progress -Activity $Param_WP.Activity -Completed
}
}


