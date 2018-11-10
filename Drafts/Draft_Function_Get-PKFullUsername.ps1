#requires -version 3
Function Get-PKWindowsFullUsername {
<# 
.SYNOPSIS
    Searches the local registry and WMI for the current user's full name

.DESCRIPTION
    Searches the local registry and WMI for the current user's full name
    Returns a PSObject

.NOTES        
    Name    : Function_Get-PKWindowsFullUsername.ps1
    Created : 2018-11-06
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2018-11-06 - Created script

.PARAMETER AllUsers
    Return results for all users with locally cached profiles (default is current user only)


#>
[CmdletBinding(DefaultParameterSetName = "CurrentUser")]
Param(
    [Parameter(
        ParameterSetName = "CurrentUser"
    )]
    [switch]$CurrentUser,

    [Parameter(
        ParameterSetName = "MatchUser"
    )]
    [string]$UserName,

    [Parameter(
        ParameterSetName = "AllUsers"
    )]
    [switch]$AllUsers
)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"
    
    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    $CurrentParams.Add("ParameterSetName",$Source)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences 
    $ProgressPreference = "Continue"

}
Process {

    
    Try {
        
        [array]$Results = @()

        # Get all user profiles
        #$SIDLocalUsers = Get-WmiObject Win32_UserProfile -ErrorAction Stop | select-Object Localpath,SID
        
        $Status = "Working"
        Switch ($Source) {
            MatchUser {
                $Msg = "Get full name for specified username on '$Env:ComputerName'"
                $Activity = $Msg
                Write-Verbose $Msg
                $CurrentOp = "Get SID and path for cached user profiles"
                Write-Progress -Activity $Activity -CurrentOp $CurrentOp -Status $Status
                [array]$SIDLocalUsers = Get-WmiObject Win32_UserProfile -Filter "LocalPath LIKE '%$UserName'" -ErrorAction SilentlyContinue | select-Object Localpath,SID
            }
            AllUsers {
                $Msg = "Get full name for all usernames on '$Env:ComputerName'"
                Write-Verbose $Msg
                $Activity = $Msg
                $CurrentOp = "Get SID and path for cached user profiles"
                Write-Progress -Activity $Activity -CurrentOp $CurrentOp -Status $Status
                [array]$SIDLocalUsers = Get-WmiObject Win32_UserProfile -ErrorAction Stop | select-Object Localpath,SID
            }
            CurrentUser {
                $Msg = "Get full name for current username on '$Env:ComputerName'"
                Write-Verbose $Msg
                $Activity = $Msg
                $CurrentOp = "Get SID and path for cached user profiles"
                Write-Progress -Activity $Activity -CurrentOp $CurrentOp -Status $Status
                [array]$SIDLocalUsers = Get-WmiObject Win32_UserProfile -ErrorAction Stop | select-Object Localpath,SID
            }
        }

        If ($SIDLocalUsers) {
        
            Switch ($Source) {
                Named {
                    $Total = $SIDLocalUsers.Count
                    $Current = 0
                    Foreach ($User in $SIDLocalUsers) {
                        
                        $Msg = $User.LocalPath
                        Write-Verbose $Msg
                        
                        $Curent ++
                        $CurrentOp = $Msg
                        $Status = "Match user SID to full name in registry"
                        Write-Progress -Activity $Activity -CurrentOp $CurrentOp -Status $Status -PercentCopmplete ($Current/$Total*100)

                        # Look up path in registry/Group Policy by SID
                        $SID = $User.sid
                        $DN = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\$SID" -ErrorAction SilentlyContinue | Select -ExpandProperty Distinguished-Name)
                
                        # Look up the username by SID
                        $Username = (Get-WMIObject -Class Win32_UserAccount -Filter "sid = '$SID'" -ErrorAction Stop).name


                        Get-WmiObject Win32_UserAccount -Filter "Domain='$($env:ComputerName)' and Name='$Username'"

                    
                        If ($DN -match "(CN=)(.*?),.*") {
                            $FullName = $Matches[2]
                            $Output = New-Object PSObject -Property ([ordered] @{
                                ComputerName      = $Env:ComputerName
                                Username          = $UserName
                                FullName          = $FullName
                                ProfilePath       = $User.LocalPath
                                SID               = $SID
                                DistinguishedName = $($DN)
                                Messages          = $Null
                            })
                            $Results += $Output
                        }
                        Else {
                            $Msg = "Failed to find user registry match for '$($User.LocalPath)'"
                            $Host.UI.WriteErrorLine("ERROR: $Msg")

                            $Output = New-Object PSObject -Property ([ordered] @{
                                ComputerName      = $Env:ComputerName
                                Username          = "Error"
                                FullName          = "Error"
                                ProfilePath       = $User.LocalPath
                                SID               = "Error"
                                DistinguishedName = "Error"
                                Messages          = $Msg
                            })
                            $Results += $Output
                        } 
                    } #end foreach
        
        
                
                }
                Current {}
                All {}
            
            }
        
        
        
        
        
        
        }

        If ($AllUsers.IsPresent) {
            
            Foreach ($Profile in $SIDLocalUsers) {
                
                $Msg = $Profile.LocalPath
                Write-Verbose $Msg

                # Look up path in registry/Group Policy by SID
                $SID = $Profile.sid
                $DN = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\$SID" -ErrorAction SilentlyContinue | Select -ExpandProperty Distinguished-Name)
                
                # Look up the username by SID
                $Username = (Get-WMIObject -Class Win32_UserAccount -Filter "sid = '$SID'" -ErrorAction Stop).name
                    
                If ($DN -match "(CN=)(.*?),.*") {
                    $FullName = $Matches[2]
                    $Output = New-Object PSObject -Property ([ordered] @{
                        ComputerName      = $Env:ComputerName
                        Username          = $UserName
                        FullName          = $FullName
                        ProfilePath       = $Profile.LocalPath
                        SID               = $SID
                        DistinguishedName = $($DN)
                        Messages          = $Null
                    })
                    $Results += $Output
                }
                Else {
                    $Msg = "Failed to find user registry match for '$($Profile.LocalPath)'"
                    $Host.UI.WriteErrorLine("ERROR: $Msg")

                    $Output = New-Object PSObject -Property ([ordered] @{
                        ComputerName      = $Env:ComputerName
                        Username          = "Error"
                        FullName          = "Error"
                        ProfilePath       = $Profile.LocalPath
                        SID               = "Error"
                        DistinguishedName = "Error"
                        Messages          = $Msg
                    })
                    $Results += $Output
                } 
            } #end if found in profile path
        
        } #end if geting for all profiles
        
        Else {
            
            $UserName = (Get-WMIObject -class Win32_ComputerSystem -Property UserName -ErrorAction Stop).UserName
            $UserOnly = $UserName.Split("\")[1]

            $Msg = $UserName
            Write-Verbose $Msg

            Foreach ($Profile in $SIDLocalUsers) {
            
                # Match profile to current user
                If ($Profile.localpath -like "*$UserOnly"){

                    # Look up path in registry/Group Policy by SID
                    $SID = $Profile.sid
                    $DN = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\$SID" -ErrorAction SilentlyContinue | Select -ExpandProperty Distinguished-Name)
                    
                    If ($DN -match "(CN=)(.*?),.*") {
                        $FullName = $Matches[2]
                        $Output = New-Object PSObject -Property ([ordered] @{
                            ComputerName      = $Env:ComputerName
                            Username          = $UserName
                            FullName          = $FullName
                            ProfilePath       = $Profile.LocalPath
                            SID               = $SID
                            DistinguishedName = $($DN)
                            Messages          = $Null
                        })
                        $Results += $Output
                    }
                } #end if found in profile path
            } #end for each profile

            If ($Results.Count -eq 0) {
                $Msg = "No fullname data found"
                $Host.UI.WriteErrorLine("ERROR: $Msg")
                $Output = New-Object PSObject -Property ([ordered] @{
                    ComputerName      = $Env:ComputerName
                    Username          = $UserName
                    FullName          = $FullName
                    ProfilePath       = $Profile.LocalPath
                    SID               = $SID
                    DistinguishedName = $($DN)
                    Messages          = $Msg
                })
                $Results += $Output
            }

        } #end if current user only
        
    }
    Catch {
        
        $Msg = "Operation failed"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
        $Host.UI.WriteErrorLine("ERROR: $Msg")
        
        $Output = New-Object PSObject -Property ([ordered] @{
            ComputerName      = $Env:ComputerName
            Username          = $UserName
            FullName          = "Error"
            ProfilePath       = $Profile.LocalPath
            DistinguishedName = "Error"
            Messages          = $Msg
        })
        $Results += $Output
    }
    Finally {
        Write-Progress -Activity * -Completed
    }

    Write-Output $Results
}

} #end Get-PKWindowsFullUserName