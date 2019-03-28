#requires -Version 3
 Function Get-PKLocalGroupMember {
 <#
 .SYNOPSIS
    Uses 'net localgroup' to return members of a named local group or all groups

.DESCRIPTION
    Uses 'net localgroup' to return members of a named local group or all groups
    Returns a PSobject

.NOTES        
    Name    : Function_Get-PKLocalGroupMember.ps1
    Created : 2019-03-26
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2019-03-26 - Created script

.EXAMPLE
    PS C:\> Get-PKLocalGroupMember Administrators -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value                 
        ---           -----                 
        Verbose       True                  
        Name          {Administrators}      
        All           False                 
        Quiet         False                 
        ScriptName    Get-PKLocalGroupMember
        ScriptVersion 1.0.0                 
        
        BEGIN  : Get local group membership

        VERBOSE: [PKINGSLEY-05122] Getting membership for group 'Administrators'
        VERBOSE: [PKINGSLEY-05122] Found group 'Administrators'
        VERBOSE: [PKINGSLEY-05122] Group 'Administrators' has 12 direct member(s)

        END    : Get local group membership
        Computername    Group          Members                                                                                                       
        ------------    -----          -------                                                                                                       
        PKINGSLEY-05122 Administrators {Administrator, DOMAIN\Domain Admins, DOMAIN\Workstation Admins, DOMAIN\IT Support gr...

.EXAMPLE
    PS C:\> Get-PKLocalGroupMember -Quiet

        Computername    Group                          Members                                                                                 
        ------------    -----                          -------                                                                                 
        PKINGSLEY-05122 Backup Operators                                                                                                       
        PKINGSLEY-05122 ConfigMgr Remote Control Users {DOMAIN\helpdesk, DOMAIN\plevy, DOMAIN\tsmith...}
        PKINGSLEY-05122 Cryptographic Operators                                                                                                
        PKINGSLEY-05122 Distributed COM Users                                                                                                  
        PKINGSLEY-05122 Event Log Readers              NT AUTHORITY\NETWORK SERVICE                                                            
        PKINGSLEY-05122 Guests                         Guest                                                                                   
        PKINGSLEY-05122 Hyper-V Administrators                                                                                                 
        PKINGSLEY-05122 IIS_IUSRS                                                                                                              
        PKINGSLEY-05122 Network Configuration Opera...                                                                                         
        PKINGSLEY-05122 Offer Remote Assistance Hel... {DOMAIN\helpdesk, DOMAIN\plevy, DOMAIN\tsmith...}
        PKINGSLEY-05122 Performance Log Users                                                                                                  
        PKINGSLEY-05122 Performance Monitor Users                                                                                              
        PKINGSLEY-05122 Power Users                                                                                                            
        PKINGSLEY-05122 Remote Desktop Users           {DOMAIN\jbloggs, DOMAIN\jbloggs, OLDCORP\kzuecher}                      
        PKINGSLEY-05122 Remote Management Users                                                                                                
        PKINGSLEY-05122 Replicator                                                                                                             
        PKINGSLEY-05122 System Managed Accounts Group  DefaultAccount                                                                          
        PKINGSLEY-05122 Users                          {DOMAIN\Domain Users, NT AUTHORITY\Authenticated Users, NT AUTHORITY\INTERACTIVE}   


 
 
 #>
 [CmdletBinding(DefaultParameterSetName="All")]
 Param(
    [Parameter(
        ParameterSetName = "All",
        HelpMessage = "Return group membership for all local groups"
    )]
    [switch]$All,

    [Parameter(
        ParameterSetName = "Named",
        Position = 0,
        HelpMessage = "One or more local group names"
    )]
    [ValidateNotNullOrEmpty()]
    [string[]]$Name,

    [Parameter(
        HelpMessage = "Hide all non-verbose/non-error console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [Switch] $Quiet
 )
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    #$Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = (-not $PSBoundParameters.ContainsKey("ComputerName")) -and (-not $ComputerName)
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    #$CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )" 

    If (-not $CurrentParams.Name) {[Switch]$All = $True}

    # Output
    $Output = @()

    # Console output
    $Activity = "Get local group membership"
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "BEGIN  : $Activity"
    $FGColor = "Yellow"
    If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"`n$Msg`n")}
    Else {Write-Verbose $Msg}

 
}
Process {

    If ($All.IsPresent) {
        
        $Msg = "Get all local groups"
        Write-Verbose "[$Env:ComputerName] $Msg"

        Write-Progress -Status $Env:ComputerName -Activity $Activity -CurrentOperation $Msg 

        $Cmd = "net localgroup"
        $Groups = ((Invoke-Expression $Cmd | Where-Object {$_ -AND ($_ -notmatch "Aliases for \\$Env:ComputerName | Alias | Comment")} | Select-Object -Skip 4) | Where-Object {$_ -notmatch "The command completed successfully."}).Replace("*",$Null)

        If ($Groups) {
            $Msg = "Found $($Groups.Count) local group(s)"
            Write-Verbose "[$Env:ComputerName] $Msg"
        }

    }
    Else {
        $Groups = $Name
    }

    $Total = $Groups.Count
    $Current = 0
    Foreach ($Group in $Groups) {
        
        $Current ++ 
        Write-Progress -Status $Env:Computername -Activity $Activity -CurrentOperation "$Group" -PercentComplete ($Current/$Total*100)

        $Msg = "Getting membership for group '$Group'"
        Write-Verbose "[$Env:ComputerName] $Msg"

        Try {
            $Cmd = "net localgroup '$Group'"
            Try {
                Function IsGroup($Grp) {
                    Try {
                        (net localgroup "$Grp") 2>&1
                    }
                    Catch {}
                }

                $Results = (Invoke-Expression $Cmd)
                }
                Catch {}#$ErrorDetails = $_.Exception.Message}
            If ($Results) {
        
                $Msg = "Found group '$Group'"
                Write-Verbose "[$Env:ComputerName] $Msg"

                $Members = ($Results | Where-Object {$_ -AND ($_ -notmatch "Aliases for \\$Env:ComputerName | Alias | Comment")} | Select-Object -Skip 4).Replace("*",$Null) | Where-Object {$_ -notmatch "The command completed successfully."}

                $Msg = "Group '$Group' has $($Members.Count) direct member(s)"
                Write-Verbose "[$Env:ComputerName] $Msg"

                $Results = New-Object PSObject -Property ([ordered] @{
                    Computername = $env:COMPUTERNAME
                    Group        = $Group
                    Members      = $Members
                })

                #$Output += $Results
                Write-Output $Results
            }
            Else {
                $Msg = "No results found for group '$Group'"
                $Host.UI.WriteErrorLine("[$Env:ComputerName] $Msg")
            }
        }
        Catch {
            $Msg = "Operation failed for group '$Group'"
            $Host.UI.WriteErrorLine("[$Env:ComputerName] $Msg")
        }
        
    } #end for each group

    #Write-Output $Output
    
}
End {
    
    $Msg = "END    : $Activity"
    $FGColor = "Yellow"
    If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"`n$Msg")}
    Else {Write-Verbose $Msg}
}    
}
 
