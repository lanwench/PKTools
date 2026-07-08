#requires -Version 4
Function Get-PKTaniumClient {
    <#
    .SYNOPSIS 
        Gets the Tanium Client service and registry configuration from one or more computers, using Get-WMIObject for downlevel compatibility

    .DESCRIPTION
        Retrieves Tanium Client service state and registry configuration using Get-WmiObject for downlevel compatibility
        Defaults to the local computer; uses Invoke-Command for remote registry queries
        Accepts pipeline input
        Returns a PSObject
        Windows only

    .NOTES
        Name    : Function_Get-PKTaniumClient.ps1
        Created : 2024-02-12
        Author  : Paula Kingsley
        Version : 01.03
        History :

            ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

            v01.00 - 2024-02-12 - Created script
            v01.03 - 2026-06-17 - Cosmetic updates; converted version format to major.minor


    .PARAMETER ComputerName
        One or more computer names or FQDNs (if not specified, uses local computer)
    
    .PARAMETER Credential
        Valid credentials on target computers (incompatible with local computer)

    .EXAMPLE
        PS > Get-PKTaniumClient -Verbose

            VERBOSE: PSBoundParameters: 

            Key           Value
            ---           -----
            Verbose       True
            ComputerName  
            Credential    System.Management.Automation.PSCredential
            ScriptName    Get-PKTaniumClient
            ScriptVersion 01.03
            PipelineInput False


            VERBOSE: [BEGIN: Get-PKTaniumClient] Get Tanium Client service and settings
            VERBOSE: [LAPTOP17] Getting Tanium Client service
            VERBOSE: [LAPTOP17] Getting Tanium Client details from registry

            ComputerName       : LAPTOP17                                                                                        
            ServiceName        : Tanium Client                                                                                      
            ServiceState       : Running                                                                                            
            ServiceStartMode   : Auto                                                                                               
            ServicePath        : "C:\Program Files (x86)\Tanium\Tanium Client\TaniumClient.exe" --service                           
            Version            : 7.4.10.1075                                                                                        
            ServerNameList     : tanium-app-prod.internal.domain,com,public-tanium.domain.com
            ServerPort         : 17472                                                                                              
            LastGoodServerName : public-tanium.domain.com                                                                         
            Tags               : {PATCHGRP8,Compliance}
            FirstInstall       : 05/10/2022 13:02:57                                                                                
            ErrorMessage       :                                                                                                    
                                                                                                                                    
            VERBOSE: [END: Get-PKTaniumClient]  
    #>
    [CmdletBinding(DefaultParameterSetName = "Local")]
    Param(
        [Parameter(
            ParameterSetName = "Named",
            Position = 0,
            ValueFromPipeline,
            ValueFromPipelineByPropertyName,
            HelpMessage = "One or more computer names or FQDNs (if not specified, uses local computer)"
        )]
        [ValidateNotNullOrEmpty()]
        [string]$ComputerName,

        [Parameter(
            ParameterSetName = "Named",
            HelpMessage = "Valid credentials on target computers (incompatible with local computer)"
        )]
        [pscredential]$Credential = [System.Management.Automation.PSCredential]::Empty
    
    )
    Begin {
    
        # Current version (please keep up to date from comment block)
        [version]$Version = "01.03"

        # How did we get here?
        $ScriptName = $MyInvocation.MyCommand.Name
        $Source = $PSCmdlet.ParameterSetName
        [switch]$PipelineInput = $MyInvocation.ExpectingInput
        $CurrentParams = $PSBoundParameters
        $MyInvocation.MyCommand.Parameters.keys | Where-Object { $CurrentParams.keys -notContains $_ } | 
        Where-Object { Test-Path Variable:$_ } | Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
        $CurrentParams.Add("ParameterSetName", $Source)
        $CurrentParams.Add("PipelineInput", $PipelineInput)
        $CurrentParams.Add("ScriptName", $ScriptName)
        $CurrentParams.Add("ScriptVersion", $Version)
        Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

        If ($IsWindows -eq $False) {
            $Msg = "This function requires Windows"
            Write-Error $Msg
            Return
        }

        $Param = @{
            ErrorAction = "Stop"
            Verbose = $False
        }

        $Activity = "Get Tanium Client service and settings" 
        Write-Verbose "[BEGIN: $ScriptName] $Activity"

    }
    
    Process {
        Switch ($PSCmdlet.ParameterSetName) {
            Local {
                $Computer = $Env:COMPUTERNAME
                Try {
                    $Msg = "Getting Tanium Client service"
                    Write-Verbose "[$Computer] $Msg"
                    Write-Progress -Activity $Activity -CurrentOperation $Msg -Status $Computer
                    #$Service = Get-CIMInstance -Query "select * FROM win32_service WHERE name LIKE '%tanium client%'" @Param 
                    $Service = Get-WMIObject -Query "select * FROM win32_service WHERE name LIKE '%tanium client%'" @Param 
                    Try {
                        $Msg = "Getting Tanium Client details from registry"
                        Write-Verbose "[$Computer] $Msg"
                        Write-Progress -Activity $Activity -CurrentOperation $Msg -Status $Computer
                        $Props = "Path,Version,ServerNameList,LastGoodServerName,ServerPort,FirstInstall,ComputerID" -split(",")
                        $Reg = Get-ItemProperty "HKLM:\Software\Wow6432node\Tanium\Tanium Client" @Param | Select-Object $Props
                        $Tags =  Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Tanium\Tanium Client\Sensor Data\Tags"  @Param  
                        [string[]]$CustomTags = ($Tags | Get-Member -MemberType NoteProperty | Where-Object {$_.Name -notmatch "^PS"}).Name
                        $ErrorMessage = $Null
                    }
                    Catch {
                        $ErrorMessage = "Unable to get registry data ($($_.Exception.Message))"
                        Write-Warning "[$Computer] $ErrorMessage"
                    }  
                }
                Catch {
                    $ErrorMessage = "Unable to get Tanium Client service ($($_.Exception.Message))"
                    Write-Warning "[$Computer] $ErrorMessage"
                }
                Finally {
                    $Service | Select-Object  @{N="ComputerName";E={$Computer}},
                        @{N="ServiceName";E={$_.Name}},
                        @{N="ServiceState";E={$_.State}},
                        @{N="ServiceStartMode";E={$_.StartMode}},
                        @{N="ServicePath";E={$_.PathName}},
                        @{N="Version";E={$Reg.Version.ToString()}},
                        @{N="ServerNameList";E={$Reg.ServerNameList}},
                        @{N="ServerPort";E={$Reg.ServerPort}},
                        @{N="LastGoodServerName";E={$Reg.LastGoodServerName}},
                        @{N="Tags";E={$CustomTags}},
                        @{N="FirstInstall";E={$Reg.FirstInstall}},
                        @{N="ErrorMessage";E={$ErrorMessage}}
                }
            }
            Named {
                If ($PSBoundParameters.ContainsValue($Credential)) {$Param["Credential"] = $Credential}
                $Total = $ComputerName.Count
                $Current = 0
                Foreach ($Computer in $ComputerName) {
                    $Computer = $Computer.Trim()
                    $Current ++
                    Try {
                        $Msg = "Getting Tanium Client service"
                        Write-Verbose "[$Computer] $Msg"
                        Write-Progress -Activity $Activity -CurrentOperation $Msg -Status $Computer -PercentComplete ($Current/$Total*100)
                        #$sop = New-CimSessionOption -UseSsl
                        #$CimSession = New-CIMSession -ComputerName $Computer -Authentication Negotiate -SessionOption $Sop @Param
                        #$Service = Get-CIMInstance -CimSession $CimSession -Query "select * FROM win32_service WHERE name LIKE '%tanium client%'" @Param 
                        $Service = Get-WMIObject -ComputerName $Computer -Query "select * FROM win32_service WHERE name LIKE '%tanium client%'" @Param 
                        Try {
                            $Msg = "Getting Tanium Client details from registry"
                            Write-Verbose "[$Computer] $Msg"
                            Write-Progress -Activity $Activity -CurrentOperation $Msg -Status $Computer -PercentComplete ($Current/$Total*100)
                            $ScriptBlock = {
                                $Props = "Path,Version,ServerNameList,LastGoodServerName,ServerPort,FirstInstall,ComputerID" -split(",")
                                $Client = Get-ItemProperty "HKLM:\Software\Wow6432node\Tanium\Tanium Client" -ErrorAction Stop | Select-Object $Props
                                $Tags =  Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Tanium\Tanium Client\Sensor Data\Tags"  
                                [string[]]$CustomTags = ($Tags | Get-Member -MemberType NoteProperty | Where-Object {$_.Name -notmatch "^PS"}).Name
                                [PsCustomObject]@{RegSvc = $Client;RegTags = $CustomTags}
                            }
                            $Reg = Invoke-Command -ComputerName $Computer -ScriptBlock $ScriptBlock -Authentication Negotiate @Param
                            $ErrorMessage = $Null
                        }
                        Catch {
                            $ErrorMessage = "Unable to get registry data ($($_.Exception.Message))"
                            Write-Warning "[$Computer] $ErrorMessage"
                        }
                    }
                    Catch {
                        $ErrorMessage = "Unable to get Tanium Client service ($($_.Exception.Message))"
                        Write-Warning "[$Computer] $ErrorMessage"
                    }
                    Finally {
                        $CimSession | Remove-CimSession -ErrorAction SilentlyContinue
                        $Service | Select-Object  @{N="ComputerName";E={$Computer}},
                            @{N="ServiceName";E={$_.Name}},
                            @{N="ServiceState";E={$_.State}},
                            @{N="ServiceStartMode";E={$_.StartMode}},
                            @{N="ServicePath";E={$_.PathName}},
                            @{N="Version";E={$Reg.RegSvc.Version.ToString()}},
                            @{N="ServerNameList";E={$Reg.RegSvc.ServerNameList}},
                            @{N="ServerPort";E={$Reg.RegSvc.ServerPort}},
                            @{N="LastGoodServerName";E={$Reg.RegSvc.LastGoodServerName}},
                            @{N="Tags";E={$Reg.RegTags}},
                            @{N="FirstInstall";E={$Reg.RegSvc.FirstInstall}},
                            @{N="ErrorMessage";E={$ErrorMessage}}
                    }
                }
            }
        } #end switch
    }
    End {
        Write-Progress -Activity * -Completed
        Write-Verbose "[END: $ScriptName] Script ran successfully"
    }
} #end Get-PSGitStatus
