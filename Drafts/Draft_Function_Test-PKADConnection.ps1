 Function Test-PKADConnection {
        [Cmdletbinding()]
        Param(
             
            [Parameter(
                Mandatory = $False,
                HelpMessage = "Active directory domain (default is current"
            )]
            [ValidateNotNullOrEmpty()]
            [Alias("Name")]
            [string]$ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name,
    
            [Parameter(
                Mandatory = $False,
                HelpMessage = "Domain controller FQDN (default is to use first available in site)"
            )]
            [ValidateNotNullOrEmpty()]
            [Alias("DomainController")]
            [string]$Server = ([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()).FindDomainController().Name,

            [Parameter(
                Mandatory = $False,
                HelpMessage = "Credentials"
            )]
            [ValidateNotNullOrEmpty()]
            [PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty
        )
        Begin {

            [Version]$Version = "01.00.0000"

            # Display our settings
            $ScriptName = $MyInvocation.MyCommand.Name
            $CurrentParams = $PSBoundParameters
            $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
                Where {Test-Path variable:$_}| Foreach {
                    $CurrentParams.Add($_, (Get-Variable $_).value)
                }
            $CurrentParams.Add("ScriptName",$ScriptName)
            $CurrentParams.Add("ScriptVersion",$Version)
            Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"
            
            $ErrorActionPreference = "SilentlyContinue"
        }
        Process {
            Try {
                
                [switch]$Continue = $False
                If ($CurrentParams.ADDomain -ne [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name) {
                    $Msg = "Verify AD domain"
                    Write-Verbose $Msg
                    $Context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("domain", $ADDomain)
                    $DomainObj = [system.directoryservices.activedirectory.Domain]::GetDomain($context)
                    If (-not $DomainObj) {
                        $Msg = "Invalid domain '$ADDomain'"
                        $Host.UI.WriteErrorLine($Msg)
                    }
                    Else {
                        $Msg = "Found AD domain '$ADDomain'"
                        Write-Verbose $Msg
                        $Continue = $True
                    }
                }
                Else {
                    $DomainObj = [DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()                        
                    $ADDomain = $DomainObj.Name
                    $Msg = "Use current AD domain '$ADDomain'"
                    Write-Verbose $Msg
                    $Continue = $True
                }

                If ($Continue.IsPresent) {
                    
                    $Continue = $False

                    If ($CurrentParams.Server -ne ([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()).FindDomainController().Name) {
                        $Msg = "Verify domain controller"
                        Write-Verbose $Msg
                        If ($ServerObj = $DomainObj.FindAllDomainControllers().Name  | Where-Object {$_ -match $Server}) {
                            $Msg = "Found domain controller '$($ServerObj.Name)'"
                            Write-Verbose $Msg
                            $Continue = $True
                        }
                        Else {
                            $Msg = "Invalid domain controller '$Server'"
                            $Host.UI.WriteErrorLine($Msg)
                        }
                    }
                    Else {
                        $Msg = "Use nearest domain controller '$Server'"
                        Write-Verbose $Msg
                        $Server = $DomainObj.FindDomainController().Name
                        $Msg = "Found domain controller '$Server'"
                        $Continue = $True
                    }
                }

                If ($Continue.IsPresent) {
                    
                    Try {
                        If ($CurrentParams.Credential.UserName) {
                            $Msg = "Test authentication as '$($Credential.Username)'"
                            Write-Verbose $Msg
                            $TestAuth = New-Object DirectoryServices.DirectoryEntry("LDAP://$Server",$Credential.UserName, $Credential.GetNetworkCredential().Password)
                            If ($TestAuth) {$True}
                        }
                        Else {
                            $Msg = "Test authentication as current user"
                            Write-Verbose $Msg
                            $TestAuth = New-Object DirectoryServices.DirectoryEntry("LDAP://$Server")  
                            If ($TestAuth) {$True}
                        }
                    }
                    Catch {
                        $Msg = "Failed to authenticate to '$ADDomain' using server '$Server'"
                        If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
                        $Host.UI.WriteErrorLine("$Msg")
                    }
                }
            }
            Catch {
                $Msg = "Operation failed"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
                $Host.UI.WriteErrorLine("$Msg")
            }

        } #end process
    } #end function