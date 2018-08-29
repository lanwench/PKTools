#requires -Version 3
Function Open-PKWindowsMSC {
<#
.SYNOPSIS
    Launches an MMC snapin using current or alternate credentials

.DESCRIPTION
    Launches an MMC snapin using current or alternate credentials
    Separate parameters for local v domain tools
    Brings up a list/menu of available tools (dynamic parameters, for fun)
    Credential parameter available for domain tools only (also converts
    UPN format to DOMAIN\username)
    
.NOTES
    Name    : Function_Open-PKWindowsMSC
    Created : 2017-10-16
    Version : 01.00.0000
    Author  : Paula Kingsley
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK

        v01.00.0000 - 2017-10-16 - Created script


.EXAMPLE
    PS C:\> Open-PKWindowsMSC -Local EventViewer -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value            
        ---                   -----            
        Verbose               True             
        Local                 EventViewer      
        Credential                             
        SuppressConsoleOutput False            
        ParameterSetName      Local            
        ScriptName            Open-PKWindowsMSC
        ScriptVersion         1.0.0            

        Launching Event Viewer (eventvwr.msc)
        

.EXAMPLE
    PS C:\> Open-PKWindowsMSC -Domain DNSManagement -Credential $AdminCredential -SuppressConsoleOutput -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                                    
        ---                   -----                                    
        Credential            System.Management.Automation.PSCredential
        SuppressConsoleOutput True                                     
        Verbose               True                                     
        Domain                DNSManagement                            
        ParameterSetName      Domain                                   
        ScriptName            Open-PKWindowsMSC                        
        ScriptVersion         1.0.0                                    


        VERBOSE: Launching DNS Management (dnsmgmt.msc) as DOMAIN\jbloggs-admin



#>
[CmdletBinding(DefaultParameterSetName = "Local")]
Param(
    
    [Parameter(
        Mandatory = $True,
        Position = 0,
        HelpMessage = "MSC name"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateSet('ADDomainsAndTrusts','ADSIEdit','ADUsersAndComputers','CertificationAuthority','ComputerManagement','DeviceManager','DHCPManager','DiskManagement','DNSManagement','EventViewer','GroupPolicyManagement','HyperVManager','IISManager','LocalGroupPolicyEditor','LocalSecurityPolicy','LocalUsersAndGroups','PerformanceMonitor','ServerManager','Services','SharedFolders','TaskScheduler','WSUSManager')]
    [string]$Name,

    [Parameter(
        ParameterSetName = "Domain",
        Mandatory = $False,
        Position = 1,
        HelpMessage = "Credentials (applicable to domain tools only)"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential] $Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(
        #ParameterSetName='Domain',
        Position = 2,
        Mandatory = $False
    )] 
    [ValidateNotNullOrEmpty()]    
    [string]$ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name,

    [Parameter(
        #ParameterSetName='Domain',
        Position = 2,
        Mandatory = $False
    )] 
    [ValidateNotNullOrEmpty()]
    [string]$Server,

    [Parameter(
        Mandatory   = $False,
        Position = 4,
        HelpMessage = "Suppress all non-verbose / non-error console output"
    )]
    [switch] $SuppressConsoleOutput

)

Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "02.00.0000"

    # How did we get here
    $PipelineInput = (-not $PSBoundParameters.ContainsKey("ComputerName")) -and (-not $ComputerName)

    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preference
    $ErrorActionPreference = "Stop"

#region OLD
    <#

    # PSObject for selecting filename/descriptive name from selection
    [pscustomobject]$LookupObj = @'
    "Name","Desc","Target","File"
    "ADDomainsAndTrusts","Active Directory Domains and Trusts","Domain","domain.msc"
    "ADSIEdit","ADSI Edit","Domain","adsiedit.msc"
    "ADUsersAndComputers","Active Directory Users and Computers","Domain","dsa.msc"
    "CertificationAuthority","Certification Authority","Local","certsrv.msc"
    "ComputerManagement","Computer Management","Local","compmgmt.msc"
    "DeviceManager","Device Manager","Local","devmgmt.msc"
    "DHCPManager","DHCP Manager","Domain","dhcpmgmt.msc"
    "DiskManagement","Disk Management","Local","diskmgmt.msc"
    "DNSManagement","DNS Management","Domain","dnsmgmt.msc"
    "EventViewer","Event Viewer","Local","eventvwr.msc"
    "GroupPolicyManagement","Group Policy Management","Domain","gpmc.msc"
    "HyperVManager","Hyper-V Manager","Local","virtmgmt.msc"  
    "IISManager","IIS Manager","Local","iis.msc"
    "LocalGroupPolicyEditor","Local Group Policy Editor","Local","gpedit.msc"    
    "LocalSecurityPolicy","Local Security Policy","Local","secpol.msc"
    "LocalUsersAndGroups","Local Users and Groups","Local","lusrmgr.msc"    
    "PerformanceMonitor","Performance Monitor","Local","perfmon.msc"  
    "ServerManager","Server Manager","Domain","ServerManager.msc"  
    "Services","Services","Local","Services.msc"
    "SharedFolders","Shared Folders","Local","fsmgmt.msc"
    "TaskScheduler","Task Scheduler","Local","taskschd.msc"
    "WSUSManager","WSUS Manager","Domain","wsus.msc"
'@ | ConvertFrom-CSV -Delimiter "," 

    #>

<#

    # PSObject for selecting filename/descriptive name from selection
    [pscustomobject]$LookupObj = @'
    "Name","Desc","File"
    "ADDomainsAndTrusts","Active Directory Domains and Trusts","domain.msc"
    "ADSIEdit","ADSI Edit","adsiedit.msc"
    "ADUsersAndComputers","Active Directory Users and Computers","dsa.msc"
    "CertificationAuthority","Certification Authority","certsrv.msc"
    "ComputerManagement","Computer Management","compmgmt.msc"
    "DeviceManager","Device Manager","devmgmt.msc"
    "DHCPManager","DHCP Manager","dhcpmgmt.msc"
    "DiskManagement","Disk Management","diskmgmt.msc"
    "DNSManagement","DNS Management","dnsmgmt.msc"
    "EventViewer","Event Viewer","eventvwr.msc"
    "GroupPolicyManagement","Group Policy Management","gpmc.msc"
    "HyperVManager","Hyper-V Manager","virtmgmt.msc"  
    "IISManager","IIS Manager","iis.msc","gpedit.msc"
    "LocalGroupPolicyEditor","Local Group Policy Editor","gpedit.msc"    
    "LocalSecurityPolicy","Local Security Policy","secpol.msc"
    "LocalUsersAndGroups","Local Users and Groups","lusrmgr.msc"    
    "PerformanceMonitor","Performance Monitor","perfmon.msc"  
    "ServerManager","Server Manager","ServerManager.msc"  
    "Services","Services","Services.msc"
    "SharedFolders","Shared Folders","fsmgmt.msc"
    "TaskScheduler","Task Scheduler","taskschd.msc"
    "WSUSManager","WSUS Manager","wsus.msc"
'@ | ConvertFrom-CSV -Delimiter "," 

#>

<#    
    # Look up filename, description based on entry
    $Selection = $LookupObj | Where-Object {$_.Name -eq $Name}
    

    # Start creating string for argument list
    [string]$ArgList = "/c $($Selection.File)"

    # If domain tools, check domain/server, normalize credentials if provided
    Switch ($Selection.Target) {
        Local {
            $Msg = "Credentials, domain name, and server name will be ignored for local MMCs"
            Write-Warning $Msg
        }
        Domain {
            
            # Get the domain context (either default, current user ... or provided)
            $Msg = "Get domain context for '$ADDomain'"
            Write-Verbose $Msg
            Try {
                $Context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("domain",$ADDomain)
                $DomainObj = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($Context)

                # Now add to string
                $ArgList += " /Domain=$ADDomain"
            }
            Catch {
                $Msg = "Failed to get domain context for '$ADDomain'"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$($_.Exception.Message)"}
                $Host.UI.WriteErrorLine("$Msg")
            }
            
            # !!!! Need to figure out how to make sure this is in the correct domain
            If ($CurrentParams.Server) {
                $Msg = "Get domain controller '$Server'"
                Write-Verbose $Msg
                Try {
                    $Type = [System.DirectoryServices.ActiveDirectory.DirectoryContextType]::DirectoryServer
                    $Context = New-Object -TypeName System.DirectoryServices.ActiveDirectory.DirectoryContext -ArgumentList $type, $Server
                    $Server = [System.DirectoryServices.ActiveDirectory.DomainController]::GetDomainController($context).Name
                }
                Catch {
                    $Msg = "Failed to get domain controller '$Server'"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$($_.Exception.Message)"}
                    $Host.UI.WriteErrorLine("$Msg")
                }
            }
            Else {
                $Msg = "Get domain controller in '$ADDomain'"
                Write-Verbose $Msg
                Try {
                    $context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("domain",$ADDomain)
                    $Server = [System.DirectoryServices.ActiveDirectory.DomainController]::findone($Context).Name
                }
                Catch {
                    $Msg = "Failed to get domain controller '$Server'"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$($_.Exception.Message)"}
                    $Host.UI.WriteErrorLine("$Msg")
                }
            }

            # Add server to arguments
            $ArgList += " /SERVER=$Server"
            
            If ($CurrentParams.Credential.UserName) {
            
                $Param_Launch.Add("Credential",$Credential)
                [switch]$GetNetBIOSName = $False

                # Standardize username as DOMAIN\user
                If ($Credential.UserName -match "@") {
                    #$UserName = "$NetBIOSName\$(($Credential.UserName -split "@")[0])"
                    $UserName = $($Credential.UserName -split "@")[0]
                    $GetNetBIOSName = $True
                } 
                Elseif ($Credential.Username -notmatch '^\w+[\\]\w+$') {
                    #$Username = "$NetBIOSName\$($Credential.Username)"
                    $Username = $Credential.Username
                    $GetNetBIOSName = $True
                }
                
                If ($GetNetBIOSName.IsPresent) {
                # Get NetBIOS name of domain (need this for credential string in mmc.exe)
                $Msg = "Get NetBIOS name of domain"
                Write-Verbose $Msg
                Try {
                    $DomainEntry = New-Object -TypeName System.DirectoryServices.DirectoryEntry "LDAP://$Server",$($Credential.UserName),$($Credential.GetNetworkCredential().password)
                    $NetBIOSName = $DomainEntry.name
                    $Username = "$NetBIOSName\$UserName"
                }
                Catch {
                    $Msg = "Failed to get NetBIOS name of domain '$ADDomain'"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$($_.Exception.Message)"}
                    $Host.UI.WriteErrorLine("$Msg")
                }
            }
            

                # Now add to string
                $ArgList += " /Domain=$ADDomain"
            }


                $Activity = "$Activity as $UserName"

            }


        }
    }

    # Start creating string for argument list
    [string]$ArgList = "/c $($Selection.File)"



    # start ""/B MMC.exe /A $($Selection.File) /DOMAIN="$ADDomain" /SERVER="%strDC%.%strDomain%"

    #region Splats

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose = $False
    }

    # Splat for write-progress
    $Activity = "Launching $($Selection.Target.ToLower()) MMC '$($Selection.Desc)' ($($Selection.File))"
    
    $Param_WP = @{}
    $Param_WP = @{
        Activity = $Null
        Status   = "Working"
    }
    
    # Splat for Start-Process
    $Param_Launch = @{}
    $Param_Launch = @{
        FilePath         = "$Env:WinDir\System32\cmd.exe"
        ArgumentList     = $Null
        WorkingDirectory = $PSHome
        NoNewWindow      = $True
        ErrorAction      = "Stop"
        Verbose          = $False
    }
    If ($CurrentParams.Credential) {
        
        Switch ($Selection.Target) {

            Local {
                $Msg = "Credentials, domain name, and server name will be ignored for local MMCs"
                Write-Warning $Msg
            }
            Domain {

                # UPNs don't work. 
                If ($Credential.UserName -match "@") {

                    # Get the username part of the UPN
                    $UserName = ($Credential.UserName -split "@")[0]

                    # Get AD Root
                    $RootDSE = [ADSI]"LDAP://RootDSE"
                    $RootDSEConfig = $RootDSE.Get("configurationNamingContext")
                    $ADSearchRoot=New-object System.DirectoryServices.DirectoryEntry("LDAP://CN=Partitions," + $RootDSEConfig)

                    # Search for Netbiosname of the specified domain
                    #$DomainFQDN = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name
                    $DomainFQDN = $ADDomain
                    $SearchString = "(&(objectclass=Crossref)(dnsRoot="+$DomainFQDN+")(netBIOSName=*))"
                    $Search = New-Object directoryservices.DirectorySearcher($ADSearchRoot,$SearchString)
                    $sNetBIOSName = ($Search.FindOne()).Properties["netbiosname"]

                    # Set the username as domain\username
                    $UserName = "$sNetBIOSName\$UserName"

                    # Create a new object
                    $InternalCred = New-Object System.Management.Automation.PSCredential -ArgumentList @($UserName,$AdminCredential.Password)
            
                    # Add to the splat
                    $Param_Launch.Add("Credential",$InternalCred)
                }
                Else {
                    # Add to splat
                    $InternalCred = $Credential
                    $UserName = $UserName = $Credential.UserName
                    $Param_Launch.Add("Credential",$InternalCred)
                }
                $Activity = "$Activity as $UserName"
            }
        } #end switch
        
    } # end if credentials

    Switch ($Selection.Target) {
        Domain {
            $Activity += " in domain $ADDomain"
            $
        }
        Local {}
    }

    #endregion Splats


    #>


#endregion OLD

    #region functions to create and remove temp files

    function Local:New-TempFile {
        [cmdletbinding()]
        Param($CmdStr)
        #Generate Temporary File Name in c:\windows
        Try {
            $sTempFile = "$("$env:WinDir\Temp\")$(([System.IO.Path]::GetRandomFileName()).Replace(".",$Null)).bat"
            $CmdStr | Out-File -FilePath $sTempFile -Encoding ascii -Force -Confirm:$False -EA Stop
            Write-Output $sTempFile
        }
        Catch {
            $Host.UI.WriteErrorLine($_.Exception.Message)
        }   
    }

    
    function Local:Remove-TempFile {
        [cmdletbinding()]
        Param($File)
        #Generate Temporary File Name in c:\windows
        Try {
            Get-Item $File -Force -ErrorAction SilentlyContinue | Remove-Item -Force -Confirm:$False -ErrorAction Stop
        }
        Catch {
            $Host.UI.WriteErrorLine($_.Exception.Message)
        }   
    }


    #endregion functions to create and remove temp files


    # Splat for Write-Progress
    $Activity = "Run MSC as another user"
    $Param_WP = @{}
    $Param_WP = @{
        Activity = $Activity
        CurrentOperation = $Null
    }

    #region AD stuff

    $AdminUser = $Credential.UserName
    #$ADDomain = "enterprisenet.org"

    # Get the domain context (either default, current user ... or provided)
    $Msg = "Get domain context for '$ADDomain'"
    Write-Verbose $Msg
    
    $Param_WP.CurrentOperation = $Msg
    Write-Progress @Param_WP

    Try {
        $Context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("domain",$ADDomain)
        $DomainObj = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($Context)

        # Now add to string
        $ArgList += " /Domain=$ADDomain"
    }
    Catch {
        $Msg = "Failed to get domain context for '$ADDomain'"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$($_.Exception.Message)"}
        $Host.UI.WriteErrorLine("$Msg")
    }
    
    # !!!! Need to figure out how to make sure this is in the correct domain
    If ($CurrentParams.Server) {
        $Msg = "Get domain controller '$Server'"
        Write-Verbose $Msg
        
        $Param_WP.CurrentOperation = $Msg
        Write-Progress @Param_WP

        Try {
            $Type = [System.DirectoryServices.ActiveDirectory.DirectoryContextType]::DirectoryServer
            $Context = New-Object -TypeName System.DirectoryServices.ActiveDirectory.DirectoryContext -ArgumentList $type, $Server
            $Server = [System.DirectoryServices.ActiveDirectory.DomainController]::GetDomainController($context).Name
        }
        Catch {
            $Msg = "Failed to get domain controller '$Server'"
            If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$($_.Exception.Message)"}
            $Host.UI.WriteErrorLine("$Msg")
        }
    }
    Else {
        $Msg = "Get domain controller in '$ADDomain'"
        Write-Verbose $Msg
        
        $Param_WP.CurrentOperation = $Msg
        Write-Progress @Param_WP

        Try {
            $context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("domain",$ADDomain)
            $Server = [System.DirectoryServices.ActiveDirectory.DomainController]::findone($Context).Name
        }
        Catch {
            $Msg = "Failed to get domain controller '$Server'"
            If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$($_.Exception.Message)"}
            $Host.UI.WriteErrorLine("$Msg")
        }
    }


    #endregion AD stuff

    #region Lookup table for menu/matching

    # Create oSObject for selecting filename/descriptive name from selection
    [pscustomobject]$LookupObj = @'
        "Name","Desc","Target","File"
        "ADDomainsAndTrusts","Active Directory Domains and Trusts","Domain","domain.msc"
        "ADSIEdit","ADSI Edit","Domain","adsiedit.msc"
        "ADUsersAndComputers","Active Directory Users and Computers","Domain","dsa.msc"
        "CertificationAuthority","Certification Authority","Local","certsrv.msc"
        "ComputerManagement","Computer Management","Local","compmgmt.msc"
        "DeviceManager","Device Manager","Local","devmgmt.msc"
        "DHCPManager","DHCP Manager","Domain","dhcpmgmt.msc"
        "DiskManagement","Disk Management","Local","diskmgmt.msc"
        "DNSManagement","DNS Management","Domain","dnsmgmt.msc"
        "EventViewer","Event Viewer","Local","eventvwr.msc"
        "GroupPolicyManagement","Group Policy Management","Domain","gpmc.msc"
        "HyperVManager","Hyper-V Manager","Local","virtmgmt.msc"  
        "IISManager","IIS Manager","Local","iis.msc"
        "LocalGroupPolicyEditor","Local Group Policy Editor","Local","gpedit.msc"    
        "LocalSecurityPolicy","Local Security Policy","Local","secpol.msc"
        "LocalUsersAndGroups","Local Users and Groups","Local","lusrmgr.msc"    
        "PerformanceMonitor","Performance Monitor","Local","perfmon.msc"  
        "ServerManager","Server Manager","Domain","ServerManager.msc"  
        "Services","Services","Local","Services.msc"
        "SharedFolders","Shared Folders","Local","fsmgmt.msc"
        "TaskScheduler","Task Scheduler","Local","taskschd.msc"
        "WSUSManager","WSUS Manager","Domain","wsus.msc"
'@ | ConvertFrom-CSV -Delimiter "," 
    

    #endregion Lookup table for menu/matching

    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "Activity: $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg}
    
}
Process {

    If ($CurrentParams.Name) {
        
        $Msg = "Validate entry"
        Write-Verbose $Msg
        
        $Param_WP.CurrentOperation = $Msg
        Write-Progress @Param_WP

        If ($Selection = $LookupObj | Where-Object {$_.Name -eq $Name}) {
            $Msg = "Selected '$($Selection.Desc)'"
            Write-Verbose $Msg
        }
        Else {
            $Msg = "Invalid entry '$Name'; please omit and run again to select from menu"
            $Host.UI.WriteErrorLine("ERROR: $Msg")
        }
    }
    Else {
        $Msg = "Select name from menu"
        Write-Verbose $Msg
        
        $Param_WP.CurrentOperation = $Msg
        Write-Progress @Param_WP
        
        $Title = "Please make a selection to run an MSC"
        If ($CurrentParams.Credential.Username) {
            $Title += " as user '$AdminUser'"
        }
        If
        If ($Selection = $LookupObj | Out-GridView -Title "Please make a selection to run an MSC in the '$ADDomain' domain as user '$AdminUser'" -OutputMode Single) {
            $Msg = "Selected '$($Selection.Desc)'"
            Write-Verbose $Msg
        }
        Else {
            $Msg = "No selection made"
            $Host.UI.WriteErrorLine("$Msg")
        }
    }
    
    # Create string for expression
    If ($Selection.Target -eq "Local") {
        $Msg = "Username and domain will be ignored for local commands"
        Write-Warning $Msg
    }

    $CmdStr = "%SYSTEMROOT%\System32\runas.exe /netonly /user:$AdminUser ""%SystemRoot%\system32\mmc.exe %SystemRoot%\system32\$($Selection.File) /domain=$ADDomain /server=$Server"" "
    $Msg = "Created command string"
    Write-Verbose $Msg
    Write-Verbose $CmdStr



Try {
        $Msg = "Create temporary script file"
        Write-Verbose $Msg
        
        $Param_WP.CurrentOperation = $Msg
        Write-Progress @Param_WP

        # Create batch file 
        $TempScript = New-TempFile -CmdStr $CmdStr -ErrorAction Stop
        $Msg = "Created temporary script file $TempScript"
        Write-Verbose $Msg

        Try {
            $Msg = "Run script file"
            Write-Verbose $Msg
            
            $Param_WP.CurrentOperation = $Msg
            Write-Progress @Param_WP

            # Run it
            $Run = Powershell -command $TempScript -Verb runas
            
    
        }
        Catch {
            $Msg = "Failed to run script file"
            If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
            $Host.UI.WriteErrorLine("ERROR: $Msg")
        }
    }
    Catch {
        $Msg = "Failed to create script file"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
        $Host.UI.WriteErrorLine("ERROR: $Msg")

    }
    Finally {
        
        $Msg = "Remove temp file"
        $Param_WP.CurrentOperation = $Msg
        Write-Progress @Param_WP

        Remove-TempFile -File $TempScript
    }
  

}
End {
    Write-Progress @Param_WP -Completed
    Remove-TempFile -File $TempScript
}
} #end Open-PKWindowsMSC





<#
DynamicParam{

    # Set the dynamic parameters' name
    $ParamName_Local = 'Local'
    # Create the collection of attributes
    $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
    # Create and set the parameters' attributes
    $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
    $ParameterAttribute.Mandatory = $True
    $ParameterAttribute.Position = 1
    $ParameterAttribute.ParameterSetName = "Local"
    # Add the attributes to the attributes collection
    $AttributeCollection.Add($ParameterAttribute) 
    # Create the dictionary 
    $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
    # Generate and set the ValidateSet 
    $arrSet = "ComputerManagement","DeviceManager","DiskManagement","EventViewer","LocalGroupPolicyEditor","LocalSecurityPolicy","LocalUsersAndGroups","PerformanceMonitor","Services","SharedFolders","TaskScheduler"
    $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
    # Add the ValidateSet to the attributes collection
    $AttributeCollection.Add($ValidateSetAttribute)
    # Create and return the dynamic parameter
    $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ParamName_Local, [string], $AttributeCollection)
    $RuntimeParameterDictionary.Add($ParamName_Local, $RuntimeParameter)

    
    # Set the dynamic parameters' name
    $ParamName_Domain = 'Domain'
    # Create the collection of attributes
    $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
    # Create and set the parameters' attributes
    $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
    $ParameterAttribute.Mandatory = $True
    $ParameterAttribute.Position = 2
    $ParameterAttribute.ParameterSetName = "Domain"
    # Add the attributes to the attributes collection
    $AttributeCollection.Add($ParameterAttribute)  
    # Generate and set the ValidateSet 
    $ArrSet = "ADDomainsAndTrusts","ADSIEdit","ADUsersAndComputers","CertificationAuthority","DHCPManager","DNSManagement","GroupPolicyManagement","HyperVManager","IISManager","ServerManager","WSUSManager"
    $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
    # Add the ValidateSet to the attributes collection
    $AttributeCollection.Add($ValidateSetAttribute)
    # Create and return the dynamic parameter
    $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ParamName_Domain, [string], $AttributeCollection)
    $RuntimeParameterDictionary.Add($ParamName_Domain, $RuntimeParameter)
    return $RuntimeParameterDictionary

}
  
#>