#requires -Module VMware.VimAutomation.Core
#requires -Module ActiveDirectory
#requires -Version 3
Function Get-PKWindowsReport {
<#
.SYNOPSIS
    Gets AD, WMI, and WSUS details for a Windows computer

.DESCRIPTION
    Gets AD, WMI, and WSUS details for a Windows computer
    Returns a custom object
    Returns either stacked strings or collections for multivalue properties
    Accepts pipeline input
     
.NOTES
    Name    : Function_Get-PKWindowsReport.ps1
    Created : 2016-05-27
    Author  : Paula Kingsley
    Version : 04.01.0000
    History :

        *** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00.0000 - 2018-01-00 - Created script
        
                                          
.EXAMPLE



#>
[CmdletBinding(
)]
Param (
    [Parameter(
        Mandatory = $True,
        Position = 0,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "Computer name or object"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Name","Hostname","DNSHostname")]
    [object[]]$ComputerName,

    [Parameter(
        Mandatory   = $False,
        HelpMessage = "AD domain FQDN or DN (default is current user)"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain(),

    [Parameter(
        Mandatory = $False,
        HelpMessage = "Additional nameservers to query (by default uses locally specified only)"
    )]
    [ValidateNotNullOrEmpty()]
    [string[]]$AdditionalNameserver,

    [Parameter(
        Mandatory   = $False,
        HelpMessage = "Valid AD credentials/admin credentials on target computer"
    )]
    [ValidateNotNullOrEmpty()]
    [PSCredential]$Credential,

    [Parameter(
        Mandatory   = $False,
        HelpMessage = "Try DCOM for downlevel client CIM connections (default is WSMAN)"
    )]
    [ValidateNotNullOrEmpty()]
    [switch] $UseDCOM,

    [Parameter(
        Mandatory   = $False,
        HelpMessage = "Test whether Windows is running GUI or Server Core installation"
    )]
    [ValidateNotNullOrEmpty()]
    [switch] $TestWindowsMode ,

    [Parameter(
        Mandatory   = $False,
        HelpMessage = "Return multivalue properties as strings instead of collections"
    )]
    [ValidateNotNullOrEmpty()]
    [switch] $CollectionsToStrings ,

    [Parameter(
        Mandatory   = $False,
        HelpMessage = "Hide non-verbose console output"
    )]
    [ValidateNotNullOrEmpty()]
    [switch] $SuppressConsoleOutput = $False

)

Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Preference
    $ErrorActionPreference = "Stop"
    $ProgressPreference = "Continue"

    # How did we get here
    $PipelineInput = (-not $PSBoundParameters.ContainsKey("VM")) -and (-not $ComputerName)
    $Source = $PSCmdlet.ParameterSetName

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object {$CurrentParams.keys -notContains $_} | 
        Where-Object {Test-Path variable:$_}| Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    #region Prerequisites 
    $Activity = "Prerequisites"

    # Make sure we can get to the domain
    $Msg = "Test AD connectivity"
    Write-Progress -Activity $Activity -CurrentOperation $Msg

    $LDAPRoot = $DomainDN = $DN = $DomainDNSRoot = $Null
    If (-not $ADDomain) {
        $DomainObj = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        $DomainDNSRoot = $DomainObj.Name
        $DomainDN = $DomainObj.GetDirectoryEntry() | Select -ExpandProperty DistinguishedName
        $LDAPRoot = ([adsi]"LDAP://$DomainDN").Path
    }
    Else {
        If ($ADDomain -match ".") {
            $DomainDNSRoot = $ADDomain
            $DN = "DC=$($ADDomain.Replace(".",",DC="))"
            $LDAP = ([adsi]"LDAP://$DN")
            $DomainDN = $LDAP.DistinguishedName
            $LDAPRoot = $LDAP.Path
        }
        Elseif ($ADDomain -match ("DC=")) {
            $DN = $ADDomain
            $DomainDNSRoot = ($ADDomain.Replace(",DC=",".")).Replace("DC=","")
            $LDAP = ([adsi]"LDAP://$ADDomain")
            $DomainDN = $LDAP.DistinguishedName
            $LDAPRoot = $LDAP.Path
        }
    }
    If (-not $DomainDN) {
        $Msg = "No valid domain DistinguishedName provided or found; please specify a valid entry"
        $Host.UI.WriteErrorLine("ERROR: $Msg")
        Break
    }
    Else {
        $Msg = "Test Active Directory authentication"
        Write-Verbose $Msg

        If ($CurrentParams.Credential.Username) {    
            Try {
                $User = $Credential.Username
                $PW = $Credential.GetNetworkCredential().password
                $DomainInfo = New-Object System.DirectoryServices.DirectoryEntry($LDAPRoot,$User,$PW) -ErrorAction Stop
                $searcher = New-Object System.DirectoryServices.DirectorySearcher($DomainInfo)
                $Msg = "Successfully authenticated to '$LDAPRoot' as '$User'"
                Write-Verbose $Msg
            }
            Catch {
                $Msg = "Authentication failed"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
                $Host.UI.WriteErrorLine("ERROR: $Msg")
                Break
            }
        }
        Else {
            Try {
                $Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher -ErrorAction Stop
                $Msg = "Successfully connected to '$LDAPRoot'"
                Write-Verbose $Msg
            }
            Catch {
                $Msg = "Authentication failed"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
                $Host.UI.WriteErrorLine("ERROR: $Msg")
                Break
            }
        }
        If ($Searcher) {
            $Searcher.SizeLimit = 100
            $Searcher.PropertiesToLoad.AddRange(('name','canonicalname','serviceprincipalname','useraccountcontrol','lastlogonTimeStamp','distinguishedname','description','operatingsystem','location','whencreated','dnshostname'))
            $Searcher.SearchRoot = $LDAPRoot
        }
    }

    #endregion Prerequisites 

    #region Functions

    # Function to get AD object info using ADSISearcher
    Function GetADComputer {
        [Cmdletbinding()]
        Param(
            $ComputerName,
            $Searcher,
            $OUTrimDepth,
            $Credential,
            [switch]$ExactMatchOnly
        )
        Begin {
            
            # DNS lookup (trap/hide errors)
            Function Get-DNS{
                [Cmdletbinding()]
                Param($Name)
                Try {
                    [Net.Dns]::GetHostEntry($Name)
                }
                Catch {}
            }

            # Convert useraccountcontrol
            Function Get-ComputerState {
                [Cmdletbinding()]
                Param($UAC)
                Process {
                    Switch ($UAC) {
                        4096 {
                            New-Object PSObject -Property @{
                                Type = "Computer (WORKSTATION_TRUST_ACCOUNT)"
                                Enabled = $True
                            }
                        }
                        4098 {
                            New-Object PSObject -Property @{
                                Type = "Computer (WORKSTATION_TRUST_ACCOUNT)"
                                Enabled = $False
                            }
                        }
                        4128 {
                            New-Object PSObject -Property @{
                                Type = "Computer (WORKSTATION_TRUST_ACCOUNT and PASSWD_NOTREQD)"
                                Enabled = $True
                            }
                        }
                        4130 {
                            New-Object PSObject -Property @{
                                Type = "Computer (WORKSTATION_TRUST_ACCOUNT and PASSWD_NOTREQD and ACCOUNTDISABLE)"
                                Enabled = $False
                            }
                        }
                        532480 {
                            New-Object PSObject -Property @{
                                Type = "Domain controller"
                                Enabled = $True
                            }
                        }
                        Default {
                            New-Object PSObject -Property @{
                                Type = "Unknown"
                                Enabled = "Unknown"
                            }
                        }
                    }
                }
            }

            $Results = @()
            $InitialValue = "Error"
            $OutputTemplate = New-Object -TypeName PSObject -Property ([ordered] @{
                Name                 = $InitialValue
                IsEnabled            = $InitialValue
                OU                   = $InitialValue
                IPAddress            = $InitialValue
                DNSHostName          = $InitialValue
                Description          = $InitialValue
                OperatingSystem      = $InitialValue
                ObjectType           = $InitialValue        
                Location             = $InitialValue
                WhenCreated          = $InitialValue
                LastLogon            = $InitialValue
                ServicePrincipalName = $InitialValue
                CanonicalName        = $InitialValue
                DistinguishedName    = $InitialValue
                Messages             = $InitialValue
            })
        }
        Process {

            Foreach ($Computer in $ComputerName) {
                
                $Searcher.Filter = "(&(objectCategory=Computer)(name=$Computer))"

                If ($ExactMatchOnly.IsPresent) {
                    $Found = $Searcher.FindOne()
                }
                Else {
                    $Found = $Searcher.FindAll()
                }

                If ($Found.Count -gt 0) {
                
                    Foreach ($Obj in $Found) {
                
                        $NameStr = ($Obj.properties.name -as [string])
                        $Msg = $NameStr
                        
                        # Copy & populate output object
                        $Output = $OutputTemplate.PSObject.Copy()
                        $Output.Name                 = $NameStr
                        $Output.OperatingSystem      = $Obj.Properties.operatingsystem -as [string]
                        $Output.WhenCreated          = (($Obj.Properties.whencreated -as [string]) -as [datetime]).ToString()
                        $Output.ServicePrincipalName = $Obj.Properties.serviceprincipalname
                        $Output.CanonicalName        = $Obj.properties.canonicalname -as [string]
                        $Output.DistinguishedName    = $Obj.properties.distinguishedname -as [string]
                        $Output.Messages             = $Null

                        # Type & state
                        $UACType = (Get-ComputerState -UAC $($Obj.Properties.useraccountcontrol -as [string]))
                        $Output.IsEnabled = $UACType.Enabled
                        $Output.ObjectType = $UACType.Type

                        # If we're trimming out some of the OU path string
                        $CN = ($Obj.properties.canonicalname -as [string])
                        If ($OUTrimDepth) {
                            $Output.OU = (($CN.Split("/",$OUTrimDepth)[-1]) -Split("/$NameStr") -as [string])
                        }
                        Else {
                            $Output.OU = $CN.Substring(0, $CN.lastIndexOf('/'))
                        }

                        # DNS lookup
                        $DNSEntity = $Null
                        If ($DNSEntity = Get-DNS -Name $NameStr) {[string[]]$Output.IPAddress = $DNSEntity.AddressList | Foreach-Object {$_.IPAddressToString}}
                        Else {
                            $Msg = "DNS lookup failed for $NameStr"
                            Write-Warning $Msg
                            [string[]]$Output.IPAddress = "(none)"
                        }
                    
                        # Convert last logon date
                        If (-not ($Output.LastLogon = [datetime]::FromFileTime($($Obj.Properties.lastlogontimestamp)))) {$Output.LastLogon = "(unknown)"}
                    
                        # Get or create FQDN
                        If (-not ($Output.DNSHostName = $Obj.Properties.dnshostname -as [string])) {
                            $Msg = "DNSHostName not available; constructing FQDN from DNS lookup"
                            If ($DNSEntity.HostName -match $DomainDNSRoot) {$Output.DNSHostName = $DNSEntity.HostName -as [string]}
                            Else {$Output.DNSHostName = $Null}
                        }

                        # Description and location
                        If (-not ($Output.Description = ($Obj.properties.description -as [string]))) {$Output.Description = "(none)"}
                        If (-not ($Output.Location = $Obj.properties.location -as [string])) {$Output.Location = "(none)"}
                    
                        Write-Output $Output

                    } # end foreach obj found
                } # end if obj found
            } #end foreach computer
        }
        } # end function GetAdComputer

    # Function if we want to check additional nameservers
    Function LookupDNS { 
        [Cmdletbinding()]
        Param(
            [string[]]$Target,
            $NS
        )
        [switch]$NSLookup = $False
        If ($NS) {
            Write-Verbose "Nameserver: $NS"
            If ($TestNS = (new-object Net.Sockets.TcpClient -ErrorAction Stop -Verbose)) {
                If (-not ($TestNS.ConnectAsync($NS,53))) {
                    $Msg = "No response from nameserver $NS"
                    $Host.UI.WriteErrorLine($Msg)
                    Break
                }
            }
            Else {$NSLookup = $True}
        }
        Else {
            Write-Verbose "Nameserver: Local default"
        }

        Foreach ($T in $Target) {

            $Output = New-Object PSObject -Property ([ordered]@{
                Input      = $T
                Mode       = "Error"
                Name       = "Error"
                IPAddress  = "Error"
                Messages   = "Error"
            })

            # If IP address
            If ($T -as [ipaddress]) {
                
                $Output.Mode = "Reverse"

                If ($NSLookup.IsPresent) {

                    If ($Lookup = Invoke-Expression -Command "nslookup $T $NS 2>&1" -OutVariable Results -EA SilentlyContinue) {
                        $Output.IPAddress = ($Results | Select-String address | select-Object -last 1).toString().split(":")[1].trim()
                        $Output.Name = ($Results | Select-String name | select-Object -last 1).toString().split(":")[1].trim()
                        $Output.Messages = $Null
                    }
                    Else {$Output.Messages = "Reverse lookup failed for $T"}
                }
                Else {
                    If ($Results = [System.Net.DNS]::GetHostByAddress($T)) {
                        $Output.IPAddress = $($Results.AddressList)                
                        $Output.Name = ($Results.hostname.tolower())
                        $Output.Messages = $Null
                    }
                    Else {$Output.Messages = "Reverse lookup failed for $T"}
                }
            }
            
            # If hostname 
            Else {

                $Output.Mode = "Forward"

                If ($NSLookup.IsPresent) {
                    
                    If ($Lookup = Invoke-Expression -Command "nslookup $T $NS 2>&1" -OutVariable Results -EA SilentlyContinue) {
                        $Output.IPAddress = ($Results | Select-String address | select-Object -last 1).toString().split(":")[1].trim()
                        $Output.Name = ($Results | Select-String name | select-Object -last 1).toString().split(":")[1].trim()
                        $Output.Messages = $Null
                    }
                    Else {$Output.Messages = "Forward lookup failed for $T"}
                }
                Else {
                    If ($Results = [System.Net.DNS]::GetHostByName($T)) {
                        $Output.IPAddress = $($Results.AddressList)                
                        $Output.Name = ($Results.hostname.tolower())
                        $Output.Messages = $Null
                    }
                    Else {$Output.Messages = "Forward lookup failed for $T"}
                }
            }
            
            Write-Output $Output
        }
    }
    
    # DNS lookup (trap/hide errors)
    Function Get-DNS{
        [Cmdletbinding()]
        Param($Name)
        Try {
            [Net.Dns]::GetHostEntry($Name)
        }
        Catch {}
    }

    # Convert useraccountcontrol
    Function Get-ComputerState {
        [Cmdletbinding()]
        Param($UAC)
        Process {
            Switch ($UAC) {
                4096 {
                    New-Object PSObject -Property @{
                        Type = "Computer (WORKSTATION_TRUST_ACCOUNT)"
                        Enabled = $True
                    }
                }
                4098 {
                    New-Object PSObject -Property @{
                        Type = "Computer (WORKSTATION_TRUST_ACCOUNT)"
                        Enabled = $False
                    }
                }
                4128 {
                    New-Object PSObject -Property @{
                        Type = "Computer (WORKSTATION_TRUST_ACCOUNT and PASSWD_NOTREQD)"
                        Enabled = $True
                    }
                }
                4130 {
                    New-Object PSObject -Property @{
                        Type = "Computer (WORKSTATION_TRUST_ACCOUNT and PASSWD_NOTREQD and ACCOUNTDISABLE)"
                        Enabled = $False
                    }
                }
                532480 {
                    New-Object PSObject -Property @{
                        Type = "Domain controller"
                        Enabled = $True
                    }
                }
                Default {
                    New-Object PSObject -Property @{
                        Type = "Unknown"
                        Enabled = "Unknown"
                    }
                }
            }
        }
    }

    # Inner function to get site
    If ($GetADSite.IsPresent) {
        function Get-ComputerSite($ComputerName){
           Try {
                If ($SiteName = Invoke-Expression -Command "nltest /server:$ComputerName /dsgetsite 2>&1") {$Sitename[0]}    
            }
            Catch {"(unknown)"}
        }
    }
    
    # Inner function to ping
    If ($TestConnection.IsPresent) {
        #https://gist.github.com/mbrownnycnyc/9913361
        function Test-Ping{
            [CmdletBinding()]
            param([String]$ComputerName = "127.0.0.1",[int]$delay = 100)
            # see http://msdn.microsoft.com/en-us/library/system.net.networkinformation.ipstatus%28v=vs.110%29.aspx
            try {
                $ping = new-object System.Net.NetworkInformation.Ping
                If ($ping.send($ComputerName,$delay).status -ne "Success") {$false}
                Else {$true}
            } catch {$false}
        }    
    }

    #endregion Functions

    #region Splats

    # General-purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose = $False
    }

    # Splat for GetADComputer
    $Param_AD = @{}
    $Param_AD = @{
        ComputerName   = ""
        Searcher       = $Searcher
        OUTrimDepth    = 6
        ExactMatchOnly = $True
        ErrorAction    = "SilentlyContinue"
        Verbose        = $False
    }


    # Splat for New-CIMSession
    $Param_CIMSession = @{}
    $Param_CIMSession = @{
        ComputerName        = ""
        OperationTimeoutSec = 5
        Authentication      = "Kerberos"
        Credential          = $Credential
        ErrorAction         = "SilentlyContinue"
        Verbose             = $False
    }
    If ($UseDCOM.IsPresent) {
        $Opt = New-CimSessionOption -Protocol Dcom
        $Param_CIMSession.Add("SessionOption",$Opt)
    }

    # Splat for Get-CIMInstance
    $Param_CIMInstance = @{}
    $Param_CIMInstance = @{
        CIMSession          = ""
        OperationTimeoutSec = 5
        Query               = ""
        ErrorAction         = "Stop"
        Verbose             = $False
    }

    # Splat to get the virtual portgroup/vlan
    $Param_PG = @{}
    $Param_PG = @{
        Name        = ""
        VMHost      = ""
        ErrorAction = "Stop"
        Server      = $VIServer
        Verbose     = $False
        Debug       = $False
    }

    # Splat for write-progress
    $Activity = "Get Windows computer report"
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        PercentComplete  = $Null
        CurrentOperation = $Null
        Status           = "Working"
    }
    
    #endregion Splats

    #region Output objects

    # Placeholder output object
    $InitialValue = "Error"
    $OutputTemplate = @()
    $OutputTemplate = New-Object PSObject -Property ([ordered] @{
        ComputerName          = $InitialValue
        CreateDate            = $InitialValue
        Memory                = $InitialValue
        CPU                   = $InitialValue
        DNSHostName           = $InitialValue
        FQDNs                 = $InitialValue
        OU                    = $InitialValue
        Description           = $InitialValue
        Location              = $InitialValue
        CanonicalName         = $InitialValue
        DistinguishedName     = $InitialValue
        ServicePrincipalNames = $InitialValue
        ObjectType            = $InitialValue
        OSName                = $InitialValue
        OSVersion             = $InitialValue
        HardDrive             = $InitialValue
        NetworkAlias          = $InitialValue
        IPv4Address           = $InitialValue
        SubnetMask            = $InitialValue
        Gateway               = $InitialValue
        DNSServers            = $InitialValue
        PTR                   = $InitialValue
        VLANID                = $InitialValue
        WSUSServer            = $InitialValue
        WSUSGroup             = $InitialValue
        Messages              = $Null
    })
    
    # Placeholder output object for networking
    $NicOutputTemplate = @()
    $NicOutputTemplate = New-Object PSObject -Property ([ordered] @{
        AdapterAlias = $InitialValue
        Description  = $InitialValue
        IPv4Address  = $InitialValue
        SubnetMask   = $InitialValue
        Gateway      = $InitialValue
        DNS          = $InitialValue
        DHCPEnabled  = $InitialValue
        MACAddress   = $InitialValue
        VMPortgroup  = $InitialValue
        VLANID       = $InitialValue
    })
   
    #endregion Output objects

    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "Action: $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg}
                
}

Process {

    $Total = $ComputerName.Count
    $Current = 0

    Foreach ($Computer in $ComputerName) {

        $Current ++
        
        $Msg = $Computer
        Write-Verbose $Msg

        $Param_WP.Status = $Msg
        $Msg =  "Look for computer object in Active Directory"
        $Param_WP.CurrentOperation = $Msg
        $Param_WP.PercentComplete = ($Current / $Total * 100)
        Write-Progress @Param_WP

        $ErrorMessages = @()
        [switch]$Continue = $False

        $Output = $OutputTemplate.PSObject.Copy()
        $Output.ComputerName = $Computer


        # Verify it's in AD
        Try {
            $Param_AD.ComputerName = $Computer
            If ($ComputerObj = GetADComputer @Param_AD) {
                $Msg = "Found '$($ComputerObj.DNSHostName)' in '$DomainDNSRoot'"
                Write-Verbose $Msg

                # Reset flag
                $Continue = $True

                $Output.ComputerName = $ComputerObj.Name
                $Output.CreateDate = $ComputerObj.WhenCreated
                $Output.Description = $ComputerObj.Description
                $Output.OU = $ComputerObj.OU
                $Output.CanonicalName = $ComputerObj.CanonicalName
                $Output.DistinguishedName = $ComputerObj.DistinguishedName
                $Output.ServicePrincipalNames = $ComputerObj.ServicePrincipalName
                $Output.Location = $ComputerObj.Location

            }
            Else {
                $Msg = "No matching computer object found for '$Computer'"
                $Host.UI.WriteErrorLine("ERROR: $Msg")
                Break
            }
        }
        Catch {
            $Msg = "No matching computer object found for '$Computer'"
            If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
            $Host.UI.WriteErrorLine("ERROR: $Msg")
            Break
        }


        If ($Continue -eq $True) {
            
            # Set flag
            [switch] $Continue = $False

            
                # CIM connection
                If ($Continue.IsPresent) {

                    # Make sure we can connect to Windows; not resetting the flag because we will continue even if this fails 
                    $Select = $SelectNoWMI
                    [switch]$CIM = $False
                    If (-not $SkipWMICheck.IsPresent) {

                        $Msg = "Create remote CIM session"
                        $Param_WP.CurrentOperation = $Msg
                        Write-Progress @Param_WP

                        $Param_CIMSession.ComputerName = $ComputerNameObj.Name

                        Try {
                            If ($CIMSession = New-CIMSession @Param_CIMSession) {
                                $CIM = $True

                                # Output will contain all properties
                                $Select = $SelectAll

                            }
                            Else {
                                $Msg = "CIM session failed on $($ComputerNameObj.Name)"
                                $Host.UI.WriteErrorLine("ERROR: $Msg; skipping checks for internal Windows data")
                                If ($ErrorDetails = $_.Exception.Message) {$Msg = "$Msg`n$ErrorDetails"}
                                $ErrorMessages += "$Msg"
                                $CIM = $False
                                $SkipWMICheck = $True
                                
                                # Output will contain all properties
                                $Select = $SelectAll
                            }
                        }
                        Catch {
                            $SkipWMICheck = $True

                            $Msg = "CIM session failed on $($ComputerNameObj.Name)"
                            If ($ErrorDetails = "$($_.exception.message -replace '\s+', ' ')") {$Msg = "$Msg`n$Msg"}
                            $Host.UI.WriteErrorLine("ERROR: $Msg`nWindows internal checks will be skipped")
                            $ErrorMessages += "$Msg"
                            
                            # Output won't contain WMI properties
                            $Select = $SelectNoWMI
                        }
                        
                    } #end if not skipping WMI check
                
                    # Proceeding anyway
                    $Continue = $True
                }

                # Get VM disk details
                If ($Continue.IsPresent) {
                    
                    # Reset flag
                    [switch] $Continue = $False

                    Try {
                        $Msg = "Get virtual hard disk details"
                        $Param_WP.CurrentOperation = $Msg
                        Write-Progress @Param_WP

                        $DiskArr = @()
                        Foreach ($Disk in ($ComputerNameObj | Get-HardDisk @StdParams -Debug:$False)) {
                            $SCSIController = Get-ScsiController -HardDisk $Disk @StdParams -Debug:$False
                            $SCSIID = $ComputerNameView.Config.Hardware.Device | Where-Object {$_.GetType().Name -eq "VirtualDisk"} | Where-Object {$_.DeviceInfo.Label -eq $Disk.Name}
                            $SCSIInfo = "$($SCSIController.Name):$($SCSIID.UnitNumber)"
                            [array]$DiskArr += "$($Disk.Name) [$SCSIInfo] $('{0:N2}' -f $($Disk.CapacityGB)) GB" 
                        }
                        If ($CollectionsToStrings.IsPresent) {
                            $Separator = "`n"
                            If ($DiskArr.Count -eq 1) {$Separator = ""}
                            $Output.VMDisk = ($DiskArr -join($Separator))
                        }
                        Else {
                            $Output.VMDisk = $DiskArr
                        }
                    
                        $Msg = "Get virtual SCSI controller details"
                        $Param_WP.CurrentOperation = $Msg
                        Write-Progress @Param_WP
                    
                        $ControllerArr = @()
                        Foreach ($Disk in ($ComputerNameObj | Get-HardDisk @StdParams -Debug:$False)) {
                            $Controller = Get-ScsiController -HardDisk $Disk @StdParams -Debug:$False
                            $ControllerStr = "$($Controller.Name) ($($Controller.Type))"
                            $ControllerArr += $ControllerStr
                        }
                        $ControllerArr = ($ControllerArr | Select-Object -Unique)
                        If ($CollectionsToStrings.IsPresent) {
                            $Separator = "`n"
                            If ($ControllerArr.Count -eq 1) {$Separator = ""}
                            $Output.VMSCSIController = ($ControllerArr -join($Separator))
                        }
                        Else {
                            $Output.VMSCSIController = $ControllerArr
                        }

                        $Continue = $True
                    }
                    Catch {
                        $Msg = "Virtual hard drive/SCSI controller lookup failed for VM $($ComputerNameObj.Name)"
                         If ($ErrorDetails = "$($_.exception.message -replace '\s+', ' ')") {$Msg = "$Msg`n$Msg"}
                        $Host.UI.WriteErrorLine("ERROR: $Msg")
                        $ErrorMessages += "$Msg"
                        #$Continue = $False   
                        $Continue = $True
                    }
                
                    # Proceeding anyway
                    $Continue = $True
                }

                # Datastores
                If ($Continue.IsPresent) {
                    
                    # Reset flag
                    [switch]$Continue = $False 
                    
                    $Msg = "Get datastore"
                    Write-Verbose $Msg
                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP

                    Try {
                        $Datastore = @()
                        $DataStore = ($ComputerNameObj | Get-DataStore @StdParams -Debug:$False).Name
                         
                        $Continue = $True

                        If ($CollectionsToStrings.IsPresent) {
                            $Separator = "`n"
                            If ($Datastore.Count -eq 1) {$Separator = ""}
                            $Output.VMDatastore = $Datastore -join("`n")
                        }
                        Else {
                            $Output.VMDatastore = $Datastore
                        }                   
                    }
                    Catch {
                        $Msg = "Datastore lookup failed for VM $($ComputerNameObj.Name) "
                        If ($ErrorDetails = "$($_.exception.message -replace '\s+', ' ')") {$Msg = "$Msg`n$Msg"}
                        $Host.UI.WriteErrorLine("ERROR: $Msg")
                        $ErrorMessages += "$Msg"
                    }               
                
                    # Proceeding anyway
                    $Continue = $True
                }

                # Network
                If ($Continue.IsPresent) {
                    
                    # Reset flag
                    [switch]$Continue = $False 

                    $Msg = "Get virtual network adapter"
                    Write-Verbose $Msg
                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP

                    Try {
                        $ComputerNameNICArr = @()
                        $PortgroupArr = @()
                        $VLANArr = @()
                        $IPArr = @()
                        $PTRArr = @()
                    
                        #Foreach ($ComputerNameNIC in ($ComputerNameObj | Get-NetworkAdapter -ErrorAction SilentlyContinue -Verbose:$False -Debug:$False)) {

                        $ComputerNameObj | Get-NetworkAdapter -ErrorAction SilentlyContinue -Verbose:$False -Debug:$False | Foreach-Object {
                            
                            $ComputerNameNIC = $_
                            $ComputerNameNICArr += $ComputerNameNIC

                            $ComputerNameNICStr = "$($ComputerNameNIC.Name) ($($ComputerNameNic.Type), $($ComputerNameNIC.NetworkName), $($ComputerNameNIC.MACAddress))"
                            
                            $PortGroup = $($ComputerNameNIC.NetworkName)
                            $PortGroupArr += $PortGroup

                            $VLAN = (Get-VirtualPortGroup -Name $PortGroup -VM $ComputerNameObj -ErrorAction SilentlyContinue -Verbose:$False -Debug:$False).VLANID
                            $VLANArr += $VLAN
                        }

                        # Although we can/will get this later from the guest, we want something if we're skipping WMI
                        $IPArr = $ComputerNameObj.Guest.IPAddress | Where-Object {$_ -match "\."}
                        $PTRArr = $IPArr | Foreach-Object {
                            [System.Net.Dns]::GetHostEntry($_).HostName
                        }

                        If ($ComputerNameNICArr.Count -gt 0) {
                            If ($CollectionsToStrings.IsPresent) {
                                $Separator = "`n"
                                If ($ComputerNameNicArr.Count -eq 1) {$Separator = ""}
                                $Output.VMNetworkAdapter = $ComputerNameNICArr -join($Separator)
                                $Separator = "`n"
                                If ($PortGroup.Count -eq 1) {$Separator = ""}
                                $Output.VMPortgroup = $PortgroupArr -join($Separator)
                                $Separator = "`n"
                                If ($VLANArr.Count -eq 1) {$Separator = ""}
                                $Output.VLANID = $VLANArr -join($Separator)
                                $Separator = "`n"
                                If ($IPArr.Count -eq 1) {$Separator = ""}
                                $Output.IPv4Address = $IPArr -join($Separator)
                                $Separator = "`n"
                                If ($PTRArr.Count -eq 1) {$Separator = ""}
                                $Output.PTR = $PTRArr -join($Separator)
                            }
                            Else {
                                $Output.VMNetworkAdapter = $($ComputerNameNICArr)
                                $Output.VMPortgroup = $($PortgroupArr)
                                $Output.VLANID = $($VLANArr)
                                $Output.IPv4Address = $($IPArr)
                                $Output.PTR = $($PTRArr)
                            }
                        }
                        Else {$Output.VMNetworkAdapter = "(none)"}

                        # Reset flag
                        $Continue = $True
                    }
                    Catch {
                        $Msg = "Network adapter query failed for VM $($ComputerNameObj.Name) "
                        If ($ErrorDetails = "$($_.exception.message -replace '\s+', ' ')") {$Msg = "$Msg`n$Msg"}
                        $Host.UI.WriteErrorLine("ERROR: $Msg")
                        $ErrorMessages += "$Msg"   
                        #$Continue = $False  
                    }                
                
                    # Proceeding anyway
                    $Continue = $True
                }

                # CPU / RAM (already have RAM, but console output should be consistent
                If ($Continue.IsPresent) {
                    
                    # Reset flag
                    [switch]$Continue = $False
                
                    $Msg = "Get CPU/RAM"
                    Write-Verbose $Msg
                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP

                    Try {
                        $ComputerNameView = $Null

                        If ($ComputerNameView = (Get-View -ViewType VirtualMachine -Filter @{"Name" = "^$($ComputerNameObj.Name)$"} @StdParams -Debug:$False)) {

                            $TotalCPUs      = $ComputerNameView.config.hardware.NumCPU
                            $TotalSockets   = ($($ComputerNameView.config.hardware.NumCPU) / $($ComputerNameView.config.hardware.NumCoresPerSocket) )
                            $CoresPerSocket = $($ComputerNameView.config.hardware.NumCoresPerSocket)
                            $CPUs = "Count: $TotalCPUs","TotalSockets: $TotalSockets","CoresPerSocket:$CoresPerSocket"

                            If ($CollectionsToStrings.IsPresent) {
                                $Separator = "`n"
                                If ($CPUs.Count -eq 1) {$Separator = ""}
                                $Output.CPU = $CPUs -join($Separator)
                            }
                            Else {
                                $Output.CPU = $($CPUs)
                            }
                        }
                        Else {
                            $Output.CPU = $ComputerNameObj.NumCPU
                        }
                        
                        $Continue = $True
                    }
                    Catch {
                        $Msg = "CPU/RAM lookup failed for VM $($ComputerNameObj.Name) "
                        If ($ErrorDetails = "$($_.exception.message -replace '\s+', ' ')") {$Msg = "$Msg`n$Msg"}
                        $Host.UI.WriteErrorLine("ERROR: $Msg")
                        $ErrorMessages += "$Msg"  
                        #$Continue = $False  
                    }
                
                    # Proceeding anyway
                    $Continue = $True
                }
                
                # AD (but if all the following fails we won't break out of the script)
                If ($Continue.IsPresent) {
                    
                    $Msg =  "Get Active Directory details"
                    Write-Verbose $Msg
                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP

                    Try {
                        # We're trimming the OU string to the 6th slash for neatness
                        If ($ADObj = GetADComputer -ComputerName $ComputerNameObj.Name -LDAPRoot $LDAPRoot -OUTrimDepth 6 -ExactMatchOnly) {

                            # This can happen, so just in case...
                            If ($ADObj.Count -gt 1) {
                                $Msg = "Multiple matching AD objects found"
                                $Host.UI.WriteWarningLine($Msg)
                                $ErrorMessages += $Msg
                            }
                            
                            $Output.CreateDate        = $ADObj.WhenCreated
                            $Output.HostName          = $ADObj.Name
                            $Output.OU                = $ADObj.OU
                            $Output.ADDescription     = $ADObj.Description
                            $Output.ADLocation        = $ADObj.Location
                            $Output.DistinguishedName = $ADObj.DistinguishedName

                        } 
                        Else {
                            $Msg = "No matching Active Directory object found"
                            $Host.UI.WriteErrorLine($Msg)
                            $ErrorMessages += $Msg
                        }
                    }
                    Catch {
                        $Msg = "No matching Active Directory object found"
                        If ($ErrorDetails = $_.Exception.Message) {$Msg = "$Msg`n$ErrorDetails"}
                        $Host.UI.WriteErrorLine($Msg)
                        $ErrorMessages += $Msg
                    }
                    
                }

                # DNS 
                If ($Continue.IsPresent) {

                    #  Network
                    $Msg = "Test/get DNS" 
                    Write-Verbose $Msg
                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP

                    $FQDNArr = @()
                    $FQDNArr += $ADObj.DNSHostname
                    If (-not $ADObj.DNSHostName) {
                        $Msg = "No DNSHostname value found for $($ADObj.Name); looking up FQDN in DNS"
                        Write-Warning $Msg
                        $ErrorMessages += $Msg
                        $FQDNArr += (LookupDNS -Target $ADObj.Name).Name
                    }
                    ElseIf ($ADObj.DNShostName  -notmatch $DomainDNSRoot) {
                        $Msg = "DNSHostname value $($ADObj.DNShostName) does not match domain DNSRoot"
                        Write-Warning $Msg
                        $ErrorMessages += $Msg
                        $FQDNArr += "$($ADObj.Name.tolower()).$DomainDNSRoot"
                    }
                   
                    If ($CurrentParams.AdditionalNameserver) {
                        Foreach ($NS in $AdditionalNameServer) {
                            $FQDNArr += (LookupDNS -Target $ADObj.Name -NS $NS).Name
                        }
                    }

                    # Update output
                    If ($FQDNArr) {
                        If ($CollectionsToStrings.IsPresent) {
                            $Separator = "`n"
                            If ($FQDNArr.Count -eq 1) {$Separator = ""}
                            $Output.FQDNs = $FQDNArr -join($Separator)
                        }
                        Else {
                            $Output.FQDNs = $FQDNArr
                        }
                    }
                    Else {
                        $Output.FQDNs = "(none)"
                    }
                }

                # CIM/OS-level data
                If ($Continue.IsPresent) {

                    # Get OS-level details using WMI

                    If ($CIM.IsPresent) {                        
                        
                        Try {
                            
                            #  OS
                            $Msg = "Get operating system details via WMI"
                            Write-Verbose $Msg
                            $Param_WP.CurrentOperation = $Msg
                            Write-Progress @Param_WP

                            $Param_CIMInstance.CIMSession = $CIMSession
                            $Param_CIMInstance.Query = "Select * FROM win32_operatingsystem"

                            $OS = (Get-CIMInstance @Param_CIMInstance)
                            $Output.OSName = $OS.Caption
                            $Output.OSVersion = $OS.Version
        
                            If ($TestWindowsMode.IsPresent) {
                                    
                                $Msg = "Get operating system installation mode (Core or GUI)"
                                Write-Verbose $Msg
                                Write-Verbose $Msg
                                $Param_WP.CurrentOperation = $Msg
                                Write-Progress @Param_WP

                                Try {
                                    $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerNameObj.Name)
                                    $RegKey= $Reg.OpenSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion")
                                    Switch ($RegKey.GetValue("InstallationType")) {
                                        "Server" {$Output.OSName = "$($OS.Caption) (GUI)"}
                                        "Server Core" {$Output.OSName = "$($OS.Caption) (Core)"}
                                        Default {$Output.OSName = "$($OS.Caption)"}
                                    }
                                                                }
                                Catch {
                                    $Msg = "Can't get OS installation mode via remote registry"
                                    Write-Warning $Msg
                                    $ErrorMessages += $Msg
                                    $Output.OSName = "$($OS.Caption)"
                                }
                            }
                            
                            # Network
                            $Msg = "Get Windows network configuration"
                            Write-Verbose $Msg
                            $Param_WP.CurrentOperation = $Msg
                            Write-Progress @Param_WP

                            Try {
                                $Param_CIMInstance.Query = "Select * FROM Win32_NetworkadapterConfiguration WHERE IPEnabled = TRUE"

                                [array]$WinNICS = Get-CIMInstance @Param_CIMInstance 
                            
                                $WinNICArr = @()
                                ForEach ($NIC in $WinNICs) {
                                
                                    # Create new hashtables
                                    $NICOutput = $NicOutputTemplate.PSObject.Copy()
                                    
                                    # update splats
                                    $Param_CIMInstance.Query = "Select * FROM Win32_Networkadapter WHERE DeviceID=$($NIC.Index)"
                                    $Param_PG.VMHost   = $ComputerNameObj.VMHost.Name
                                
                                    #DNS 
                                    If ($CollectionsToStrings.IsPresent) {
                                        $Separator = "`n"
                                        If ($NIC.DNSServerSearchOrder.Count -eq 1) {$Separator = ""}
                                        $NICOutput.DNS = ($NIC.DNSServerSearchOrder -join($Separator))
                                    }
                                    Else {
                                        $NICOutput.DNS = $NIC.DNSServerSearchOrder
                                    }
                                    
                                    # Get the port group for each NIC
                                    Try {
                                        $PortGroup = ($ComputerNameObj | Get-NetworkAdapter @StdParams -Debug:$False | Where-Object {$_.MacAddress -eq $NIC.MACAddress}).NetworkName
                                        $Param_PG.Name = $PortGroup
                                        $VLAN = (Get-VirtualPortGroup @Param_PG) | Select-Object Name,VLanID

                                        $NICOutput.AdapterAlias = (Get-CIMInstance @Param_CIMInstance ).NetConnectionID
                                        $NICOutput.Description  = $NIC.Description
                                        $NICOutput.IPv4Address  = $NIC.IPAddress | Where-Object {$_ -notmatch ":"}
                                        $NICOutput.SubnetMask   = $NIC.IPSubnet | Where-Object {$_ -match "\."}
                                        $NICOutput.Gateway      = $NIC.DefaultIPGateway | Where-Object {$_ -notmatch ":"}
                                        $NICOutput.MACAddress   = $NIC.MACAddress
                                        $NICOutput.DHCPEnabled  = $NIC.DHCPEnabled
                                        $NICOutput.VMPortGroup  = $VLAN.Name
                                        $NICOutput.VLANID       = $VLAN.VLANID    
                                    }
                                    Catch {
                                        $Msg = "Can't get network adapter details"
                                        $ErrorDetails = "$($_.exception.message -replace '\s+', ' ')"
                                        $Host.UI.WriteErrorLine("ERROR: $Msg; $ErrorDetails")
                                        $ErrorMessages += "$Msg; $ErrorDetails"
                                    }   

                                    $WinNICArr += $NICOutput

                                } #end for each NIC

                                # Get PTRs
                                $PTRArr = $WinNICArr.IPv4Address | Foreach-Object {
                                    [System.Net.Dns]::GetHostEntry($_).HostName
                                }
                                
                                # Update output 
                                If ($CollectionsToStrings.IsPresent) {
                                    $Separator = "`n"
                                    If ($WinNICArr.IPv4Address.Count -eq 1) {$Separator = ""}
                                    $Output.IPv4Address = $WinNICArr.IPv4Address -join($Separator)
                                    $Separator = "`n"
                                    If ($PTRArr.Count -eq 1) {$Separator = ""}
                                    $Output.PTR = $PTRArr -join($Separator)
                                    $Separator = "`n"
                                    If ($WinNICArr.AdapterAlias.Count -eq 1) {$Separator = ""}
                                    $Output.NetworkAlias = $WinNICArr.AdapterAlias -join($Separator)
                                    $Separator = "`n"
                                    If ($WinNICArr.SubnetMask.Count -eq 1) {$Separator = ""}
                                    $Output.SubnetMask = $WinNICArr.SubnetMask -join($Separator)
                                    $Separator = "`n"
                                    If ($WinNICArr.Gateway.Count -eq 1) {$Separator = ""}
                                    $Output.Gateway = $WinNICArr.Gateway -join($Separator)
                                    $Separator = "`n"
                                    If ($WinNICArr.DNS.Count -eq 1) {$Separator = ""}
                                    $Output.DNSServers = $WinNICArr.DNS -join($Separator)
                                    $Separator = "`n"
                                    If ($WinNICArr.VMPostGroup.Count -eq 1) {$Separator = ""}
                                    $Output.VMPortGroup = $WinNICArr.VMPortGroup -join($Separator)
                                    $Separator = "`n"
                                    If ($WinNICArr.VLANID.Count -eq 1) {$Separator = ""}
                                    $Output.VLANID = $WinNICArr.VLANID -join($Separator)
                                }
                                Else {
                                    $Output.NetworkAlias  = $WinNICArr.AdapterAlias
                                    $Output.IPv4Address   = $WinNICArr.IPv4Address
                                    $OUtput.PTR           = $PTRArr
                                    $Output.SubnetMask    = $WinNICArr.SubnetMask
                                    $Output.Gateway       = $WinNICArr.Gateway
                                    $Output.DNSServers    = $WinNICArr.DNS
                                    $Output.VMPortGroup   = $WinNICArr.VMPortGroup
                                    $Output.VLANID        = $WinNICArr.VLANID
                                }   
                            }
                            Catch {
                                $Msg = "Can't get network adapter(s)"
                                $ErrorDetails = "$($_.exception.message -replace '\s+', ' ')"
                                $Host.UI.WriteErrorLine("ERROR: $Msg; $ErrorDetails")
                                $ErrorMessages += "$Msg; $ErrorDetails"
                            }

                            #  Windows drives
                            $Msg = "Get Windows hard drives"
                            Write-Verbose $Msg
                            $Param_WP.CurrentOperation = $Msg
                            Write-Progress @Param_WP

                            $DriveArr = @()
                            $Separator = "`n"
                            
                            $Param_CIMInstance.Query = "Select * FROM Win32_Volume WHERE DriveType=3 AND (NOT Driveletter = null)"

                            Foreach ($Drive in (Get-CIMInstance @Param_CIMInstance | Sort-Object Name)) {
                                If ($Drive.Label -eq $Null) {$Label = "(none)"}
                                Else {$Label = $Drive.Label}
                                [array]$DriveArr += "$($Drive.Name) [$Label] $('{0:N2}' -f $($Drive.Capacity/1GB)) GB" 
                            }
                            If ($CollectionsToStrings.IsPresent) {
                                $Separator = "`n"
                                If ($DriveArr.Count -eq 1) {$Separator = ""}
                                $Output.HardDrive = ($DriveArr -join($Separator))
                            }
                            Else {
                                $Output.HardDrive = $DriveArr
                            }
                            
                            # WSUS 
                            $Msg =  "Get WSUS server and target group"
                            Write-Verbose $Msg
                            $Param_WP.CurrentOperation = $Msg
                            Write-Progress @Param_WP

                            Try {
                                $WSUSGroup = "(none)"
                                $WSUSServer = "(none)"

                                If ($ServerReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine",$ADObj.Name)) {

                                    If ($WSUSEnv = $ServerReg.OpenSubKey('Software\Policies\Microsoft\Windows\WindowsUpdate')) {
                                
                                        $WSUSServer= $WSUSEnv.GetValue('WUServer')
                                        $WSUSGroup = $WSUSEnv.GetValue('TargetGroup')
                                        
                                    } # end if WU
                                    Else {
                                        $Msg = "No WSUS server or target group found in registry for $V"
                                        $ErrorMessages += $Msg
                                        $Host.UI.WriteWarningLine("$Msg")
                                    }

                                } # end if registry
                                Else {
                                    $Msg = "Remote registry lookup failed for WSUS configuration"
                                    $ErrorMessages += $Msg
                                    $Host.UI.WriteWarningLine("$Msg")
                                }
                            }
                            Catch {
                                $Msg = "Remote registry lookup failed for WSUS server name or target group for $V"
                                $ErrorDetails = $_.exception.message
                                $Host.UI.WriteErrorLine("ERROR: $Msg; $ErrorDetails")
                                
                                $ErrorMessages += "$Msg; $ErrorDetails"
                                $WSUSGroup = $WSUSServer = "Error"
                            }

                            $Output.WSUSServer = $WSUSServer
                            $Output.WSUSGroup = $WSUSGroup

                            $Null = $CIMSession | Remove-CIMSession -Verbose:$False -ErrorAction SilentlyContinue -Confirm:$False
                        } #end try gwmi
                        
                        Catch {
                            $Msg = "WMI lookup failed for $V"
                            $ErrorDetails = $_.exception.message
                            $Host.UI.WriteErrorLine("ERROR: $Msg; $ErrorDetails")
                            $ErrorMessages += "$Msg; $ErrorDetails"
                        }
                        } #end try getting CIM data
                    Else {
                        $Output.OSName = $ComputerNameObj.ExtensionData.Guest.GuestFullName
                    }
                }
                
                # Display any error messages
                Switch ($ErrorMessages.Count) {
                    0 {$Output.Messages = $Null}
                    1 {$Output.Messages = $ErrorMessages | Out-String}
                    Default {
                        $Separator = "`n"
                        $Output.Messages = ($ErrorMessages -join($Separator))
                    }
                }
                
                # Return PSObject 
                Write-Output $Output | Select $Select

            } # end for each VM found

        } #end if VM found
                
    } #end for each VM

} #end process

End {
    
    Write-Progress -Activity $Activity -Completed

}

} #end Get-PKWindowsReport

