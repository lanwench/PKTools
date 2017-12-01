#requires -Version 3
Function Get-PKADComputerMiniReport {
<#
.SYNOPSIS
    Uses the ADSI type accelerator to return AD computer object AD details (no ActiveDirectory module required)

.DESCRIPTION
    Uses the ADSI type accelerator to return AD computer object AD details (no ActiveDirectory module required)
    Optional switches look up the AD site name and ping the computer 
    Accepts pipeline input
    Returns a PSobject

.NOTES        
    Name    : Function_Test-PKConnection.ps1
    Author  : Paula Kingsley
    Created : 2017-11-20
    Version : 01.00.0000
    History:

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2017-11-20 - Created script based on jrich's original
        
.LINK
    https://gallery.technet.microsoft.com/scriptcenter/Powershell-Test-Server-e0cdea9a?ranMID=24542&ranEAID=je6NUbpObpQ&ranSiteID=je6NUbpObpQ-1p7VW5KEESFnLgSaMed_Bw&tduid=(05cc9a118b47a445311a4d27bb47d63a)(256380)(2459594)(je6NUbpObpQ-1p7VW5KEESFnLgSaMed_Bw)()

.PARAMETER ComputerName
    Name of computer to check; separate multiple names with commas

.PARAMETER SizeLimit
    Maximum number of objects to return per search string (default is 100)

.PARAMETER DomainDN
    LDAP path to AD domain distinguishedname (default is local computer domain)

.PARAMETER GetADSite
    Look up the AD site for the computer if it's enabled (will slow processing)

.PARAMETER TestConnection
    Ping the computer if it's enabled (will slow processing)

.PARAMETER Credential
    Valid credential on target

.PARAMETER Tests
    Tests to perform (Default is DNS, Ping, RDP, Registry, SSH, WinRM and WMI; ShowMenu allows for selection, including CredSSP)

.PARAMETER SuppressconsoleOutput
    Suppress all non-verbose/non-error console output

.EXAMPLE
    PS C:\> Get-PKADComputerMiniReport -ComputerName qa-webserver -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                                    
        ---                   -----                                    
        ComputerName          {qa-webserver}                           
        Verbose               True                                     
        SizeLimit             100                                      
        DomainDN              LDAP://DC=domain,DC=local
        GetADSite             False                                    
        TestConnection        False                                    
        Credential            System.Management.Automation.PSCredential
        SuppressConsoleOutput False                                    
        PipelineInput         True                                     
        ScriptName            Get-PKADComputerMiniReport               
        ScriptVersion         1.0.0                                    

        Action: Get AD computer mini report

        VERBOSE: Search for qa-webserver
        VERBOSE: 1 computer object(s) found

        Name                 : qa-webserver
        IsEnabled            : True
        IPAddress            : {10.11.178.23}
        DNShostName          : qa-webserver.domain.local
        Description          : Test IIS box
        OperatingSystem      : Windows Server 2012 R2 Datacenter
        ObjectType           : Computer (WORKSTATION_TRUST_ACCOUNT)
        Location             : Reno
        WhenCreated          : 2017-03-09 21:20:33
        LastLogon            : 2017-11-21 19:07:47
        ServicePrincipalName : {WSMAN/qa-webserver, WSMAN/qa-webserver.domain.local, TERMSRV/qa-webserver.domain.local, 
                               TERMSRV/qa-webserver...}
        CanonicalName        : domain.local/Servers/qa-webserver
        DistinguishedName    : CN=qa-webserver,OU=Servers,DC=domain,DC=local
        Messages             : 


.EXAMPLE
    PS C:\> $ $Arr | Get-PKADComputerMiniReport -Verbose | FT

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                                    
        ---                   -----                                    
        Verbose               True                                     
        ComputerName                                                   
        SizeLimit             100                                      
        DomainDN              LDAP://DC=domain,DC=local
        GetADSite             False                                    
        TestConnection        False                                    
        Credential            System.Management.Automation.PSCredential
        SuppressConsoleOutput False                                    
        PipelineInput         True                                     
        ScriptName            Get-PKADComputerMiniReport               
        ScriptVersion         1.0.0                                    

        Action: Get AD computer mini report

        VERBOSE: Search for sqlserver-14
        VERBOSE: 1 computer object(s) found
        VERBOSE: Search for testbox-old
        VERBOSE: 1 computer object(s) found
        VERBOSE: Search for foo
        ERROR: No AD object match found for foo
        VERBOSE: Search for webserver1
        VERBOSE: 1 computer object(s) found
        VERBOSE: Search for webserver-2
        VERBOSE: 1 computer object(s) found
        
        Name            IsEnabled IPAddress       DNShostName                 Description                                         
        ----            --------- ---------       -----------                 -----------                                         
        sqlserver-14         True {10.11.128.208} sqlserver-14.domain.local   prod sql server for project X
        testbox-old         False {10.11.178.58}  testbox-old.domain.local    Karen's test SQL VM 
        foo                 Error Error           Error                       Error
        webserver-1          True {10.11.178.164} webserver1.domain.local     Dev - server
        webserver-2          True {10.11.178.67}  webserver-2.domain.local    Dev - server
        
.EXAMPLE
    PS C:\> Get-PKADComputerMiniReport -ComputerName ops-* -GetADSite -TestConnection 

        Action: Get AD computer mini report
        WARNING: Retrieving AD site information will slow down processing, especially with a large number of computer names
        WARNING: Pinging computers will slow down processing, especially with a large number of computer names

        Name                 : OPS-PATCHSRV-3
        IsEnabled            : True
        IsAlive              : True
        IPAddress            : {10.11.129.248}
        DNShostName          : OPS-PATCHSRV-3.domain.local
        Description          : Ops - WSUS server for workstations
        OperatingSystem      : Windows Server 2012 R2 Datacenter
        ObjectType           : Computer (WORKSTATION_TRUST_ACCOUNT)
        Site                 : WEST
        Location             : Reno
        WhenCreated          : 2015-12-10 22:17:55
        LastLogon            : 2017-11-24 16:27:20
        ServicePrincipalName : {WSMAN/OPS-PATCHSRV-3, WSMAN/OPS-PATCHSRV-3.domain.local, TERMSRV/OPS-PATCHSRV-3, 
                               TERMSRV/OPS-PATCHSRV-3.domain.local...}
        CanonicalName        : domain.local/Dev/Servers/ops-patchsrv-3
        DistinguishedName    : CN=OPS-PATCHSRV-3,OU=Servers,DC=domain,DC=local
        Messages             : 

        Name                 : ops-admin-1
        IsEnabled            : True
        IsAlive              : True
        IPAddress            : {10.11.178.104}
        DNShostName          : ops-admin-1.domain.local
        Description          : Ops - Production application server for DAM (D/R for ops-monitorDB-201)
        OperatingSystem      : Windows Server 2012 R2 Datacenter
        ObjectType           : Computer (WORKSTATION_TRUST_ACCOUNT)
        Site                 : WEST
        Location             : Reno
        WhenCreated          : 2017-04-20 20:31:27
        LastLogon            : 2017-11-22 20:02:47
        ServicePrincipalName : {WSMAN/ops-admin-1.domain.local, WSMAN/ops-admin-1, TERMSRV/ops-admin-1, 
                               TERMSRV/ops-admin-1.domain.local...}
        CanonicalName        : domain.local/Dev/Servers/ops-admin-1
        DistinguishedName    : CN=ops-admin-1,OU=Servers,DC=domain,DC=local
        Messages             : 

        Name                 : ops-monitorDB-201
        IsEnabled            : True
        IsAlive              : True
        IPAddress            : {10.187.36.229}
        DNShostName          : ops-monitorDB-201.domain.local
        Description          : Ops - Production application server for DAM
        OperatingSystem      : Windows Server 2012 R2 Datacenter
        ObjectType           : Computer (WORKSTATION_TRUST_ACCOUNT)
        Site                 : LosAngeles
        Location             : LosAngeles
        WhenCreated          : 2017-03-22 13:34:43
        LastLogon            : 2017-11-27 15:23:47
        ServicePrincipalName : {WSMAN/ops-monitorDB-201, WSMAN/ops-monitorDB-201.domain.local, 
                               TERMSRV/ops-monitorDB-201.domain.local, TERMSRV/ops-monitorDB-201...}
        CanonicalName        : domain.local/Servers/ops-monitordb-201
        DistinguishedName    : CN=ops-monitorDB-201,OU=Servers,DC=domain,DC=local
        Messages             : 

        Name                 : ops-testvm
        IsEnabled            : True
        IsAlive              : True
        IPAddress            : {10.187.37.229}
        DNShostName          : ops-testvm.domain.local
        Description          : Ops - test VM
        OperatingSystem      : Windows Server 2012 R2 Datacenter
        ObjectType           : Computer (WORKSTATION_TRUST_ACCOUNT)
        Site                 : SoCal
        Location             : Los Angeles
        WhenCreated          : 2017-03-22 19:55:56
        LastLogon            : 2017-11-26 21:10:01
        ServicePrincipalName : {TERMSRV/ops-testvm, TERMSRV/ops-testvm.domain.local, WSMAN/ops-testvm, 
                               WSMAN/ops-testvm.domain.local...}
        CanonicalName        : domain.local/Dev/Servers/Ops/ops-testvm
        DistinguishedName    : CN=ops-testvm,OU=Ops,OU=Servers,OU=Dev,DC=domain,DC=local

.EXAMPLE
    PS C:\> Get-PKADComputerMiniReport -ComputerName archiv* -DomainDN LDAP://DC=foo,DC=bar -GetADSite -TestConnection -Credential $Cred -Verbose | FT

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                                    
        ---                   -----                                    
        ComputerName          {archiv*}                            
        GetADSite             True                                     
        TestConnection        True                                     
        Verbose               True                                     
        SizeLimit             100                                      
        DomainDN              LDAP://DC=foo,DC=bar
        Credential            System.Management.Automation.PSCredential
        SuppressConsoleOutput False                                    
        PipelineInput         True                                     
        ScriptName            Get-PKADComputerMiniReport               
        ScriptVersion         1.0.0                                    

        Action: Get AD computer mini report
        WARNING: Retrieving AD site information will slow down processing, especially with a large number of computer names
        WARNING: Pinging computers will slow down processing, especially with a large number of computer names

        VERBOSE: Search for archiv*
        VERBOSE: 2 computer object(s) found
        Site name testing unavailable for disabled computer archival3
        Connection testing unavailable for disabled computer archival3
        Site name testing unavailable for disabled computer ARCHIVE-BOX
        Connection testing unavailable for disabled computer ARCHIVE-BOX

        Name         IsEnabled IsAlive IPAddress       DNShostName         Description   OperatingSystem                  
        ----         --------- ------- ---------       -----------         -----------   ---------------                  
        archival3        False   False {10.11.178.118} archival3.foo.bar   test vm2      Windows Server 2012 R2 Datacenter
        archive-box      False   False {192.168.30.4}  archive-box.foo.bar (none)        Windows Server 2012 R2 Datacenter


#>

[CmdletBinding()]
Param(
    [Parameter(
        Position = 0,
        ValueFromPipelineByPropertyName = $true,
        ValueFromPipeline = $true,
        HelpMessage = "One or more computer names to search for"
    )]
    [Alias("Computer","Name","VM","Hostname","FQDN","DNSHostName")]
    [String[]]$ComputerName,
    
    [Parameter(
        Mandatory = $False,
        Position = 1,
        HelpMessage = "Maximum number of objects to return"
    )]
    [Alias("ResultLimit","Limit")]
    [int]$SizeLimit = '100',

    [Parameter(
        Mandatory = $False,
        Position = 2,
        ValueFromPipelineByPropertyName=$true,
        HelpMessage = "DistinguishedName of domain (default is current computer domain)"
    )]
    [Alias("Domain")]
    [String]$DomainDN = $(([adsisearcher]"").Searchroot.path),
    
    [Parameter(
        Mandatory=$False,
        HelpMessage="Get AD site name (slows processing)"
    )]
    [Switch] $GetADSite,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Ping computer (slows processing)"
    )]
    [Switch] $TestConnection,

    [Parameter(
        Mandatory = $False,
        Position = 3,
        HelpMessage = "Credentials in domain"
    )]
    [Alias("RunAs")]
    [System.Management.Automation.Credential()]$Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Hide all non-verbose/non-error console output"
    )]
    [Switch] $SuppressConsoleOutput


)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    $PipelineInput = (-not $PSBoundParameters.ContainsKey("VM")) -and (-not $VM)

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

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference   = "Continue"
    
    #region Prerequisites

    Try {
        $Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher -ErrorVariable ErrProcessNewObjectSearcher @StdParams
        $Searcher.SizeLimit = $SizeLimit
        $searcher.PropertiesToLoad.AddRange(('name','canonicalname','serviceprincipalname','useraccountcontrol','lastlogonTimeStamp','distinguishedname','description','operatingsystem','location','whencreated','dnshostname'))

        $Searcher.SearchRoot = $DomainDN
        If ($DomainDN -notmatch "LDAP://") {$Searcher.SearchRoot = "LDAP://$DomainDN"}
        
        $DomainDNSRoot = ($DomainDN -replace('DC=',$Null) -replace(',','.')).split('/')[-1]
    }
    Catch {
        $Msg = "DirectorySearcher creation failed"
        If ($ErrorDetails = $($_.exception.message -replace '\s+', ' ')) {$Msg = "$Msg`n$ErrorDetails"}
        $Host.UI.WriteErrorLine("ERROR: $Msg") 
        Break
    }

    #endregion Prerequisites

    #region Splats

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    # Splat for Write-Progress
    $Activity = "Get AD computer mini report"
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = "Working"
        PercentComplete  = $Null
    }
    
    #endregion Splats

    #region Functions

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

    #region Output

    $InitialValue = "Error"
    $OutputTemplate = New-Object -TypeName PSObject -Property ([ordered] @{
        Name                 = $InitialValue
        IsEnabled            = $InitialValue
        IsAlive              = $InitialValue
        IPAddress            = $InitialValue
        DNShostName          = $InitialValue
        Description          = $InitialValue
        OperatingSystem      = $InitialValue
        ObjectType           = $InitialValue
        Site                 = $InitialValue        
        Location             = $InitialValue
        WhenCreated          = $InitialValue
        LastLogon            = $InitialValue
        ServicePrincipalName = $InitialValue
        CanonicalName        = $InitialValue
        DistinguishedName    = $InitialValue
        Messages             = $InitialValue
    })
    $Results = @()
    If (-not $GetADSite.IsPresent) {$OutputTemplate.PSObject.Properties.Remove("Site")}
    If (-not $TestConnection.IsPresent) {$OutputTemplate.PSObject.Properties.Remove("IsAlive")}

    #endregion Output

    # Console output
    $BGColor = $Host.UI.RawUI.BackgroundColor
    $Msg = "Action: $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg}

    If ($GetADSite.IsPresent) {
        $Msg = "Retrieving AD site information will slow down processing, especially with a large number of computer names"
        Write-Warning $Msg
    }
    If ($TestConnection.IsPresent) {
        $Msg = "Pinging computers will slow down processing, especially with a large number of computer names"
        Write-Warning $Msg
    }
    $Host.UI.WriteLine()
}
Process {
    
    $Total = $ComputerName.count
    $Current = 0

    Foreach ($Computer in $ComputerName) {
    
        $Msg = "Search for $Computer"
        Write-Verbose $Msg

        $Current ++
        $Param_WP.Status = $Computer
        $Param_WP.PercentComplete = ($Current/$Total*100)

        Try {
            $Msg = "Search ADSI"
            $Param_WP.CurrentOperation = $Msg
            Write-Progress @Param_WP

            $Searcher.Filter = "(&(objectCategory=Computer)(name=$Computer))"

            If (($Found = $Searcher.FindAll()).Count -gt 0) {
                
                $Msg = "$($Found.Count) computer object(s) found"
                Write-Verbose $Msg

                Foreach ($Obj in $Found) {
                
                    $NameStr = ($Obj.properties.name -as [string])
                    $Msg = $NameStr
                    #Write-Verbose $Msg

                    $Msg = "Create report object"
                    $Param_WP.CurrentOperation = $Msg
                    $Param_WP.Status = $NameStr
                    Write-Progress @Param_WP
                    
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

                    # DNS lookup
                    $DNSEntity = $Null
                    $DNSEntity = [Net.Dns]::GetHostEntry($NameStr)
                    [string[]]$Output.IPAddress = $DNSEntity.AddressList | Foreach-Object {$_.IPAddressToString}

                    # Convert last logon date
                    If (-not ($Output.LastLogon = [datetime]::FromFileTime($($Obj.Properties.lastlogontimestamp)))) {$Output.LastLogon = "(unknown)"}
                    
                    # Get or create FQDN
                    If (-not ($Output.DNSHostName = $Obj.Properties.dnshostname -as [string])) {
                        $Msg = "DNSHostName not available; constructing FQDN from DNS lookup"
                        If ($DNSEntity.HostName -match $DomainDNSRoot) {$Output.DNSHostName = $DNSEntity.HostName -as [string]}
                        Else {$Output.DNSHostName = "(unavailable)"}
                    }

                    # Description and location
                    If (-not ($Output.Description = ($Obj.properties.description -as [string]))) {$Output.Description = "(none)"}
                    If (-not ($Output.Location = $Obj.properties.location -as [string])) {$Output.Location = "(none)"}
                    
                    # Look up the AD site info if object is enabled
                    If ($GetADSite.IsPresent) {

                        If ($Output.IsEnabled -eq $True) {

                            $Msg = "Retrieve AD site name"
                            $Param_WP.CurrentOperation = $Msg
                            Write-Progress @Param_WP
                        
                            If (-not ($Output.Site = Get-ComputerSite -ComputerName $NameStr -ErrorAction SilentlyContinue)) {$Output.Site = "(unavailable)"}
                        }
                        Else {
                            $Msg = "Site name testing unavailable for disabled computer $NameStr"
                            $Host.UI.WriteErrorLine($Msg)
                            $Output.Site = "(Unavailable)"
                        }   
                    }

                    # Ping it if object is enabled
                    If ($TestConnection.IsPresent) {

                        If ($Output.IsEnabled -eq $True) {
                            $Msg = "Ping computer"
                            $Param_WP.CurrentOperation = $Msg
                            Write-Progress @Param_WP
                        
                            If ($DNSEntity -eq "(unavailable)") {$Output.IsAlive = "(unavailable)"}
                            Else {$Output.IsAlive = Test-Ping -ComputerName $FQDN}
                        }
                        Else {
                            $Msg = "Connection testing unavailable for disabled computer $NameStr"
                            $Host.UI.WriteErrorLine($Msg)
                            $Output.IsAlive = $False
                        }
                    }
                    
                    # Add it up
                    $Results += $Output

                }  #end foreach found

            } #end if found

            Else {
                
                $Msg = "No AD object match found"
                $Host.UI.WriteErrorLine("ERROR: $Msg for $Computer")
                $Output = $OutputTemplate.PSObject.Copy()
                $Output.Name = $Computer
                $Output.Messages = $Msg
                $Results += $Output
            }
        }
        Catch {
            $Msg = "ADSISearch failed for $Computer"
            If ($ErrorDetails = $_.exception.message) {$Msg = "$Msg`n$ErrorDetails"}
            $Host.UI.WriteErrorLine("ERROR: $Msg") 
            
            $Output = $OutputTemplate.PSObject.Copy()
            $Output.Name = $Computer
            $Output.Messages = $Msg
            $Results += $Output
        
        }

    } #end foreach computer
}
End {

    Write-Progress -Activity $Activity -Completed

    If ($Results.Count -gt 0) {
        Write-Output $Results
    }
    Else {
        $Msg = "No results found"
        $Host.UI.WriteErrorLine($Msg) 

    }
}
} #end Get-PKADComputerMiniReport
 

