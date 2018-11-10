# Function to get AD object info using ADSISearcher (no ActiveDirectory module needed)
# https://social.technet.microsoft.com/Forums/scriptcenter/en-US/22f27c8c-4389-400b-b851-11d7a1ceb527/use-adsisearcher-on-another-domain?forum=ITCG

    Function GetADComputer {
        [Cmdletbinding(
            SupportsShouldProcess = $True,
            ConfirmImpact = "High"
        )]
        Param(
            [Parameter(Mandatory=$True)]
            $ComputerName,
            $ADDomain,
            $Server,
            $SizeLimit = 100,
            $OUTrimDepth,
            $Credential
        )
        $ErrorActionPreference = "Stop"
        If ($ComputerName -match "\*") {
            $Msg = "Wildcards are not permitted in computer names"
            $Host.UI.WriteErrorLine($Msg)
            Break
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

        If ($Credential) {
            $Context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain",$ADDomain,$Credential.UserName,$Credential.GetNetworkCredential().Password)
        }
        Else {
            $Context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain",$ADDomain)
        }
        $DomainObj = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($Context)
        $Root = $DomainObj.GetDirectoryEntry()
        $Search = [adsisearcher]$Root
        $Search.Filter = "(&(objectCategory=Computer)(name=$ComputerName))"
        $SearchResults = $Search.FindOne()
        $DirectoryEntry = $SearchResults.GetDirectoryEntry()



        Try {
        
            $ADSISearch = [adsisearcher]"(&(objectCategory=Computer)(name=$ComputerName))"
            $ADSISearch.SizeLimit = $SizeLimit
            $ADSISearch.PropertiesToLoad.AddRange(('name','canonicalname','serviceprincipalname','useraccountcontrol','lastlogonTimeStamp','distinguishedname','description','operatingsystem','location','whencreated','dnshostname'))
            
            If ($Credential) {
                $Domain = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList $DomainDN ,$($Credential.UserName),$($Credential.GetNetworkCredential().password)
                $ADSISearch.SearchRoot = $Domain
            }
            
            Try {
                $ADSISearch.Filter = "(&(objectCategory=Computer)(name=$ComputerName))"
                $Obj = $ADSISearch.FindAll()

                Switch ($Obj.Count) {
                    0 {
                        $Msg = "Computer object '$ComputerName' not found"
                        $Host.UI.WriteErrorLine("ERROR: $Msg") 
                        Break
                    }
                    Default {
                        $Msg = "Computer object '$ComputerName' not found"
                        $Host.UI.WriteErrorLine("ERROR: $Msg") 
                        Break
                    }
                    1 {
                        $NameStr = $($Obj.Properties.name -as [string])
                        $Msg = "Computer object '$NameStr' found"
                        Write-Verbose $Msg

                        $Results = @()
                        $InitialValue = "Error"
                        $Output = New-Object -TypeName PSObject -Property ([ordered] @{
                            Name                 = $NameStr
                            OU                   = $InitialValue
                            IPAddress            = $InitialValue
                            DNSHostName          = $InitialValue
                            Description          = $InitialValue
                            OperatingSystem      = $InitialValue
                            Location             = $InitialValue
                            WhenCreated          = $InitialValue
                            LastLogon            = $InitialValue
                            ServicePrincipalName = $InitialValue
                            CanonicalName        = $InitialValue
                            DistinguishedName    = $InitialValue
                            Messages             = $InitialValue
                        })
      
                       If ($RemoveADComputer.IsPresent) {
            
                            $OU = (($Obj.Properties.distinguishedname -as [string])).split(",",2)[1]
                                
                            If ($PSCmdlet.ShouldProcess("`n`n`tDelete AD object '$NameStr'`n`tfrom OU`n`t$OU`n`n")) {
                                Try {
                                    $Obj.GetDirectoryEntry() | Foreach-Object {$_.DeleteObject(0)}
                                    $Msg = "Computer object '$NameStr' deleted from '$OU'"
                                    Write-Verbose $Msg
                                    $True
                                }
                                Catch {
                                    $Msg = "Object deletion failed for $ComputerName"
                                    If ($Errordetails = $_.Exception.Message) {$Msg = "$Msg`n$ErrorDetails"}
                                    $Host.UI.WriteErrorLine("ERROR: $Msg")
                                    $False
                                }
                            }
                            Else {
                                $Msg = "Object removal cancelled by user for '$($Obj.Properties.name -as [string])'"
                                Write-Verbose $Msg
                                $False
                            }

                        } #end if removing
                        Else {
                            
                            $InitialValue = "Error"
                            $Output = New-Object -TypeName PSObject -Property ([ordered] @{
                                Name                 = $NameStr
                                OU                   = $Obj.Properties.operatingsystem -as [string]
                                IPAddress            = $Null
                                DNSHostName          = $Obj.Properties.dnshostname -as [string]
                                Description          = $Obj.properties.description -as [string]
                                OperatingSystem      = $Obj.Properties.operatingsystem -as [string]
                                Location             = $Obj.properties.location -as [string]
                                WhenCreated          =  (($Obj.Properties.whencreated -as [string]) -as [datetime]).ToString()
                                LastLogon            = [datetime]::FromFileTime($($Obj.Properties.lastlogontimestamp))
                                ServicePrincipalName = $Obj.Properties.serviceprincipalname
                                CanonicalName        = $Obj.properties.canonicalname -as [string]
                                DistinguishedName    = $Obj.properties.distinguishedname -as [string]
                                Messages             = $Null
                            })
    
                    
                            # If we're trimming out some of the OU path string
                            $CN = ($Obj.properties.canonicalname -as [string])
                            If ($OUTrimDepth) {
                                $Output.OU = (($CN.Split("/",$OUTrimDepth)[-1]) -Split("/$NameStr") -as [string])
                            }
                            Else {
                                $Output.OU = $CN.Substring(0, $CN.lastIndexOf('/'))
                            }

                            # DNS lookup
                            If ($DNSEntity = Get-DNS -Name $NameStr) {[string[]]$Output.IPAddress = $DNSEntity.AddressList | Foreach-Object {$_.IPAddressToString}}
                            Else {
                                $Msg = "DNS lookup failed for $NameStr; no IP address information available"
                                $Output.Messages = $Msg
                            }
                    
                            Write-Output $Output

                        } #end if reporting
                    } #end if one object found
                } #end switch
            }
            Catch {
                $Msg = "Computer object '$ComputerName' not found"
                If ($ErrorDetails = $($_.exception.message -replace '\s+', ' ')) {$Msg = "$Msg`n$ErrorDetails"}
                #$Host.UI.WriteErrorLine("ERROR: $Msg") 
                Write-Verbose $Msg
                Break
            }
        }
        Catch {
            $Msg = "DirectorySearcher lookup failed"
            If ($ErrorDetails = $($_.exception.message -replace '\s+', ' ')) {$Msg = "$Msg`n$ErrorDetails"}
            $Host.UI.WriteErrorLine("ERROR: $Msg") 
            Break
        }

    } # end function GetADComputer

    Function TestADConnection{
        [Cmdletbinding()]
        Param(
            [Parameter(Mandatory=$True)]
            $ADDomain,
            $Server,
            $Credential
        )
        If ($Credential) {$User = $Credential.UserName}
        Else {$User = $Env:UserName}

        $Output = New-Object PSObject -Property ([ordered]@{
            Domain    = $ADDomain
            DomainDN  = "Error"
            Server    = "Error"
            UserName  = $User
            SiteName  = "Error"
            IsSuccess = $False
            Messages  = "Error"
        }) 

        $Messages = @()

        # Get the domain/DN
        Try {
            If ($Credential) {
                $Context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("domain",$ADDomain,$Credential.UserName,$Credential.GetNetworkCredential().Password) -ErrorAction Stop   
            }
            Else {
                $Context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("domain",$ADDomain) -ErrorAction Stop
            }
        
            If ($DomainObj = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($Context)) {

                $Output.DomainDN = $DomainObj.GetDirectoryEntry() | Select -ExpandProperty DistinguishedName
                $Msg = "Connected to Active Directory domain '$($DomainObj.Name)' as '$User'"
                Write-Verbose $Msg
                $Messages += $Msg

                Try {
                
                    If ($Server) {
                        
                        $Output.Server = $Server

                        If ($Credential) {
                            $Context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("DirectoryServer",$Server,$Credential.UserName,$Credential.GetNetworkCredential().Password) -ErrorAction Stop
                        }
                        Else {
                            $Context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("DirectoryServer",$Server) -ErrorAction Stop
                        }

                        If ($DC = [System.DirectoryServices.ActiveDirectory.DomainController]::GetDomainController($Context)) {
                            $Output.Server = $DC.Name
                            $Output.SiteName = $DC.SiteName
                            If ($DC.Domain.Name -eq $ADDomain) {
                                $Output.IsSuccess = $True
                                $Msg = "Verified Domain Controller '$($Server)' in domain '$($DomainObj.Name)'"
                                Write-Verbose $Msg
                                $Messages += $Msg
                            }
                        }
                        Else {
                            $Output.IsSuccess = $False
                            $Msg = "Failed to verify Domain Controller '$($Server)' in domain '$($DomainObj.Name)'"
                            $Messages += $Msg
                        }
                    }
                    Else {
                        If ($DC = [System.DirectoryServices.ActiveDirectory.DomainController]::findone($Context)) {
                            $Output.Server = $DC.Name
                            $Output.SiteName = $DC.SiteName
                            $Output.IsSuccess = $True
                            $Msg = "Got first available Domain Controller"
                            Write-Verbose $Msg
                            $Messages += $Msg
                        }
                        Else {
                            $Output.IsSuccess = $False
                            $Msg = "Failed to get first available Domain Controller"
                            $Host.UI.WriteErrorLine("ERROR: $Msg")
                            $Messages += $Msg
                        }
                    }
                }
                Catch {
                    $Output.IsSuccess = $False
                    If ($Server) {$Msg = "Failed to get Domain Controller '$Server'"}
                    Else {$Msg = "Failed to get Domain Controller"}
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
                    $Host.UI.WriteErrorLine("ERROR: $Msg")
                    $Messages += $Msg
                }
            }
            Else {
                $Output.IsSuccess = $False
                $Msg = "Failed to connect to Active Directory domain '$($ADDomain)' as '$User'"
                $Host.UI.WriteErrorLine("ERROR: $Msg")
                $Messages += $Msg
            }
        }
        Catch {
            $Output.IsSuccess = $False
            $Msg = "Failed to connect to Active Directory domain '$($ADDomain)' as '$User'"
            If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
            $Host.UI.WriteErrorLine("ERROR: $Msg")
            $Messages += $Msg
        }

        $Output.Messages = $Messages -join("`n")
        Write-Output $Output
    }

Function GetFrancoisADComputer {
<#

https://lazywinadmin.com/2013/10/powershell-get-domaincomputer-adsi.html


Note: I had to use backtick "`" to be able to fit the code in my blog. Backticks are not present in the final script.

To resume we added the following items :

* Error Handling

* TRY{"do tasks"} CATCH{"Oups Error"}

* Verbose

* [cmdletbinding()]

* Write-Verbose

* Support for Multiple ComputerName query

* [string[]]$ComputerName

* FOREACH ($item in $ComputerName)

* Support for Alternative Credential

* [System.Management.Automation.Credential()]$Credential = [System.Management.Automation.PSCredential]::Empty

* IF ($PSBoundParameters['Credential'])

* New-Object -TypeName System.DirectoryServices.DirectoryEntry-ArgumentList $DomainDN,$($Credential.UserName),$($Credential.GetNetworkCredential().password)

* Support for different Domain

* IF ($PSBoundParameters['DomainDN'])

* $Searcher.SearchRoot = $DomainDN

* OR

* New-Object -TypeName System.DirectoryServices.DirectoryEntry-ArgumentList $DomainDN,$($Credential.UserName),$($Credential.GetNetworkCredential().password)


#>
[CmdletBinding()]
PARAM(
    [Parameter(
        ValueFromPipelineByPropertyName=$true,
        ValueFromPipeline=$true
    )]
    [Alias("Computer")]
    [String[]]$ComputerName,
    
    [Alias("ResultLimit","Limit")]
    [int]$SizeLimit='100',
    
    [Parameter(
        ValueFromPipelineByPropertyName=$true
    )]
    [Alias("Domain")]
    [String]$DomainDN=$(([adsisearcher]"").Searchroot.path),

    [Alias("RunAs")]
    [System.Management.Automation.Credential()]
    $Credential = [System.Management.Automation.PSCredential]::Empty

)#PARAM

PROCESS{
    IF ($ComputerName){
    
    FOREACH ($item in $ComputerName){
    TRY{
        # Building the basic search object with some parameters
        Write-Verbose -Message "COMPUTERNAME: $item"
        $Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher `
                        -ErrorAction 'Stop' -ErrorVariable ErrProcessNewObjectSearcher
        $Searcher.Filter = "(&(objectCategory=Computer)(name=$item))"
        $Searcher.SizeLimit = $SizeLimit
        $Searcher.SearchRoot = $DomainDN

        # Specify a different domain to query
        IF ($PSBoundParameters['DomainDN']){
            IF ($DomainDN -notlike "LDAP://*") {$DomainDN = "LDAP://$DomainDN"}#IF
            Write-Verbose -Message "Different Domain specified: $DomainDN"
            $Searcher.SearchRoot = $DomainDN}#IF ($PSBoundParameters['DomainDN'])

        # Alternate Credentials
        IF ($PSBoundParameters['Credential']) {
            Write-Verbose -Message "Different Credential specified: $($Credential.UserName)"
            $Domain = New-Object -TypeName System.DirectoryServices.DirectoryEntry `
                -ArgumentList $DomainDN,$($Credential.UserName),$($Credential.GetNetworkCredential().password) `
                -ErrorAction 'Stop' -ErrorVariable ErrProcessNewObjectCred
            $Searcher.SearchRoot = $Domain}#IF ($PSBoundParameters['Credential'])

        # Querying the Active Directory
        Write-Verbose -Message "Starting the ADSI Search..."
        FOREACH ($Computer in $($Searcher.FindAll())){
            Write-Verbose -Message "$($Computer.properties.name)"
            New-Object -TypeName PSObject -ErrorAction 'Continue' `
                -ErrorVariable ErrProcessNewObjectOutput -Property @{
                "Name" = $($Computer.properties.name)
                "DNShostName"    = $($Computer.properties.dnshostname)
                "Description" = $($Computer.properties.description)
                "OperatingSystem"=$($Computer.Properties.operatingsystem)
                "WhenCreated" = $($Computer.properties.whencreated)
                "DistinguishedName" = $($Computer.properties.distinguishedname)}#New-Object
        }#FOREACH $Computer

        Write-Verbose -Message "ADSI Search completed"
    }#TRY
    CATCH{ 
        Write-Warning -Message ('{0}: {1}' -f $item, $_.Exception.Message)
        IF ($ErrProcessNewObjectSearcher){
            Write-Warning -Message "PROCESS BLOCK - Error during the creation of the searcher object"}
        IF ($ErrProcessNewObjectCred){
            Write-Warning -Message "PROCESS BLOCK - Error during the creation of the alternate credential object"}
        IF ($ErrProcessNewObjectOutput){
            Write-Warning -Message "PROCESS BLOCK - Error during the creation of the output object"}
    }#CATCH
}#FOREACH $item
}

}
} #end function



