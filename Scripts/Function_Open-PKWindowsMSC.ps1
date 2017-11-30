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
        ParameterSetName = "Domain",
        Mandatory = $False,
        Position = 2,
        HelpMessage = "Credentials (for domain tools only)"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential] $Credential,

    [Parameter(
        Mandatory   = $False,
        Position = 4,
        HelpMessage = "Suppress all non-verbose / non-error console output"
    )]
    [switch] $SuppressConsoleOutput

)


DynamicParam{

    # Set the dynamic parameters' name
    $ParamName_Local = 'Local'
    # Create the collection of attributes
    $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
    # Create and set the parameters' attributes
    $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
    $ParameterAttribute.Mandatory = $true
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
    $ParameterAttribute.Mandatory = $true
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
  

Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

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

    # Here-String for selecting filename/descriptive name from selection
    $Object = @'
"Item","Name","File"
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

    
    # Normalize the parameter
    Switch ($Source){
        Local {$Option = $CurrentParams.Local}
        Domain {$Option = $CurrentParams.Domain}
    }
    $Selection = $Object | Where-Object {$_.Item -eq $Option}

    #region Splats

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose = $False
    }

    # Splat for write-progress
    $Activity = "Launching $($Selection.Name) ($($Selection.File))"
    
    $Param_WP = @{}
    $Param_WP = @{
        Activity = $Null
        Status   = "Working"
    }
    
    # Splat for Start-Process
    $Param_Launch = @{}
    $Param_Launch = @{
        FilePath         = "$Env:WinDir\System32\cmd.exe"
        ArgumentList     = "/c $($Selection.File)"
        WorkingDirectory = $PSHome
        NoNewWindow      = $True
        ErrorAction      = "Stop"
        Verbose          = $False
    }
    If ($CurrentParams.Credential) {
        
        # UPNs don't work. 
        If ($Credential.UserName -match "@") {

            # Get the username part of the UPN
            $UserName = ($Credential.UserName -split "@")[0]

            # Get AD Root
            $RootDSE = [ADSI]"LDAP://RootDSE"
            $RootDSEConfig = $RootDSE.Get("configurationNamingContext")
            $ADSearchRoot=New-object System.DirectoryServices.DirectoryEntry("LDAP://CN=Partitions," + $RootDSEConfig)

            # Search for Netbiosname of the specified domain
            $DomainFQDN = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name
            $SearchString = "(&(objectclass=Crossref)(dnsRoot="+$DomainFQDN+")(netBIOSName=*))"
            $Search = New-Object directoryservices.DirectorySearcher($ADSearchRoot,$SearchString)
            $sNetBIOSName = ($Search.FindOne()).Properties["netbiosname"]

            # Set the username as domain\username
            $UserName = "$sNetBIOSName\$UserName"

            # Create a new object
            $InternalCred = New-Object System.Management.Automation.PSCredential -ArgumentList @($UserName,$AdminCredential.Password)
        }
        Else {
            $InternalCred = $Credential
            $Param_Launch.Add("Credential",$InternalCred)
            
        }
        $Activity = "$Activity as $UserName"
    }

    #endregion Splats

     # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "$Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg}
    
}
Process {

    $Param_WP.Activity = $Activity 
    Write-Progress @Param_WP
    Try {
        #$Param_Launch.ArgumentList = $Selection.File
        Start-Process @Param_Launch
        
    }
    Catch {
        $Msg = "Operation failed"
        If ($_.Exception.Message) {$Msg = "$Msg`n$($_.Exception.Message)"}
        $Host.UI.WriteErrorLine("ERROR: $Msg")
    }
  

}
End {
    Write-Progress -Activity $Activity -Completed
}
} #end Open-PKWindowsMSC


