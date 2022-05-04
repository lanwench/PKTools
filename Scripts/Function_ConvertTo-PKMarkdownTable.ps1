#requires -Version 4
Function ConvertTo-PKMarkdownTable {
<#
.SYNOPSIS
    Converts a PSObject to a markdown table

.DESCRIPTION
    Converts a PSObject to a markdown table (with option to format as two-column list with custom header names)
    By default does not format top row / header as bold with || (this can be changed by using -BoldHeaders)
    The -AsList switch formats the table as a two-column list (works best for arrays containing only one object,
    which is often useful for reporting) and allows for custom column names (default is Item, Value)
    Properties containing collections can be formatted as strings separated by commas or line breaks for easier
    readability
    Accepts pipeline input
    Outputs an array of strings

.NOTES 
    Name    : Function_ConvertTo-PKMarkdownTable.ps1
    Created : 2022-01-22
    Author  : Paula Kingsley
    Version : 01.00.0000
    History:  
        
        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK ** 

        v1.00.00.0000 - 2022-01-22 - Created script

.EXAMPLE
    PS: C\> Get-ADDomainController -Filter * | Select Hostname,Site,IPv4Address,OperatingSystem,IsGlobalCatalog | ConvertTo-PKMarkdownTable -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key              Value                    
        ---              -----                    
        Verbose          True                     
        PSCustomObject                            
        BoldHeaders      False                    
        AsList           False                    
        Column1Name      Item                     
        Column2Name      Value                    
        CollectionJoin   Comma                    
        ParameterSetName AsTable                  
        PipelineInput    True                     
        ScriptName       ConvertTo-PKMarkdownTable
        ScriptVersion    1.0.0                    


        VERBOSE: Converting PSCustomObject to markdown table

        |Hostname|Site|IPv4Address|OperatingSystem|IsGlobalCatalog|
        |PARDC01.domain.local|Parsippany|10.30.78.28|Windows Server 2019 Datacenter|True|
        |HOHDC01.domain.local|Hohokus|10.21.32.239|Windows Server 2019 Datacenter|True|
        |PARDC02.domain.local|Parsippany|10.30.78.27|Windows Server 2019 Datacenter|True|
        |REDDC01.domain.local|RedBank|10.32.179.119|Windows Server 2019 Datacenter|True|
        |ALLDC01.domain.local|Allentown|10.18.80.97|Windows Server 2019 Datacenter|True|
        |NUTDC01.domain.local|Nutley|10.21.32.249|Windows Server 2019 Standard|True|
        |ELIDC03.domain.local|Elizabeth|10.5.163.15|Windows Server 2019 Standard|True|


.EXAMPLE
    PS C:\> Get-ADDomainController dc04.domain.local | Select-Object $Props | 
        ConvertTo-PKMarkdownTable -BoldHeaders -AsList -Column1Name Cats -Column2Name Dogs -CollectionJoin LineBreak

        ||*Cats*||*Dogs*||
        |Name|DC04|
        |Hostname|DC04.domain.local|
        |OperatingSystem|Windows Server 2019 Datacenter|
        |IPv4Address|10.11.32.7|
        |Site|Venice|
        |Forest|domain.local|
        |IsGlobalCatalog|True|
        |IsReadyOnly||
        |LDAPPort|389|
        |ServerObjectGuid|3351ab00-bc0e-447b-b0c9-e6c6c4ee3222|
        |SSLPort|636|
        |OperationMasterRoles|SchemaMaster
        DomainNamingMaster
        PDCEmulator
        RIDMaster
        InfrastructureMaster|
        |ComputerObjectDN|CN=DC04,OU=Domain Controllers,DC=domain,DC=local|


.EXAMPLE
    PS C:\> Get-Service | Select-Object -First 20 | Sort-Object Status -Descending | Group-Object Status | Select Name,@{N="Services";E={$_.Group.DisplayName -join("`n")}} | ConvertTo-PKMarkdownTable 
      
        |Name|Services|
        |Running|AppX Deployment Service (AppXSVC)
        Base Filtering Engine
        BitLocker Drive Encryption Service
        Windows Audio
        Windows Audio Endpoint Builder
        Background Intelligent Transfer Service
        Bluetooth User Support Service_26c749
        Application Information
        Application Host Helper Service|
        |Stopped|App Readiness
        Microsoft App-V Client
        AssignedAccessManager Service
        Application Identity
        Cellular Time
        ActiveX Installer (AxInstSV)
        GameDVR and Broadcast User Service_26c749
        Application Layer Gateway Service
        AllJoyn Router Service
        Application Management
        Agent Activation Runtime_26c749|

.PARAMETER PSCustomObject
    PSCustomObject to convert to a markdown table

.PARAMETER BoldHeaders
    Format top row/headers as bold (use * and ||)        

.PARAMETER Column1Name
    When using -AsList, name for Column 1 (default is 'Item')

.PARAMETER Column2Name
    When using -AsList, name for Column 1 (default is 'Value')

.PARAMETER CollectionJoin
    Collection properties will be joined as a string using a comma, line break or semicolon (default is comma)

.EXAMPLE
    PS C:\> Get-ADComputer -SearchBase "OU=Servers,DC=domain,DC=local" -Filter * -Properties Description,ManagedBy | Select Name,DNSHostName,Enabled,Description,DistinguishedName | Select -First 5 | ConvertTo-PKMarkdownTable -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key              Value                    
        ---              -----                    
        Verbose          True                     
        PSCustomObject                            
        BoldHeaders      False                    
        AsList           False                    
        Column1Name      Item                     
        Column2Name      Value                    
        CollectionJoin   Comma                    
        ParameterSetName AsTable                  
        PipelineInput    True                     
        ScriptName       ConvertTo-PKMarkdownTable
        ScriptVersion    1.0.0                    


        VERBOSE: Converting PSCustomObject to markdown table

        |Name|DNSHostName|Enabled|Description|DistinguishedName|
        |FILESERVER|fileserver.domain.local|True||CN=fileserver,OU=Servers,DC=domain,DC=local|
        |app1|APP1.domain.local|True||CN=app1,OU=Servers,DC=domain,DC=local|
        |sql-dev-3|sql-dev-3.domain.local|True||CN=sql-dev-3,OU=Servers,DC=domain,DC=local|

.EXAMPLE
    Get-ADUser ladygaga | ConvertTo-PKMarkdownTable -BoldHeaders -AsList -Column1Name Type -Column2Name Description -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key              Value                    
        ---              -----                    
        BoldHeaders      True                     
        AsList           True                     
        Column1Name      Type                     
        Column2Name      Description              
        Verbose          True                     
        PSCustomObject                            
        CollectionJoin   Comma                    
        ParameterSetName List                     
        PipelineInput    True                     
        ScriptName       ConvertTo-PKMarkdownTable
        ScriptVersion    1.0.0                    

        VERBOSE: Converting PSCustomObject to two-column markdown table

        ||*Type*||*Description*||
        |DistinguishedName|CN=Lady Gaga,OU=Accounts,DC=domain,DC=local|
        |Enabled|True|
        |GivenName|Lady|
        |Name|Lady Gaga|
        |ObjectClass|user|
        |ObjectGUID|1e4cf0fb-5657-457b-8891-7b7df5033cda|
        |SamAccountName|ladygaga|
        |SID|S-1-5-21-2523766570-964973733-570019136-99781|
        |Surname|Gaga|
        |UserPrincipalName|ladygaga@domain.local|
        

#>
[CmdletBinding(SupportsShouldProcess,ConfirmImpact = "Low",DefaultParameterSetName = "AsTable")]
Param(
    [Parameter(
        Mandatory = $True,
        Position = 0,
        ValueFromPipeline = $True,
        HelpMessage = "PSCustomObject to convert to a markdown table"
    )]
    [PSCustomObject[]]$PSCustomObject,

    [Parameter(
        HelpMessage = "Display top row/headers as bold using * and ||"
    )]
    [switch]$BoldHeaders,

    [Parameter(
        ParameterSetName = "List",
        HelpMessage = "Format object as two-column list/table"
    )]
    [switch]$AsList,
    
    [Parameter(
        ParameterSetName = "List",
        HelpMessage = "When using -AsList, name for Column 1 (default is 'Item')"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$Column1Name = "Item",

    [Parameter(
        ParameterSetName = "List",
        HelpMessage = "When using -AsList, name for Column 2 (default is 'Value')"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$Column2Name = "Value",

    [Parameter(
        HelpMessage = "Collection properties will be joined as a string using a comma, line break or semicolon (default is comma)"
    )]
    [ValidateSet("Comma","LineBreak","Semicolon")]
    [string]$CollectionJoin = "Comma"

)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    If (-not $AsList.IsPresent) {
        $Source = "AsTable"
    }
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    $ScriptName = $MyInvocation.MyCommand.Name
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # We can display collections as strings joined by commas or line breaks
    Switch ($CollectionJoin) {
        Comma {$Joiner = ", "; $JoinLabel = $CollectionJoin.ToLower()}
        LineBreak {$Joiner = " `n"; $JoinLabel = "line break"}
        Semicolon {$Joiner = "; "; $JoinLabel = $CollectionJoin.ToLower()}
    }

    # Initialize output array & create an arraylist so we can 
    # quickly add the input object to it
    $Output = @()   
    $Interim = [System.Collections.ArrayList]::new()
    
}
Process {
    
    # Adding the object to an interim object so we don't unroll arrays in the pipeline
    # and end up with a header + row for each row
    # All the actual work will be in End
    If ($PSCustomObject -is [array]) {
        $Null = $PSCustomObject | Foreach-Object {$Interim.Add($_)}
    } 
    Else {
        $Null = $Interim.Add($PSCustomObject)
    }
}
End {
    
    # Get the property names
    $Properties = ($Interim[0].PSObject.Properties.Name)

    If ($Interim.Count -gt 1 -and $AsList.IsPresent) {
        $Msg = "The -AsList parameter outputs a table with two columns for name/value,  and may not provide the best output for array objects containing multiple items"
        Write-Warning $Msg
    }
    Else { #if ($Interim.Count -eq 1) {
        $Collections = @()
        Foreach ($P in $Properties) {
            If (($Interim.$P).Count -gt 1) {
                $Collections += $P
            }
        }
        If ($Collections.Count -gt 0) {
            $Msg = "Your input object has one or more properties containing collections; these will be joined with a $Joinlabel" 
            If ($AsList.IsPresent) {
                Write-Verbose $Msg
            }
            Else {
                $Msg += ", but you should probably select -AsList if you want this to be readable!"
                Write-Warning $Msg
            }
        }
    }

    # Convert the interim list back to a pscustomobject, which is insanity, but apparently we can't update property values in the list
    $ConvergedObject = [PSCustomObject]@{}
    Foreach ($P in $Properties) {
        $ConvergedObject | Add-Member -MemberType NoteProperty -Name $P -Value ($Interim.$P -join($Joiner))
    }
    
    # Initialize array for output table rows
    [array]$Table = @()

    If ($AsList.IsPresent) {
        
        $Msg = "Converting PSCustomObject to two-column markdown table"
        Write-Verbose $Msg

        # Get the names of the propeties
        $Header = $SortOrder = $Properties

        # Add header row to array
        If ($BoldHeaders.IsPresent) {$Table += "||*$Column1Name*||*$Column2Name*||"}
        Else {$Table += "|$Column1Name|$Column2Name|"}

        # Add the rest
        Foreach ($P in $Properties) {
            $Table += "|$P|$($ConvergedObject.$P)|"
        }
        $Output += $Table

    }
    Else {
        $Msg = "Converting PSCustomObject to markdown table"
        Write-Verbose $Msg

        # Create the header row
        If ($BoldHeaders.IsPresent) {$Table += "||*$($Properties -join("*||*"))*||"}
        Else {$Table += "|$($Properties -join("|"))|"}
        
        # Add the rest of the rows to the table, using convertto-csv because we are both sneaky and lazy
        $Table += "|$((($ConvergedObject | ConvertTo-Csv -Delimiter "|" -NoTypeInformation) | Select-Object -Skip 1).Replace('"',$Null))|"
        $Output += $Table
    }

    Write-Output $Output
}
} #end ConvertTo-PKMarkdownTable


