#Requires -version 3
Function Convert-PKJiraMarkdownTable {
<# 
.SYNOPSIS
    Creates a Jira-flavored markdown table from a PSObject

.DESCRIPTION
    Creates a Jira-flavored markdown table from a PSObject
    Optionally includes a panel
    Accepts pipeline input
    Returns a string
    
.NOTES        
    Name    : Function_Convert-PKJiraMarkdownTable.ps1
    Created : 2018-12-11
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2018-10-23 - Created script

.PARAMETER InputObject
    PSObject to convert to markdown

.PARAMETER AddHeader
    Add a header (instead of a panel)

.PARAMETER HeaderText
    Text for header

.PARAMETER HeaderTextColor
    Color for header text (default is Black)

.PARAMETER HeaderStyle
    Header style: H1, H2 or H3 (default is H2)

.PARAMETER AddPanel
    Add a Jira panel (instead of a header)

.PARAMETER PanelTitleText
    Title for a Jira panel

.PARAMETER PanelTitleTextColor
    Color for panel title text (default is Black)

.PARAMETER PanelTitleBackgroundColor
    Background color for Jira panel title (default is none)

.PARAMETER PanelBorder
    "Panel border style: Dashed or Solid (default is Solid)"

.PARAMETER SuppressConsoleOutput
    Hide all non-verbose/non-error console output

.EXAMPLE
    PS C:\> Convert-PKJiraMarkdownTable -InputObject $Report -Verbose
        VERBOSE: PSBoundParameters: 
	
        Key                       Value                                                                                                    
        ---                       -----                                                                                                    
        InputObject               @{VMName=dev-SQL-5; CreateDate=2018-12-03 04:34:22 PM; VMNotes="SQL for "...   
        Verbose                   True                                                                                                     
        Header                    False                                                                                                    
        HeaderText                                                                                                                         
        HeaderTextColor           Black                                                                                                    
        HeaderStyle               H2                                                                                                       
        AddPanel                  False                                                                                                    
        PanelTitleText                                                                                                                     
        PanelTitleTextColor       Black                                                                                                    
        PanelTitleBackgroundColor LemonChiffon                                                                                             
        PanelBackgroundColor      (None)                                                                                                   
        PanelBorder               Solid                                                                                                    
        SuppressConsoleOutput     False                                                                                                    
        ScriptName                Convert-PKJiraMarkdownTable                                                                              
        ScriptVersion             1.0.0                                                                                                    

        Action: Convert PSObject to Jira-flavored markdown table

        ||*Item*||*Description*||
        |VMName|dev-SQL-5|
        |CreateDate|2018-12-03 04:34:22 PM|
        |VMNotes|SQL for ops team test
        Created date: Monday, December 3, 2018 8:25:06 AM
        Created by: DOMAIN\jbloggs|
        |VMVersion|v10|
        |VMDataCenter|Sacramento|
        |VMHost|vmhost-520.domain.local|
        |VMFolder|dev-db|
        |VMDataStore|DS2_C501SSD_03132018_VMC03_N1
        DS1_C501SSD_03232018_VMC03_N2
        DS2_C501SSD_03292018_VMC03_N2|
        |VMSizeGB|3,397.11 GB|
        |VMCluster|vmc03|
        |VMDisk|Hard disk 1 [SCSI controller 0:0] 40.00 GB
        Hard disk 2 [SCSI controller 1:0] 205.00 GB
        Hard disk 3 [SCSI controller 2:0] 1,024.00 GB
        Hard disk 4 [SCSI controller 3:0] 2,096.00 GB|
        |VMSCSIController|SCSI controller 0 (ParaVirtual)
        SCSI controller 1 (ParaVirtual)
        SCSI controller 2 (ParaVirtual)
        SCSI controller 3 (ParaVirtual)|
        |VMNetworkAdapter|Network adapter 1|
        |VMPortGroup|1176_db_in|
        |Memory|32 GB|
        |CPU|Count: 8
        TotalSockets: 2
        CoresPerSocket:4|
        |HostName|dev-SQL-5|
        |FQDNs|dev-SQL-5.domain.local|
        |OU|Dev/SQL |
        |ADDescription|SQL for ops team test|
        |ADLocation|Sacramento|
        |DistinguishedName|CN=dev-SQL-5,OU=Servers,OU=Dev,DC=domain,DC=local|
        |OSName|Microsoft Windows Server 2012 R2 Datacenter (Core)|
        |OSVersion|6.3.9600|
        |HardDrive|C:\ [c_os_install] 39.66 GB
        D:\ [d_d_TmpDB] 205.00 GB
        E:\ [e_e_Log] 1,024.00 GB
        F:\ [f_f_Data] 2,095.87 GB|
        |NetworkAlias|Ethernet0|
        |IPv4Address|10.61.176.26|
        |SubnetMask|255.255.254.0|
        |Gateway|10.61.176.1|
        |DNSServers|10.62.179.250
        10.61.179.250|
        |PTR|dev-SQL-5.domain.local|
        |VLANID|1176|
        |WSUSServer|http://wsus01.domain.local|
        |WSUSGroup|patching_dev_group1|
        |Messages||
        |JIRAIssue|DBA-9806|
        |URL|http://jira.domain.local/browse/DBA-9806|


        
#> 

[CmdletBinding(
    DefaultParameterSetName = "__DefaultParameterSet",
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
[OutputType([Array])]
Param (
    [Parameter(
        Position = 0,
        Mandatory=$True,
        ValueFromPipeline=$True,
        HelpMessage="PSObject to convert to markdown table"
    )]
    [ValidateNotNullOrEmpty()]
    [PSObject] $InputObject,

    [Parameter(
        ParameterSetName = "Header",
        HelpMessage = "Add leading header"
    )]
    [Switch] $AddHeader,

    [Parameter(
        ParameterSetName = "Header",
        HelpMessage = "Text for header",
        Mandatory = $True
    )]
    [ValidateNotNullOrEmpty()]
    [string] $HeaderText,

    [Parameter(
        ParameterSetName = "Header",
        HelpMessage="Color for header text (default is Black)"
    )]
    [ValidateSet('Black','BurntSienna','DarkRed','LemonChiffon','LightCyan','LightGoldenRod','LightGray','NielsenBlue','OliveGreen','White','YellowGreen')]
    [string] $HeaderTextColor = "Black",

    [Parameter(
        ParameterSetName = "Header",
        HelpMessage = "Header style: H1, H2 or H3 (default is H2)"
    )]
    [ValidateSet('H1','H2','H3')]
    [string] $HeaderStyle = "H2",

    [Parameter(
        ParameterSetName = "Panel",
        HelpMessage = "Enclose table in a panel"
    )]
    [Switch] $AddPanel,

    [Parameter(
        ParameterSetName = "Panel",
        HelpMessage = "Text for panel title",
        Mandatory = $True
    )]
    [ValidateNotNullOrEmpty()]
    [string] $PanelTitleText,
    
    [Parameter(
        ParameterSetName = "Panel",
        HelpMessage="Color for panel title text (default is Black)"
    )]
    [ValidateSet('Black','BurntSienna','DarkRed','LemonChiffon','LightCyan','LightGoldenRod','LightGray','NielsenBlue','OliveGreen','White','YellowGreen')]
    [string] $PanelTitleTextColor = "Black",

    [Parameter(
        ParameterSetName = "Panel",
        HelpMessage ="Background color for panel title (default LemonChiffon)"
    )]
    [ValidateSet('Black','BurntSienna','DarkRed','LemonChiffon','LightCyan','LightGoldenRod','LightGray','NielsenBlue','OliveGreen','White','YellowGreen')]
    [string] $PanelTitleBackgroundColor = "LemonChiffon",

    [Parameter(
        ParameterSetName = "Panel",
        HelpMessage = "Background color for Jira panel title (default is none)"
    )]
    [ValidateSet('(None)','Black','BurntSienna','DarkRed','LemonChiffon','LightCyan','LightGoldenRod','LightGray','NielsenBlue','OliveGreen','White','YellowGreen')]
    [string] $PanelBackgroundColor = "(None)",

    [Parameter(
        ParameterSetName = "Panel",
        HelpMessage = "Panel border style: Dashed or Solid (default is Solid)"
    )]
    [ValidateSet('Solid','Dashed')]
    [string] $PanelBorder = "Solid",

    [Parameter(
        HelpMessage = "Hide all non-verbose/non-error console output"
    )]
    [Switch] $SuppressConsoleOutput

)

Begin { 
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = (-not $PSBoundParameters.ContainsKey("ComputerName")) -and (-not $ComputerName)
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference    = "Continue"
    
    # Separate rows
    $Separator = "`n"

    # If we want a panel, create lookup table for colors
    Switch ($Source) {
        #Panel   {
            
            # Background/foreground colors for the panel title
            ##$PanelTitleFG = $ColorHash.Item($PanelTitleTextColor)
            #$PanelTitleBG = $ColorHash.Item($PanelTitleBackgroundColor)
        #}
        _DefaultParameterSetName  {}
        Default {
            
            # Lookup table - maybe sometime will modify so parameter takes color names or hex codes directly
            $ColorHash = @{}
            $ColorHash = [ordered]@{
                DarkRed        = "#8B0000"
                BurntSienna    = "9C3108"
                LemonChiffon   = "#FFFACD"
                LightGoldenRod = "#FAFAD2"
                YellowGreen    = "#9ACD32"
                LightCyan      = "#E0FFFF"
                OliveGreen     = "#808000"
                NielsenBlue    = "#009dd9"
                LightGray      = "#D3D3D3"
                White          = "#FFFFFF"
                Black          = "#000000"
            }
        }
    }

    # Console output
    $Activity = "Convert PSObject to Jira-flavored markdown table"
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "Action: $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg`n")}
    Else {Write-Verbose $Msg}


} #end begin

Process {
    
    
        # Create ordered hashtable from array
        $TableHash = [ordered]@{}
        $TableHash = [ordered]@{}
        $InputObject.PSObject.Properties | Where-Object {$_.MemberType -eq 'NoteProperty'} | Select Name,Value | Foreach-Object {
            $Msg = "$($_.Name) = $($_.Value)"
            #Write-Verbose $Msg
            $TableHash.Add($_.Name,$_.Value)
        }

        # Initialize string
        [string]$TableStr = $Null

        Switch ($Source) {
    
            Panel  {
                If ($PanelTitleBackgroundColor -ne '(None)') {
                    $TableStr += "{panel:title=$PanelTitleText|borderStyle=$Border|titleBGColor=$($ColorHash.Item($PanelTitleBackgroundColor))|titleColor=$($ColorHash.Item($PanelTitleTextColor))|bgColor=$($Colorhash.item($PanelBackgroundColor))}"
                }
                Else {
                    $TableStr += "{panel:title=$PanelTitleText|borderStyle=$Border|titleBGColor=$($ColorHash.Item($PanelTitleBackgroundColor))|titleColor=$($ColorHash.Item($PanelTitleTextColor))}"
                }
                $TableStr += $Separator
            }
            Header {
                $TableStr += "$($HeaderStyle.ToLower()). {color:$($ColorHash.Item($HeaderTextColor))}$HeaderText{color}"

                #$TableStr += "$($HeaderStyle.ToLower()). $HeaderText"
                $TableStr += $Separator
            }
        }

        # Create headers
        $TableStr += "||*Item*||*Description*||"
        $TableStr += $Separator

        # Add hashtable rows
        $TableHash.GetEnumerator() | Foreach-Object {
            $TableStr += "|$($_.Name)|$($_.Value)|"
            $TableStr += $Separator
        }

        If ($AddPanel.IsPresent) {
            $TableStr += "{panel}"
        }

        Write-Output $TableStr
    #}
        
}
End {}

} # end Convert-PKJiraMarkdownTable



<#

break

# Hashtable to Jira markdown table

[switch]$IsPanel = $True

# Input array (must be a list!)
$InputObject = $Report
$Header = "New VM report"


If ($IsPanel.IsPresent) {
    $Title = "VM build report"
}

# Create property selection order
$Select = $InputObject.PSObject.Properties.Name

# Create ordered hashtable from array
$Hash = [ordered]@{}
$InputObject.PSObject.Properties | Where-Object {$_.MemberType -eq 'NoteProperty'} | Select Name,Value | Foreach-Object {
    Write-Verbose "$($_.Name) = $($_.Value)"
    $Hash.Add($_.Name,$_.Value)
}

# Initialize table string
[string]$Table = $Null
$Separator = "`n"

#If you want a header, add this at the top
$Table += "h2. $Header"
$Table += $Separator

# If we want a panel
If ($IsPanel.IsPresent) {

    $Title = "My panel title"

    $ColorHash = @{}
    $ColorHash = [ordered]@{
        DarkRed        = "#8B0000"
        BurntSienna    = "9C3108"
        LemonChiffon   = "#FFFACD"
        LightGoldenRod = "#FAFAD2"
        YellowGreen    = "#9ACD32"
        LightCyan      = "#E0FFFF"
        OliveGreen     = "#808000"
        NielsenBlue    = "#009dd9"
        LightGray      = "#D3D3D3"
        White          = "#FFFFFF"
        Black          = "#000000"
    }

    # Background/foreground colors for the panel title
    $PanelTitleBG = "NielsenBlue"
    $PanelTitleFG = "White"

    # Background color for the panel
    $PanelBG = "LemonChiffon"

    # Solid or dashed border
    $Border = "solid" # "dashed"
    
    # Add the panel
    $Table += "{panel:title=$Title|borderStyle=$Border|titleBGColor=$($Colorhash.item($PanelTitleBG))|titleColor=$($Colorhash.item($PanelTitleFG))|bgColor=$($Colorhash.item($PanelBG))}"
    $Table += $Separator
}


# Create headers
$Table += "||*Item*||*Description*||"
$Table += $Separator

# Add hashtable rows
$Hash.GetEnumerator() | Foreach-Object {
    $Table += "|$($_.Name)|$($_.Value)|"
    $Table += $Separator
}

If ($IsPanel.IsPresent) {
    $Table += "{panel}"
}



#>

