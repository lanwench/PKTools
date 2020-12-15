#Requires -version 4
Function Convert-PKDistinguishedNameToJSON {
<# 
.SYNOPSIS
    Uses Zachary Loeber's Get-ChildOUStructure to output the CanonicalName format of a container/OU as JSON 

.DESCRIPTION
    Uses Zachary Loeber's Get-ChildOUStructure to output the CanonicalName format of a container/OU as JSON 
    Converts DistinguishedName format to CanonicalName if needed
    Accepts pipeline input
    Returns a JSON object

.NOTES        
    Name    : Function_Convert-PKDistinguishedNameToJSON.ps1
    Created : 2020-02-21
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2020-02-21 - Created script

.PARAMETER DistinguishedName
    DistinguishedName to parse (will be converted to CanonicalName format if needed)

.PARAMETER Depth
    Depth for JSON output (default is 20)

.PARAMETER Quiet
    Suppress non-verbose console output (outputs errors as errors, not formatted strings)

.EXAMPLE
    PS C:\> Convert-PKDistinguishedNameToJSON -DistinguishedName domain.com/toplevel/next/nowthis/andmore -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key               Value                                   
        ---               -----                                   
        DistinguishedName domain.com/toplevel/next/nowthis/andmore
        Verbose           True                                    
        Depth             20                                      
        Quiet             False                                   
        PipelineInput     False                                   
        ScriptName        Convert-PKDistinguishedNameToJSON       
        ScriptVersion     1.0.0                                   

        [BEGIN: Convert-PKDistinguishedNameToJSON] Create JSON output from DistinguishedName to 20 level(s)

        VERBOSE: domain.com/toplevel/next/nowthis/andmore
        VERBOSE: Processing Subtree for: domain.com
        VERBOSE: Processing Subtree for: toplevel
        VERBOSE: Processing Subtree for: next
        VERBOSE: Processing Subtree for: nowthis
        VERBOSE: Processing Subtree for: andmore
        {
            "path":  "domain.com",
            "name":  "domain.com",
            "children":  [
                             {
                                 "path":  "domain.com/toplevel",
                                 "name":  "toplevel",
                                 "children":  [
                                                  {
                                                      "path":  "domain.com/toplevel/next",
                                                      "name":  "next",
                                                      "children":  [
                                                                       {
                                                                           "path":  "domain.com/toplevel/next/nowthis",
                                                                           "name":  "nowthis",
                                                                           "children":  [
                                                                                            {
                                                                                                "path":  "domain.com/toplevel/next/nowthis/andmore",
                                                                                                "name":  "andmore",
                                                                                                "children":  null
                                                                                            }
                                                                                        ]
                                                                       }
                                                                   ]
                                                  }
                                              ]
                             }
                         ]
        }

        [END Convert-PKDistinguishedNameToJSON] Create JSON output from DistinguishedName to 20 level(s)

.EXAMPLE
    PS C:\> Get-ADComputer "CN=LAPTOP88,OU=Tech Lab,OU=NorthAm,OU=Workstations,DC=domain,DC=local" | Convert-PKDistinguishedNameToJSON -Depth 4 -Quiet

    {
        "path":  "domain.local",
        "name":  "domain.local",
        "children":  [
                         {
                             "path":  "domain.local/Workstations",
                             "name":  "Workstations",
                             "children":  [
                                              {
                                                  "path":  "domain.local/Workstations/NorthAm",
                                                  "name":  "NorthAm",
                                                  "children":  ""
                                              }
                                          ]
                         }
                     ]
    }


#> 

[CmdletBinding()]
Param (
    
    [Parameter(
        Position = 0,
        Mandatory = $True,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "DistinguishedName (will be converted to CanonicalName format if needed)"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$DistinguishedName,

    [Parameter(
        HelpMessage = "Depth for JSON output (default is 20)"
    )]
    [ValidateNotNullOrEmpty()]
    [int] $Depth = 20,

    [Parameter(
        HelpMessage = "Suppress non-verbose console output (outputs errors as errors, not formatted strings)"
    )]
    [Alias("SuppressConsoleOutput")]
    [Switch]$Quiet


)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    $Source = $PSCmdlet.ParameterSetName
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object {$CurrentParams.keys -notContains $_} | 
        Where-Object {Test-Path variable:$_}| Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    #region Functions
    # Mix, match, use, discard, whatever

    # Function to write a console message or a verbose message
    Function Write-MessageInfo {
        Param([Parameter(ValueFromPipeline)]$Message,$FGColor,[switch]$Title)
        $BGColor = $host.UI.RawUI.BackgroundColor
        If (-not $Quiet.IsPresent) {
            If ($Title.IsPresent) {$Message = "`n$Message`n"}
            $Host.UI.WriteLine($FGColor,$BGColor,"$Message")
        }
        Else {Write-Verbose "$Message"}
    }

    # Function to write an error as a string (no stacktrace), or an error, and options for prefix to string
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message,[switch]$PrefixPrerequisites,[switch]$PrefixError)
        If ($PrefixPrerequisites.IsPresent) {$Message = "[Prerequisites] $Message"}
        Elseif ($PrefixError.IsPresent) {$Message = "ERROR  :  $Message"}
        If ($ErrorDetails = $_.Exception.Message) {$Message += " ($ErrorDetails)"}
        If (-not $Quiet.IsPresent) {
            $Host.UI.WriteErrorLine("$Message")
        }
        Else {Write-Error "$Message"}
    }
    # Function to write a warning, with any error data, and options for prefix to string
    Function Write-MessageWarning {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message,[switch]$PrefixPrerequisites,[switch]$PrefixError)
        If ($PrefixPrerequisites.IsPresent) {$Message = "[Prerequisites] $Message"}
        Elseif ($PrefixError.IsPresent) {$Message = "ERROR  :  $Message"}
        If ($ErrorDetails = $_.Exception.Message) {$Message += " ($ErrorDetails)"}
        Write-Warning $Msg
    }

    # Function to write an error/warning, collecting error data
    Function Write-MessageVerbose {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message,[switch]$PrefixPrerequisites,[switch]$PrefixError)
        If ($PrefixPrerequisites.IsPresent) {$Message = "[Prerequisites] $Message"}
        Elseif ($PrefixError.IsPresent) {$Message = "ERROR  :  $Message"}
        If ($ErrorDetails = $_.Exception.Message) {$Message += " ($ErrorDetails)"}
        Write-Verbose $Message
    }

    # Convert DN to CN
    function Get-CanonicalName {
    Param (
        [Parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string[]]$DistinguishedName
    )     
        foreach ($dn in $DistinguishedName) {      
            $d = $dn.Split(',') ## Split the dn string up into it's constituent parts 
            $arr = (@(($d | Where-Object { $_ -notmatch 'DC=' }) | ForEach-Object { $_.Substring(3) }))  ## get parts excluding the parts relevant to the FQDN and trim off the dn syntax 
            [array]::Reverse($arr)  ## Flip the order of the array. 
 
            ## Create and return the string representation in canonical name format of the supplied DN 
            $("{0}/{1}" -f  (($d | Where-Object { $_ -match 'dc=' } | ForEach-Object { $_.Replace('DC=','') }) -join '.'), ($arr -join '/')).TrimEnd("/") 
        } 
    }

    # Function that actually does the stuff (Zachary Loeber)
    function Get-ChildOUStructure {
    <#
    .SYNOPSIS
        Create JSON exportable tree view of AD OU (or other) structures.

    .DESCRIPTION
        Create JSON exportable tree view of AD OU (or other) structures in Canonical Name format.

    .PARAMETER ouarray
        Array of OUs in CanonicalName format (ie. domain/ou1/ou2)

    .PARAMETER oubase
        Base of OU

    .EXAMPLE
        $OUs = @(Get-ADObject -Filter {(ObjectClass -eq "localanizationalUnit")} -Properties CanonicalName).CanonicalName
        $test = $OUs | Get-ChildOUStructure | ConvertTo-Json -Depth 20

    .NOTES
        Author: Zachary Loeber
        Requires: Powershell 3.0, Lync
        Version History
        1.0.0 - 12/24/2014
            - Initial release
    .LINK
        https://github.com/zloeber/Powershell/blob/master/ActiveDirectory/Get-ChildOUStructure.ps1
    .LINK
        http://www.the-little-things.net
    #>
    [CmdletBinding()]
    param(
        [Parameter(
            Position=0, 
            ValueFromPipeline=$true, 
            Mandatory=$true, 
            HelpMessage='Array of OUs in CanonicalName format (ie. domain/ou1/ou2)'
        )]
        [string[]]$ouarray,

        [Parameter(Position=1, HelpMessage='Base of OU.')]
        [string]$oubase = ''
    )
    begin {
        $newarray = @()
        $base = ''
        $firstset = $false
        $ouarraylist = @()
    }
    process {
        $ouarraylist += $ouarray
    }
    end {
        $ouarraylist = $ouarraylist | Where {($_ -ne $null) -and ($_ -ne '')} | Select -Unique | Sort-Object
        if ($ouarraylist.count -gt 0) {
            $ouarraylist | Foreach {
                # $prioroupath = if ($oubase -ne '') {$oubase + '/' + $_} else {''}
                $firstelement = @($_ -split '/')[0]
                $regex = "`^`($firstelement`?`)"
                $tmp = $_ -replace $regex,'' -replace "^(\/?)",''

                if (-not $firstset) {
                    $base = $firstelement
                    $firstset = $true
                }
                else {
                    if (($base -ne $firstelement) -or ($tmp -eq '')) {
                        Write-Verbose "Processing Subtree for: $base"
                        $fulloupath = if ($oubase -ne '') {$oubase + '/' + $base} else {$base}
                        New-Object psobject -Property @{
                            'name' = $base
                            'path' = $fulloupath
                            'children' = if ($newarray.Count -gt 0) {,@(Get-ChildOUStructure -ouarray $newarray -oubase $fulloupath)} else {$null}
                        }
                        $base = $firstelement
                        $newarray = @()
                        $firstset = $false
                    }
                }
                if ($tmp -ne '') {
                    $newarray += $tmp
                }
            }
            Write-Verbose "Processing Subtree for: $base"
            $fulloupath = if ($oubase -ne '') {$oubase + '/' + $base} else {$base}
            New-Object psobject -Property @{
                'name' = $base
                'path' = $fulloupath
                'children' = if ($newarray.Count -gt 0) {,@(Get-ChildOUStructure -ouarray $newarray -oubase $fulloupath)} else {$null}
            }
        }
    }
    }

    #endregion Functions

    #region Splats

    # Splat for write-progress
    $Activity = "Create JSON output from DistinguishedName to $Depth level(s)"
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = "Working"
    }

    #endregion Splats

    # Console output
    "[BEGIN: $Scriptname] $Activity" | Write-MessageInfo -FGColor Yellow -Title
    
} #end begin
Process {

    Foreach ($DN in $DistinguishedName) {
        
        $Param_WP.CurrentOperation = $DN
        Write-Progress @Param_WP
        $DN | Write-MessageVerbose
        
        If ($DN -match "CN=") {$DN = $DN | Get-CanonicalName}
        If ($DN -match "./") {
            Get-ChildOUStructure $DN | ConvertTo-Json -Depth $Depth           
        }
        Else {
            $Msg = "Invalid DistinguishedName or CanonicalName"
            $Msg | Write-MessageError
        }
    }
        
}
End {
    
    $Null = Write-Progress -Activity $Activity -Completed
    "[END $Scriptname] $Activity" | Write-MessageInfo -FGColor Yellow -Title
}

} # end Convert-PKDistinguishedNameToJSON

