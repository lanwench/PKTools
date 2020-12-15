#requires -Version 4
Function Get-PKADReplicationFailures {
<#
.SYNOPSIS
    Gets AD replication errors for domain controllers in a forest

.DESCRIPTION
    Gets AD replication errors for domain controllers in a forest
    Accepts pipeline input
    Outputs a PSObject

.NOTES 
    Name    : Function_Get-PKADReplicationFailures.ps1
    Author  : Paula Kingsley
    Created : 2020-06-30
    Version : 01.00.0000
    History :  
        
        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK ** 

        v01.00.0000 - 2020-06-30 - Created script

.LINK
    https://techcommunity.microsoft.com/t5/itops-talk-blog/powershell-basics-how-to-check-active-directory-replication/ba-p/326364

.PARAMETER Forest
    Forest name (default is current user's)

.PARAMETER Quiet
    Hide all non-verbose/non-error console output

.EXAMPLE
    PS C:\> 

#>
[CmdletBinding()]
Param(
   [Parameter(
        HelpMessage = "Forest name (default is current user's)"
    )]
    [ValidateNotNullOrEmpty()]
    $Forest = $env:USERDNSDOMAIN,

   [Parameter(
        HelpMessage = "Hide all non-verbose/non-error console output"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("SuppressConsoleOutput")]
    [switch]$Quiet
    
)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    $ScriptName = $MyInvocation.MyCommand.Name
    
    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"
    
    #region Functions

    # Function to get error details
    Function Get-Error {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message)
        If ($ErrorDetails = $_.Exception.Message) {$Message += " ($ErrorDetails)"}
        Write-Output $Message
    }

    # Function to write verbose message, collecting error data, and optional prefixes
    Function Write-MessageVerbose {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message,[switch]$PrefixPrerequisites,[switch]$PrefixError)
        If ($PrefixPrerequisites.IsPresent) {$Message = "[Prerequisites] $Message"}
        Elseif ($PrefixError.IsPresent) {$Message = "ERROR  :  $Message"}
        If ($ErrorDetails = $_.Exception.Message) {$Message += " ($ErrorDetails)"}
        Write-Verbose $Message
    }

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

    # Function to write an error as a string (no stacktrace), or an error, with optional prefixes
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

    # Function to write an error/warning, collecting error data, with optional prefixes
    Function Write-MessageWarning {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message,[switch]$PrefixPrerequisites,[switch]$PrefixError)
        If ($PrefixPrerequisites.IsPresent) {$Message = "[Prerequisites] $Message"}
        Elseif ($PrefixError.IsPresent) {$Message = "ERROR  :  $Message"}
        If ($ErrorDetails = $_.Exception.Message) {$Message += " ($ErrorDetails)"}
        Write-Warning $Message
    }
    
    #endregion Functions


    #region Prerequisites 

    # Not using #requires bc other functions in this module don't require AD
    If (-not ($Null = Get-Module ActiveDirectory -EA SilentlyContinue)) {
        $Msg = "Active Directory module not found!"
        $Msg | Write-MessageError -PrefixPrerequisites -PrefixError
        Break
    }

    #endregion Prerequisites


    # For console
    $Activity = "Get Active Directory replication failures"
    "[BEGIN: $ScriptName] $Activity" | Write-MessageInfo -FGColor Yellow -Title

}
Process {

    
    $Msg = "Get all domain controllers"
    "[$Forest] $Msg" | Write-MessageVerbose
    Write-Progress -Activity $Activity -CurrentOperation $Msg -Status $Forest

    Try {
        $DCs = Get-ADDomainController -Filter * -Server (Get-ADForest $Forest -EA Stop) -EA Stop | 
            Select Domain,Site,Name,Hostname,OperatingSystem,IPv4Address,IsglobalCatalog,SSLPort,LDAPPort
        
        [int]$Total = ($DCs -as [array]).Count
        [int]$Current = 0

        $Msg = "[$Forest] $Total domain controllers found"
        $Msg | Write-MessageVerbose


        Foreach ($DC in ($DCs | Select -first 3)) {
            
            $Current ++
            $Msg = "Get replication partner metadata (please wait)"
            "[$($DC.Hostname)] $Msg" | Write-MessageVerbose
            Write-Progress -Activity $Activity -CurrentOperation $Msg -Status $DC.Hostname -PercentComplete ($Current/$Total * 100)

            $Output = [pscustomobject] @{
                DomainController    = $DC.Hostname
                Domain              = $DC.Domain
                Site                = $DC.Site
                OperatingSystem     = $DC.OperatingSystem
                IPv4Address         = $DC.IPv4Address
                IsGlobalCatalog     = $DC.IsGlobalCatalog
                ReplicationPartners = $null
                LastReplication     = $null
                FailureCount        = $null
                FailureType         = $null
                FirstFailure        = $null
                ErrorMessages       = $null
            }

            Try {
                $Partners = (Get-ADReplicationPartnerMetadata -Target $DC -ErrorAction Stop)
                $Output.ReplicationPartners = $Partners.Partner
                $Output.LastReplication = $Partners.LastReplicationSuccess

                Write-Verbose "Get replication errors (please wait)"
                "[$($DC.Hostname)] $Msg" | Write-MessageVerbose
                Write-Progress -Activity $Activity -CurrentOperation $Msg -Status $DC.Hostname -PercentComplete ($Current/$Total * 100)

                Try {
                    $Failures = (Get-ADReplicationFailure -Target $DC -ErrorAction Stop)
                    $Output.FailureCount  = $Failures.FailureCount
                    $Output.FailureType = $Failures.FailureType
                    $Output.FirstFailure = $Failures.FirstFailureTime
                }
                Catch {
                    $Msg = "Failed to get replication error data ($($_.Exception.Message))"
                    Write-Warning "[$($DC.Hostname)] $Msg"
                    $Output.ErrorMessages = $Msg
                }
            }
            Catch {
                    $Msg = "Failed to get replication partner metadata ($($_.Exception.Message))"
                    Write-Warning "[$($DC.Hostname)] $Msg"
                    $Output.ErrorMessages = $Msg
            }

            Write-Output $Output
        }
    }
    Catch {
        Throw "[$Forest] $($_.Exception.Message)"
    }
}
End {
    
    "[END: $ScriptName] $Activity" | Write-MessageInfo -FGColor Yellow -Title
    Write-Progress -Activity * -Completed
}
}