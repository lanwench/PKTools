#Requires -version 3
Function Get-PKIANAPorts {
<# 
.SYNOPSIS
    Uses Invoke-Webrequest to get a CSV file from iana.org and creates a PSObject of TCP/UDP names, port numbers, and descriptions 

.DESCRIPTION
    Generic function
    Accepts pipeline input
    Returns a PSobject

.NOTES        
    Name    : Function_Do-SomethingCool.ps1
    Created : 2020-02-21
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2020-02-21 - Created script

.PARAMETER URI
    URI for IANA TCP/UDP port description CSV file (default is https://www.iana.org/assignments/service-names-port-numbers/service-names-port-numbers.csv) 

.PARAMETER OutputGridview
    Send output to Grid View

.PARAMETER TimeoutSeconds
    Timeout in seconds for Invoke-WebRequest (default is 10)


.EXAMPLE
    PS C:\> Do-SomethingCool -ComputerName foo

        
#> 

[CmdletBinding(
    SupportsShouldProcess = $True,
    ConfirmImpact = "Medium"
)]
[OutputType([Array])]
Param (
    [Parameter(
        Position = 0,
        HelpMessage = "URI for IANA TCP/UDP port description CSV file"
    )]
    [ValidateNotNullOrEmpty()]
    [string] $URI = "https://www.iana.org/assignments/service-names-port-numbers/service-names-port-numbers.csv",

    [Parameter(
        HelpMessage = "Timeout in seconds for Invoke-WebRequest (default is 10)"
    )]
    [int]$TimeoutSeconds = 10,

    [Parameter(
        HelpMessage = "Hide unassigned, empty, or Discarded ports"
    )]
    [switch]$HideUnassigned,

    [Parameter(
        HelpMessage = "Show protocols: All, TCPOnly, UDPOnly (default All)"
    )]
    [ValidateSet("All","TCP","UDP")]
    [String]$Protocol = "All",

    [Parameter(
        HelpMessage = "Send output to Grid View"
    )]
    [switch]$OutputGridView
)

Begin { 
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput

    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    #region Functions

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

    #endregion Functions

    # Console output
    $Activity = "Get IANA.org list of service names and port numbers via CSV file with a $TimeoutSeconds`s timeout"
    If ($Protocol -ne "All") {$Activity += ", filtering on protocol $Protocol"}
    If ($HideUnassigned.IsPresent) {$Activity += ", hiding unassigned, empty, or discarded ports"}
    "[BEGIN: $ScriptName] $Activity" | Write-MessageInfo -FGColor Yellow -Title

} #end begin

Process {
    
    $Msg = "Get CSV file from IANA.orgURI '$URI'"
    Write-Verbose $Msg

    If ($PSCmdlet.ShouldProcess("`nGet CSV file (this can take a moment`n",$URI)) {
        
        Try {
            $GetCSV = Invoke-WebRequest -Uri $URI -UseBasicParsing -Method Get -TimeoutSec $TimeoutSeconds -ErrorAction Stop
            $CSVContent = $GetCSV.Content

            $HeaderConfirm = 'Service Name,Port Number,Transport Protocol,Description,Assignee,Contact,Registration Date,Modification Date,Reference,Service Code,Unauthorized Use Reported,Assignment Notes'
            If ($CSVContent -match $HeaderConfirm) {
        
                $Msg = "Converting CSV file to PSObject"
                #If ($Protocol -ne "All") {$Msg += ", protocol $Protocol"}
                #If ($HideUnassigned.IsPresent) {$Msg += ", hiding unassigned, empty, or discarded ports"}
                Write-Verbose $Msg
                $OutputObj = ConvertFrom-Csv $CSVContent | Select @{N="Name";E={$_."Service Name"}},@{N="Port";E={$_."Port Number"}},@{N="Protocol";E={$_."Transport Protocol"}},Description
                
                Switch ($Protocol) {
                    All {}
                    TCP {$OutputObj = ($OutputObj | Where-Object {$_.Protocol -eq "TCP"})}
                    UDP {$OutputObj = ($OutputObj | Where-Object {$_.Protocol -eq "UDP"})}
                }
                If ($HideUnassigned.IsPresent) {$OutputObj = ($OutputObj | Where-Object {($_.Name -and $_.Port) -and ("discard" -notin ($_.Name,$_.Port))})}
                
                If ($OutputGridView.IsPresent) {
                    $Msg = "Send output to grid view"    
                    Write-Verbose $Msg
                    Try {
                        $OutputObj | Out-Gridview -Title "$(Get-Date) IANA.org TCP and UDP port names and numbers" -ErrorAction Stop
                    }
                    Catch {
                        $Msg = "Out-Gridview failed ($($_.Exception.Message))" 
                        Write-Error $Msg
                    }
                }
                Else {
                    Write-Output $OutputObj
                }
            }
            Else {
                $Msg = "CSV file headers don't match expected headers; please verify CSV file manually. Expected headers are:`n'Service Name,Port Number,Transport Protocol,Description,Assignee,Contact,Registration Date,Modification Date,Reference,Service Code,Unauthorized Use Reported,Assignment Notes'"
                Write-Error $Msg
            }
        }
        Catch {
            $Msg = "Failed to get CSV file from URI ($($_.Exception.Message))" 
            Write-Error $Msg
        }
    }
    Else {
        $Msg = "Operation cancelled by user"
        Write-Verbose $Msg
    }
    
}
End {  
    
    "[END $ScriptName] $Activity" | Write-MessageInfo -FGColor Yellow -Title
}

} # end Do-SomethingCool

