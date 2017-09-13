#requires -RunAsAdministrator 
#requires -Version 3 
function Get-PKWindowsLogonInfo {
<#
.Synopsis

.Description

.Notes
    


#>
[cmdletbinding()]
param(
    [Int]$Newest = [Int]::MaxValue,
    [DateTime]$Before = (Get-Date),
    [DateTime]$After = (Get-Date).AddHours(24),
    [string[]]$ComputerName = $Env:ComputerName,
    $Authentication = '*',
    $User = '*',
    $Path = '*'
)
Begin {
    
    #Write-Verbose $MyInvocation.MyCommand.Name

    [version]$Version = "01.00.0000"

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    #$CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    #$CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams| Format-Table -AutoSize | out-string )"

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference = "SilentlyContinue"

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose = $False
    }

    $null = $PSBoundParameters.Remove('Authentication')
    $null = $PSBoundParameters.Remove('User')
    $null = $PSBoundParameters.Remove('Path')
    $null = $PSBoundParameters.Remove('Newest')


}

 
    
Process {
        
    $Msg = "Get event log data for $Env:ComputerName"
    Write-Verbose $Msg
    
    Get-EventLog -ComputerName $ComputerName -LogName Security -InstanceId 4624 @PSBoundParameters |
    
        ForEach-Object {
            [PSCustomObject]@{
                Time           = $_.TimeGenerated
                User           = $_.ReplacementStrings[5]
                Domain         = $_.ReplacementStrings[6]
                Path           = $_.ReplacementStrings[17]
                Authentication = $_.ReplacementStrings[10]
            } 
        } | 
        Where-Object Path -like $Path |
        Where-Object User -like $User |
        Where-Object Authentication -like $Authentication |
        Select-Object -First $Newest
}
}
 <##
 
$yesterday = (Get-Date).AddDays(-1)
Get-LogonInfo -After $yesterday |
Out-GridView

#>