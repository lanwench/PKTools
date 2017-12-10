#requires -Version 3
function Show-PKSubnetMaskTable {
<#

.SYNOPSIS
    Displays a table of prefix lengths and dotted decimal subnet masks

.DESCRIPTION
    Displays a table of prefix lengths and dotted decimal subnet masks
    No input
    Outputs a PSObject

.NOTES
    Name    : Function_Show-PKSubnetMaskTable.ps1
    Created : 2017-12-09
    Author  : Paula Kingsley
    Version : 01.01.0000
    History :

        # PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK

        v01.00.0000 - 2017-12-09 - Created script after Richard Siddaway's original


.LINK
    http://itknowledgeexchange.techtarget.com/powershell/subnets-and-prefixes/

#>
[CmdletBinding()]
Param()
Begin {
    
     # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    #$Source = $PSCmdlet.ParameterSetName
    #$PipelineInput = (-not $PSBoundParameters.ContainsKey("VM")) -and (-not $VM)

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    #$CurrentParams.Add("ParameterSetName",$Source)
    #$CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

}
Process {

    Foreach ($PrefixLength in 8..30) {

        switch ($PrefixLength) {

            {$_ -gt 24} {
                $bin = ('1' * ($PrefixLength – 24)).PadRight(8, '0')
                $o4 = [convert]::ToInt32($bin.Trim(),2)
                $mask = "255.255.255.$o4"
                break
            }

            {$_ -eq 24}{
                $mask = '255.255.255.0'
                break
            }

            {$_ -gt 16 -and $_ -lt 24} {
                $bin = ('1' * ($PrefixLength – 16)).PadRight(8, '0')
                $o3 = [convert]::ToInt32($bin.Trim(),2)
                $mask = "255.255.$o3.0"
                break
            }

            {$_ -eq 16} {
                $mask = '255.255.0.0'
                break
            }

            {$_ -gt 8 -and $_ -lt 16} {
                $bin = ('1' * ($PrefixLength – 8)).PadRight(8, '0')
                $o2 = [convert]::ToInt32($bin.Trim(),2)
                $mask = "255.$o2.0.0"
                break
            }

            {$_ -eq 8}{
                $mask = '255.0.0.0'
                break
            }

            default {
                $mask = '0.0.0.0'
            }
        }
        
        New-Object -TypeName PSObject -Property  ([ordered] @{
            PrefixLength = $PrefixLength
            SubnetMask = $Mask
        })
    }
}
} #end Show-PKSubnetMaskTable