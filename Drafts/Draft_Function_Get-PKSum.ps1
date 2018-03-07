Function Get-PKSum {
[cmdletbinding()]
Param(
    [Parameter(
        Mandatory = $True,
        Position = 0,
        ValueFromPipeline = $True
    )]
    [ValidateNotNullOrEmpty()]
    $InputObject,

    [switch]$Round
)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    # Detect pipeline input
    $PipelineInput = (-not $PSBoundParameters.ContainsKey("InputObject")) -and (-not $InputObject)
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    Function GetSize {
        If ($Size = ($InputObject | Measure-Object -sum -ErrorAction SilentlyContinue ).sum) {
            If ($Round.IsPresent) {[math]::round($Size)}
            Else {$Size}
        }
        Else {
            $Msg = "Invalid input"
            $Host.UI.WriteErrorLine($Msg)
        }
    }
}
Process {

    If ($PipelineInput) {$_ | GetSize}
    Else {GetSize $InputObject}

}
}