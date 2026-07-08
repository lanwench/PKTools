#Requires -version 4.0
Function Convert-PKBytes{
<#
.SYNOPSIS
    Converts one or more byte counts to a human-readable unit, auto-detecting or using a specified target unit

.DESCRIPTION
    Determines the appropriate unit (B, KB, MB, GB, TB, or PB) for one or more given byte counts, or uses -TargetUnit to force a specific unit
    Optional -TargetUnit may be especially handy when piping multiple values that should all appear in the same unit for consistent comparison
    Returns a PSObject with the raw quotient (Output), rounded display string (Formatted), and the target unit (ConvertTo)
    Without -Round, Output is the raw quotient and Formatted shows two decimal places
    With -Round, Output and Formatted both use whole numbers
    Accepts pipeline input

.NOTES
    Name    : Function_Convert-PKBytes.ps1
    Created : 2018-02-13
    Author  : Paula Kingsley
    Version : 02.00
    History :

        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v01.00 - 2018-02-13 - Created script based on Boe Prox's original
        v02.00 - 2026-06-16 - Renamed, overhauled & added helper function, added Round param

.LINK
    https://learn-powershell.net/2010/08/29/convert-bytes-to-highest-available-unit/

.PARAMETER Bytes
    Integer to convert to readable format

.PARAMETER TargetUnit
    Target unit for conversion; if not specified, the highest readable unit is auto-detected
    Valid values: B, KB, MB, GB, TB, PB

.EXAMPLE
    PS C:\> Convert-PKBytes -Bytes 40000 -Verbose

        VERBOSE: PSBoundParameters:

        Key              Value
        ---              -----
        Bytes            40000
        Round            False
        ParameterSetName __AllParameterSets
        PipelineInput    False
        ScriptName       Convert-PKBytes
        ScriptVersion    02.00
        Verbose          True

        VERBOSE: [BEGIN: Convert-PKBytes] Convert byte sizes to human-readable strings
        VERBOSE: Converting 40000 to KB

        Bytes     : 40000
        ConvertTo : KB
        Rounding  : False
        Output    : 39.0625
        Formatted : 39.06 KB

        VERBOSE: [END: Convert-PKBytes] Script ran successfully

.EXAMPLE
    PS C:\> 1048576, 1073741824 | Convert-PKBytes -Round

        Bytes     : 1048576
        ConvertTo : MB
        Rounding  : True
        Output    : 1
        Formatted : 1 MB

        Bytes     : 1073741824
        ConvertTo : GB
        Rounding  : True
        Output    : 1
        Formatted : 1 GB

.EXAMPLE
    PS C:\> 1073741824, 5368709120 | Convert-PKBytes -TargetUnit MB

        Bytes     : 1073741824
        ConvertTo : MB
        Rounding  : False
        Output    : 1024
        Formatted : 1024 MB

        Bytes     : 5368709120
        ConvertTo : MB
        Rounding  : False
        Output    : 5120
        Formatted : 5120 MB

        # Auto-detect would have chosen GB for both; -TargetUnit MB forces a consistent unit

#>

[CmdletBinding()]
Param(
    [Parameter(
        Mandatory=$True,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        Position=0,
        HelpMessage = "One or more byte sizes (int64)"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Size","Sum")]
    [int64[]]$Bytes,

    [Parameter(
        Position = 1,
        HelpMessage = "Target unit for conversion; if not specified, the highest readable unit is auto-detected"
    )]
    [ValidateSet("B","KB","MB","GB","TB","PB")]
    [string]$TargetUnit,

    [Parameter(
        HelpMessage = "Return a round number"
    )]
    [switch]$Round

)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "02.00"

    # Capture before $PSBoundParameters is mutated by the enumeration below
    $TargetUnitSpecified = $PSBoundParameters.ContainsKey("TargetUnit")
    If ($TargetUnitSpecified) { $TargetUnit = $TargetUnit.ToUpper() }

    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    $ScriptName = $MyInvocation.MyCommand.Name
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object {$CurrentParams.keys -notContains $_} |
        Where-Object {Test-Path variable:$_} | ForEach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    Function _FormatSize {
        Param([double]$Value, [double]$Divisor)
        $Decimals = If ($Round) { 0 } Else { 2 }
        [math]::Round($Value / $Divisor, $Decimals)
    }
    If ($TargetUnit) {$Msg = "Convert byte count to $TargetUnit units"}
    Else {$Msg = "Convert byte count to auto-detected highest possible unit"}
    If ($Round) {$Msg += " (rounded)"}
    Write-Verbose "[BEGIN: $ScriptName] $Msg"

}
Process {

    ForEach ($Item in $Bytes) {
        If ($TargetUnitSpecified) {
            $Divisor = @{ B=1; KB=1KB; MB=1MB; GB=1GB; TB=1TB; PB=1PB }[$TargetUnit]
            $Unit = $TargetUnit
        }
        Else {
            $Divisor, $Unit = Switch ($Item) {
                {$_ -gt 1PB} { 1PB, "PB";    Break }
                {$_ -gt 1TB} { 1TB, "TB";    Break }
                {$_ -gt 1GB} { 1GB, "GB";    Break }
                {$_ -gt 1MB} { 1MB, "MB";    Break }
                {$_ -gt 1KB} { 1KB, "KB";    Break }
                Default       { 1,   "B" }
            }
        }
        $Rounded = _FormatSize $Item $Divisor
        $OutputValue = If ($Round) { $Rounded } Else { $Item / $Divisor }
        Write-Verbose "Converting $Item bytes to $Unit"
        [PSCustomObject]@{
            Bytes     = $Item
            ConvertTo = $Unit
            Rounding  = $Round
            Output    = $OutputValue
            Formatted = "$Rounded $Unit"
        }
    }
}
End {
    Write-Verbose "[END: $ScriptName] Script ran successfully"
}
} #end Convert-PKBytes
$Null = New-Alias -Name Convert-PKBytesToSize  -Value Convert-PKBytes -Force:$True -ErrorAction SilentlyContinue