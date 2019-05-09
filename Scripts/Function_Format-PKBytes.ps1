#requires -Version 3
Function Format-PKBytes {
<#
.SYNOPSIS
    Converts bytes to human-readable form--detecting B,KB,MB,GB,TB,PB--and returning a PSObject or string

.DESCRIPTION
    Converts bytes to human-readable form--detecting B,KB,MB,GB,TB,PB--and returning a PSObject or string
    Uses regular expressions
    Accepts pipeline input
    Returns a PSObject or string

.NOTES
    Name    : Function_Format-PKBytes.ps1
    Author  : Paula Kingsley
    Created : 2019-04-26
    Version : 01.00.0000    
    History :

        *** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00.0000 - 2019-04-26 - Created script

.PARAMETER Bytes
    Bytes to convert to readable format

.PARAMETER ReturnSizeOnly
    Return only string for size (default is PSObject)

.PARAMETER Quiet
    Suppress all non-verbose console output

.EXAMPLE
    PS C:\Users\JBloggs\Documents\Logfiles> (Get-ChildItem -Recurse | Measure-Object -Property Length -Sum).Sum | Format-PKSize -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key            Value        
        ---            -----        
        Verbose        True         
        Bytes                       
        ReturnSizeOnly False        
        Quiet          False        
        PipelineInput  True         
        ScriptName     Format-PKSize
        ScriptVersion  1.0.0        
        InputObject                 

        BEGIN  : Format bytes into human-readable form

          Bytes Size   
          ----- ----   
        1394087 1.33 MB

        END    : Format bytes into human-readable form

.EXAMPLE
    PS C:\> Format-PKBytes -Bytes 123456 -ReturnSizeOnly -Quiet

        120.56 KB

.EXAMPLE
    PS C:\> Format-PKBytes -Bytes "kittens"

        BEGIN  : Format bytes into human-readable form
        ERROR  : Cannot convert argument "a", with value: "kittens", for "Log" to type "System.Double": 
        "Cannot convert value "kittens" to type "System.Double". Error: "Input string was not in a correct format.""
        END    : Format bytes into human-readable form


#>
[Cmdletbinding()]
Param (
    [Parameter(
        Mandatory = $True,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True ,
        HelpMessage = "Bytes to convert to readable format"
    )]
    [Alias("Sum")]
    $Bytes,

    [Parameter(
        HelpMessage = "Return only string for size (default is PSObject)"
    )]
    [ValidateNotNullOrEmpty()]
    [switch]$ReturnSizeOnly,

    [Parameter(
        HelpMessage = "Suppress all non-verbose console output"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("SuppressConsoleOutput")]
    [switch]$Quiet
)
Begin {
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $CurrentParams = $PSBoundParameters
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    If ($PipelineInput.IsPresent) {$CurrentParams.InputObject = $Null}
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference    = "Continue"
    
    # Console output
    $Activity = "Format bytes into human-readable form"
    If ($ReturnSizeOnly.IsPresent) {$Activity += " (return string for size only)"}
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "BEGIN  : $Activity"
    $FGColor = "Yellow"
    If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg")}
    Else {Write-Verbose $Msg}

} 
Process {
    
    Try {
        $Size = switch -Regex ([math]::truncate([math]::log($Bytes,1024))) {
            '^0' {"$Bytes Bytes"}
            '^1' {"{0:n2} KB" -f ($Bytes / 1KB)}
            '^2' {"{0:n2} MB" -f ($Bytes / 1MB)}
            '^3' {"{0:n2} GB" -f ($Bytes / 1GB)}
            '^4' {"{0:n2} TB" -f ($Bytes / 1TB)}
             Default {"{0:n2} PB" -f ($Bytes / 1pb)}
        }
        If (-not $Size) {$Size = "Error"}
        If (-not $ReturnSizeOnly.IsPresent) {
             New-Object PSObject -Property ([ordered]@{
                    Bytes = $Bytes
                    Size  = $Size
                })
        }
        Else {
            $Size
        }
    }
    Catch {
        $Msg = ($_.Exception.Message)
        If (-not $Quiet.IsPresent) {$Host.UI.WriteErrorLine("ERROR  : $Msg")}
        Else {Write-Warning $Msg}
    }
}
End {
    
    $Msg = "END    : $Activity"
    $FGColor = "Yellow"
    If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg")}
    Else {Write-Verbose $Msg}
}
}

$Null =  New-Alias -Name Format-PKByteSize -Value Format-PKBytes -Confirm:$False -Force -ErrorAction SilentlyContinue -Description "For guessability"
$Null =  New-Alias -Name Format-PKSize -Value Format-PKBytes -Confirm:$False -Force -ErrorAction SilentlyContinue -Description "For guessability"