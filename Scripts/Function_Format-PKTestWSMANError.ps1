#Requires -version 3
Function Format-PKTestWSMANError {
<# 
.SYNOPSIS
    Formats error messages from Test-WSMAN into human-readable strings

.DESCRIPTION
    Formats error messages from Test-WSMAN into human-readable strings
    Accepts pipeline input
    Returns a string

.NOTES        
    Name    : Function_Format-PKTestWSMANError.ps1
    Created : 2019-03-11
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2019-03-11 - Created script

.PARAMETER Message
    Error message output from Test-WSMAN

.PARAMETER SuppressConsoleOutput
    Hide all non-verbose/non-error console output

.EXAMPLE
    PS C:\> Do-SomethingCool -ComputerName foo

        
#> 

[CmdletBinding()]
Param (
    
    [Parameter(
        ValueFromPipeline = $True,
        Mandatory = $True,
        HelpMessage="Error message output from Test-WSMAN"
    )]
    [ValidateNotNullOrEmpty()]
    [string] $Message,

    [Parameter(
        HelpMessage="Hide all non-verbose/non-error console output"
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
    
    # Console output
    $Activity = "Convert error output from Test-WSMAN into readable message"
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "BEGIN  : $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg`n")}
    Else {Write-Verbose $Msg}


} #end begin

Process {

   If ($Message -match "<f:WSManFault ") {

        #parse the product version line into separate properties
        [regex]$rx = "OS:\s(?\S+)\sSP:\s(?\d\.\d)\sStack:\s(?\d\.\d)"
        $pv = $Message.productVersion
        
        [string]$os = ($rx.Matches($pv)).foreach({$_.groups["OS"].value})
 
        #force these as strings and later treat them as [decimal]
        [string]$sp = ($rx.Matches($pv)).foreach({$_.groups["SP"].value})
        [string]$stack = ($rx.Matches($pv)).foreach({$_.groups["Stack"].value})


        [regex]:: match($Message,'(?<=\<f\:Message\>).+(?=\<\/f\:Message\>)',"singleline").value.trim()



    }
    Else {
        $Msg = "Message does not appear to be a Test-WSMan error"
        $Host.UI.WriteErrorLine($Msg)
    }
        
}
End {

    $Activity = "Convert error output from Test-WSMAN into readable message"
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "END    : $Activity"
    $FGColor = "Green"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg`n")}
    Else {Write-Verbose $Msg}
   
    
}

} # end Format-PKTestWSMANError


<#

Try {

    $TEst = Test-WSMAN -ComputerName acr-content-1 -Authentication Kerberos -ErrorAction Stop
}
Catch {
    $Message = $_.Exception.Message
}




[regex]$rx = "OS:\s(?<OS>\S+)\sSP:\s(?<SP>\d\.\d)\sStack:\s(?<Stack>\d\.\d)"
        $pv = $test.productVersion
        
        [string]$os = ($rx.Matches($pv)).foreach({$_.groups["OS"].value})

        #force these as strings and later treat them as [decimal]
        [string]$sp = ($rx.Matches($pv)).foreach({$_.groups["SP"].value})
        [string]$stack = ($rx.Matches($pv)).foreach({$_.groups["Stack"].value})
       
        #write custom result to the pipeline
        $test | Select-Object -property @{Name="Computername";Expression={$Computer.ToUpper()}},
        wsmid,protocolVersion,ProductVendor,
        @{Name="OS";Expression={$OS}},@{Name="SP";Expression={$SP -as [decimal]}},
        @{Name="Stack";Expression={$Stack -as [decimal]}}

        #>