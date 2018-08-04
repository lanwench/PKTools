#requires -version 4
Function Convert-PKSubnetMask {
<#
.SYNOPSIS
    Converts a dotted-decimal subnet mask to CIDR index, or vice versa
    
.DESCRIPTION
    Converts a dotted-decimal subnet mask to CIDR index, or vice versa
    Uses string manipulation and regular expressions
    Accepts pipeline input
    Returns a string or integer
     
.NOTES
    Name    : Function_Convert-PKSubnetMask.ps1
    Author  : Paula Kingsley
    Created : 2018-06-26
    Version : 01.00.0000
    History :

        *** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00.0000 - 2018-06-26 - Created script to replace more complicated prior version

.EXAMPLE
    PS C:\> Convert-PKSubnetMask -InputObj 24 -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                          Value
        ---                          -----
        InputObj                        24
        Verbose                       True
        PipelineInput                False
        ScriptName    Convert-PKSubnetMask
        ScriptVersion                1.0.0

        VERBOSE: Convert CIDR index '24' to dotted-decimal subnet mask
        255.255.255.0

.EXAMPLE
    PS C:\>  Convert-PKSubnetMask -InputObj 255.255.0.0 -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value               
        ---           -----               
        InputObj      255.255.0.0         
        Verbose       True                
        PipelineInput False               
        ScriptName    Convert-PKSubnetMask
        ScriptVersion 1.0.0               

        VERBOSE: Convert subnet mask 255.255.0.0 to CIDR index
        16
                                                                       

.EXAMPLE
    PS C:\> "255.255.254.0",23,"kittens","/16" | Convert-PKSubnetMask

        23
        255.255.254.0
        ERROR: Unrecognized input object 'kittens'; please enter a valid CIDR index or dotted-decimal subnet mask
        255.255.0.0


 #>
 [CmdletBinding()]
 Param(
    
    [Parameter(
        Mandatory = $True,
        Position = 0,
        ValueFromPipeline = $True
    )]
    [ValidateNotNullOrEmpty()]
    $InputObj   
 )
 Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Display parameters
    $PipelineInput = (-not $PSBoundParameters.ContainsKey("InputObj")) -and (-not $InputObj)
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    #$CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Regex
    $SubnetPattern = "^(((255\.){3}(255|254|252|248|240|224|192|128|0+))|((255\.){2}(255|254|252|248|240|224|192|128|0+)\.0)|((255\.)(255|254|252|248|240|224|192|128|0+)(\.0+){2})|((255|254|252|248|240|224|192|128|0+)(\.0+){3}))$"
    $CIDRPattern = "(/\d|\d)"

 }
 Process{

    If ($InputObj -match $SubnetPattern) {$Source = "Mask"}
    Elseif ($InputObj -match $CIDRPattern) {$Source = "Index"}
    Else {$Source = "Unknown"}

    Switch ($Source) {
        Mask {
            $SubnetMask = $InputObj
            $Msg = "Convert subnet mask $SubnetMask to CIDR index"
            Write-Verbose $Msg

            $result = 0
            [IPAddress] $IP = $SubnetMask.Trim()
            $Octets = $IP.IPAddressToString.Split('.')
            Foreach ($Octet in $Octets){
                While (0 -ne $Octet) {
                    $Octet = ($Octet -shl 1) -band [byte]::MaxValue
                    $Result++
                }
            }
            Write-Output $Result
        }
        Index {
            $CIDRIndex = $InputObj
            $Msg = "Convert CIDR index '$CIDRIndex' to dotted-decimal subnet mask"
            Write-Verbose $Msg
            
            If ($CIDRIndex -match "/") {$Index = ($($CIDRIndex.Replace("/",$Null)).Trim()) -as [int]}
            Else {[int]$Index = $CIDRIndex}
            
            $Result = (('1'*$Index+'0'*(32-$Index)-split'(.{8})')-ne''| Foreach-Object {[convert]::ToUInt32($_,2)})-join'.'
            Write-Output $Result
        }
        Unknown {
            $Msg = "Unrecognized input object '$InputObj'`nPlease enter a valid CIDR index or dotted-decimal subnet mask"
            $Host.UI.WriteErrorLine("ERROR: $Msg")
        }
    }
}
} #end Convert-PKSubnetMask
 