#requires -Version 3
Function Convert-PKSubnetCIDR {
<#
.SYNOPSIS
    Converts between dotted-decimal subnet mask and CIDR length, or displays full table with usable networks/hosts

.DESCRIPTION
    Converts between dotted-decimal subnet mask and CIDR length, or displays full table with usable networks/hosts
    Accepts pipeline input
    Returns a PSObject

.NOTES
    Name    : Function_Convert-PKSubnetCIDR.ps1 
    Created : 2018-03-29
    Author  : Paula Kingsley
    Version : 05.02.0000
    History :

        ** PLEASE KEEP $VERSION UPDATED IN BEGIN BLOCK
        
        v01.00.0000 - 2018-03-29 - Created script

.EXAMPLE
    PS C:\> Convert-PKSubnetCIDR -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value               
        ---                   -----               
        Verbose               True                
        InputObj                                  
        DisplayLookupTable    True                
        SuppressConsoleOutput False               
        ParameterSetName      All                 
        PipelineInput         True                
        ScriptName            Convert-PKSubnetCIDR
        ScriptVersion         1.0.0               

        Action: Display CIDR lookup table

        CIDRLength Mask            Class   UsableNetworks UsableHosts                
        ---------- ----            -----   -------------- -----------                
        /1         128.0.0.0       Class A 128            2,147,483,392              
        /2         192.0.0.0       Class A 64             1,073,741,696              
        /3         224.0.0.0       Class A 32             536,870,848                
        /4         240.0.0.0       Class A 16             268,435,424                
        /5         248.0.0.0       Class A 8              134,217,712                
        /6         252.0.0.0       Class A 4              67,108,856                 
        /7         254.0.0.0       Class A 2              33,554,428                 
        /8         255.0.0.0       Class A 1              16,777,214                 
        /9         255.128.0.0     Class B 128            8,388,352                  
        /10        255.192.0.0     Class B 64             4,194,176                  
        /11        255.224.0.0     Class B 32             2,097,088                  
        /12        255.240.0.0     Class B 16             1,048,544                  
        /13        255.248.0.0     Class B 8              524,272                    
        /14        255.252.0.0     Class B 4              262,136                    
        /15        255.254.0.0     Class B 2              131,068                    
        /16        255.255.0.0     Class B 1              65,024                     
        /17        255.255.128.0   Class C 128            32,512                     
        /18        255.255.192.0   Class C 64             16,256                     
        /19        255.255.224.0   Class C 32             8,128                      
        /20        255.255.240.0   Class C 16             4,064                      
        /21        255.255.248.0   Class C 8              2,032                      
        /22        255.255.252.0   Class C 4              1,016                      
        /23        255.255.254.0   Class C 2              508                        
        /24        255.255.255.0   Class C 1              254                        
        /25        255.255.255.128 Class C 2 subnets      124                        
        /26        255.255.255.192 Class C 4 subnets      62                         
        /27        255.255.255.224 Class C 8 subnets      30                         
        /28        255.255.255.240 Class C 16 subnets     14                         
        /29        255.255.255.248 Class C 32 subnets     6                          
        /30        255.255.255.252 Class C 64 subnets     2                          
        /31        255.255.255.254 Class C none           (point to point links only)
        /32        255.255.255.255 Class C 1/256          1                          

.EXAMPLE
    PS C:\> "255.255.254.0","foo","7","255.252.0.0" | Convert-PKSubnetCIDR -Verbose

        VERBOSE: PSBoundParameters: 	
        Key                   Value               
        ---                   -----               
        Verbose               True                
        InputObj                                  
        DisplayLookupTable    False               
        SuppressConsoleOutput False               
        ParameterSetName      Lookup              
        PipelineInput         True                
        ScriptName            Convert-PKSubnetCIDR
        ScriptVersion         1.0.0               

        Action: Convert between CIDR length & dotted-decimal subnet mask formats

        VERBOSE: Convert dotted-decimal subnet mask '255.255.254.0' to CIDR length
        23
        ERROR: Invalid input object 'foo'; please enter CIDR length as integer (1-32), or valid dotted-decimal subnet mask
        VERBOSE: Convert dotted-decimal subnet mask 'foo' to CIDR length
        No match found; please re-run Convert-PKSubnetCIDR without parameters (or with -DisplayLookupTable)
        VERBOSE: Convert CIDR length '7' to dotted-decimal subnet mask
        254.0.0.0
        VERBOSE: Convert dotted-decimal subnet mask '255.252.0.0' to CIDR length
        14

.EXAMPLE
    PS C:\> "255.255.254.0",13,24,"255.252.0.0" | Convert-PKSubnetCIDR -SuppressConsoleOutput

        23
        255.252.0.0
        255.255.255.128
        14

#>
[CmdletBinding(DefaultParameterSetName = "All")]
Param(
    [Parameter(
        ParameterSetName = "Lookup",
        Position = 0,
        ValueFromPipeline = $True,
        Mandatory = $True,
        HelpMessage = "CIDR length or dotted-decimal subnet mask to convert"
    )]
    [ValidateNotNullOrEmpty()]
    [object]$InputObj,

    [Parameter(
        ParameterSetName = "All",
        Mandatory = $False,
        HelpMessage = "Convert CIDR length to dotted-decimal subnet mask"
    )]
    [switch]$DisplayLookupTable,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Suppress non-verbose/non-error console output"
    )]
    [Switch]$SuppressConsoleOutput
)
Begin {
    

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    $PipelineInput = (-not $PSBoundParameters.ContainsKey("InputObj")) -and (-not $InputObj)
    $Source = $PSCmdlet.ParameterSetName
    If ($Source -eq "All") {$DisplayLookupTable = $True}

    # Display parameters
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Regex to confirm valid subnet mask
    $SubnetPattern = "^(((255\.){3}(255|254|252|248|240|224|192|128|0+))|((255\.){2}(255|254|252|248|240|224|192|128|0+)\.0)|((255\.)(255|254|252|248|240|224|192|128|0+)(\.0+){2})|((255|254|252|248|240|224|192|128|0+)(\.0+){3}))$"

    #region Create array & lookup tables
 
    $LookupArr = @'
        "CIDR","CIDRLength","Mask","Class","NumNetworks","NumHosts"
        "/1","1","128.0.0.0","Class A","128","2,147,483,392"
        "/2","2","192.0.0.0","Class A","64","1,073,741,696"
        "/3","3","224.0.0.0","Class A","32","536,870,848"
        "/4","4","240.0.0.0","Class A","16","268,435,424"
        "/5","5","248.0.0.0","Class A","8","134,217,712"
        "/6","6","252.0.0.0","Class A","4","67,108,856"
        "/7","7","254.0.0.0","Class A","2","33,554,428"
        "/8","8","255.0.0.0","Class A","1","16,777,214"
        "/9","9","255.128.0.0","Class B","128","8,388,352"
        "/10","10","255.192.0.0","Class B","64","4,194,176"
        "/11","11","255.224.0.0","Class B","32","2,097,088"
        "/12","12","255.240.0.0","Class B","16","1,048,544"
        "/13","13","255.248.0.0","Class B","8","524,272"
        "/14","14","255.252.0.0","Class B","4","262,136"
        "/15","15","255.254.0.0","Class B","2","131,068"
        "/16","16","255.255.0.0","Class B","1","65,024"
        "/17","17","255.255.128.0","Class C","128","32,512"
        "/18","18","255.255.192.0","Class C","64","16,256"
        "/19","19","255.255.224.0","Class C","32","8,128"
        "/20","20","255.255.240.0","Class C","16","4,064"
        "/21","21","255.255.248.0","Class C","8","2,032"
        "/22","22","255.255.252.0","Class C","4","1,016"
        "/23","23","255.255.254.0","Class C","2","508"
        "/24","24","255.255.255.0","Class C","1","254"
        "/25","25","255.255.255.128","Class C","2 subnets","124"
        "/26","26","255.255.255.192","Class C","4 subnets","62"
        "/27","27","255.255.255.224","Class C","8 subnets","30"
        "/28","28","255.255.255.240","Class C","16 subnets","14"
        "/29","29","255.255.255.248","Class C","32 subnets","6"
        "/30","30","255.255.255.252","Class C","64 subnets","2"
        "/31","31","255.255.255.254","Class C","none","(point to point links only)"
        "/32","32","255.255.255.255","Class C","1/256","1"
'@ | ConvertFrom-CSV


    #[pscustomobject]$LookupArr = ConvertFrom-CSV $LookupStr # -Header "CIDR","CIDRLength","Mask","Class","NumNetworks","NumHosts"
    #$LookupArr | Foreach-Object { $_.CIDRLength = [string]$_.CIDRLength}

    # To look up CIDR length by subnet mask
    $MaskHash = New-Object ([ordered] @{})
    foreach($r in ($LookupArr | Sort @{E={[int]$_.CIDRLength}})){
        $MaskHash.Add($r.Mask,$r.CIDRLength)
    }

    # To look up subnet mask by CIDR length
    $CIDRHash = New-Object ([ordered] @{})
    foreach($r in ($LookupArr | Sort @{E={[int]$_.CIDRLength}})){
        $CIDRHash.Add($r.CIDRLength,$r.Mask)
    }

    #endregion Create array & lookup tables

    $Activity = "Convert between CIDR length & dotted-decimal subnet mask formats"
    If ($Source -eq "All") {$Activity = "Display CIDR lookup table"}
    
    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "Action: $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg}
    $Host.UI.WriteLine()

}
Process {
    
    If ($Source -eq "All") {$Type = "All"}
    Else {
        If ($InputObj -match $SubnetPattern) {$Type = "Mask"}
        Elseif (($InputObj -as [int]) -in 1..32) {$Type = "CIDR"}
        Else {
            $Msg = "Invalid input object '$InputObj'; please enter a CIDR length as integer (1-32), or a valid dotted-decimal subnet mask`nRun $($MyInvocation.MyCommand.Name) without parameters for the full table"
            $Host.UI.WriteErrorLine("ERROR: $Msg")
        }
    }

    Switch ($Type) {
        CIDR {
            $Msg = "Convert CIDR length '$InputObj' to dotted-decimal subnet mask"
            Write-Verbose $Msg
            If ($Results = $CIDRHash["$InputObj"]) {
                Write-Output $Results
            }
            Else {
                $Msg = "No match found; please re-run $($MyInvocation.MyCommand.Name) without parameters (or with -DisplayLookupTable)"
                $Host.UI.WriteErrorLine($Msg)
            }
        }
        Mask {
            $Msg = "Convert dotted-decimal subnet mask '$InputObj' to CIDR length"
            Write-Verbose $Msg

            If ($Results = $MaskHash[$InputObj]) {
                Write-Output $Results
            }
            Else {
                $Msg = "No match found; please re-run $($MyInvocation.MyCommand.Name) without parameters (or with -DisplayLookupTable)"
                $Host.UI.WriteErrorLine($Msg)
            }
        }
        All  {
            Write-Output $LookupArr | Sort-Object @{E={[int]$_.CIDRLength}} | Select @{N="CIDRLength";E={$_.CIDR}},Mask,Class,@{N="UsableNetworks";E={$_.NumNetworks}},@{N="UsableHosts";E={$_.NumHosts}} | Format-Table -AutoSize -Wrap
        }
    }
}
}#end Convert-PKSubnetCIDR


