#requires -Version 3
Function Convert-PKSubnetMask {
<#
.SYNOPSIS
    Converts between dotted-decimal subnet mask/CIDR prefix length

.DESCRIPTION
    Converts between dotted-decimal subnet mask/CIDR prefix length
    Accepts pipeline input
    Outputs a PSObject

.NOTES
    Name    : Function_Convert-PKSubnetMask.ps1
    Created : 2017-12-10
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :
        
        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK

        v01.00.0000 - 2017-12-09 - Created script based on Ben Schulz's original (see link)

.LINK
    https://gallery.technet.microsoft.com/scriptcenter/Convert-SubnetMask-80401493

.PARAMETER InputObject
    Dotted decimal subnet mask or CIDR prefix length

.PARAMETER SuppressConsoleOutput
    Suppress non-verbose/non-error console output

.EXAMPLE
    PS C:\> Convert-PKSubnetMask 255.255.0.0,24,1.2.3.4,255.255.254.0,foo -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                                       
        ---                   -----                                       
        Verbose               True                                        
        InputObject           {255.255.0.0, 24, 1.2.3.4, 255.255.254.0...}
        ReturnValidOnly       False                                       
        SuppressConsoleOutput False                                       
        PipelineInput         False                                       
        ScriptName            Convert-PKSubnetMask                        
        ScriptVersion         1.0.0                                       

        Action: Convert between CIDR prefix length and dotted decimal subnet mask

        VERBOSE: Dotted-decimal 255.255.0.0
        VERBOSE: CIDR prefix length 24
        Invalid input '1.2.3.4'
        VERBOSE: Dotted-decimal 255.255.254.0
        Invalid input 'foo'


        InputType Input                Output
        --------- -----                ------
        Mask      255.255.0.0              16
        Prefix    24            255.255.255.0
        -         1.2.3.4               Error
        Mask      255.255.254.0            23
        -         foo                   Error


.EXAMPLE
    PS C:\> $Arr | Convert-PKSubnetMask -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value               
        ---                   -----               
        Verbose               True                
        InputObject                               
        SuppressConsoleOutput False
        PipelineInput         True               
        ScriptName            Convert-PKSubnetMask
        ScriptVersion         1.0.0               

        Action: Convert between CIDR prefix length and dotted decimal subnet mask

        VERBOSE: CIDR prefix length 12
        Invalid input '50'
        VERBOSE: CIDR prefix length 24
        VERBOSE: CIDR prefix length 23
        Invalid input 'foo'
        VERBOSE: Dotted-decimal 255.255.0.0

        InputType       Input Output       
        ---------       ----- ------       
        Prefix             12 255.240.0.0  
        -                  50 Error        
        Prefix             24 255.255.255.0
        Prefix             23 255.255.254.0
        -                 foo Error        
        Mask      255.255.0.0 16   

.EXAMPLE
    PS C:\> Convert-PKSubnetMask 255.255.0.0,24,1.2.3.4,255.255.254.0,foo -ReturnValidOnly -SuppressConsoleOutput

        Invalid input '1.2.3.4'
        Invalid input 'foo'

        InputType Input                Output
        --------- -----                ------
        Mask      255.255.0.0              16
        Prefix    24            255.255.255.0
        Mask      255.255.254.0            23

#>

[Cmdletbinding()]
param( 
    [Parameter(
        Position = 0,
        HelpMessage = "Dotted decimal subnet mask or CIDR prefix length to convert",
        ValueFromPipeline = $True,
        Mandatory=$True
    )]
    [ValidateNotNullOrEmpty()]
    [object[]]$InputObject, 

    [Parameter(
        Mandatory = $false,
        HelpMessage = "Don't return objects for invalid entries"
    )]
    [Switch] $ReturnValidOnly,

    [Parameter(
        Mandatory = $false,
        HelpMessage = "Suppress non-verbose console output"
    )]
    [Switch] $SuppressConsoleOutput
)

Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    $PipelineInput = (-not $PSBoundParameters.ContainsKey("InputObject")) -and (-not $InputObject)

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preference
    $ErrorActionPreference = "Stop"
    $ProgressPreference = "Continue"

    # Output object template
    $OutputTemplate = New-Object PSObject -Property ([ordered] @{
        InputType = "Error"
        Input     = "Error"
        Output    = "Error"
    })
    $Results = @()

    # Inner function to determine entry type
    Function WhatIsIt($Entry) {
        $DottedDecimalPattern = "^(((255\.){3}(255|254|252|248|240|224|192|128|0+))|((255\.){2}(255|254|252|248|240|224|192|128|0+)\.0)|((255\.)(255|254|252|248|240|224|192|128|0+)(\.0+){2})|((255|254|252|248|240|224|192|128|0+)(\.0+){3}))$"
        #$DottedDecimalPattern = "^(?:(?:0?0?\d|0?[1-9]\d|1\d\d|2[0-5][0-5]|2[0-4]\d)\.){3}(?:0?0?\d|0?[1-9]\d|1\d\d|2[0-5][0-5]|2[0-4]\d)$"
        If (($Entry -as [int]) -in 1..32) {"PrefixLength"}
        Elseif ($Entry -match $DottedDecimalPattern) {"SubnetMask"}
        Else {"Error"}
    }

    # Splat for write-progress (although this goes very fast)
    $Activity = "Convert between CIDR prefix length and dotted decimal subnet mask"
    $Param_WP = @{}
    $Param_WP = @{
        Activity = $Activity
        CurrentOperation = ""
        PercentComplete = ""
        Status = "Working..."
    }

    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "Action: $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg}

    $Host.UI.WriteLine()
}

Process {

    $Total = $InputObject.Count
    $Current = 0

    Foreach ($Item in $InputObject) {

        $Output = $OutputTemplate.PSObject.Copy()
        $Output.Input = $Item

        $Current ++
        $Param_WP.CurrentOperation = $Item
        $Param_PercentComplete = ($Current/$Total * 100)
        Write-Progress @Param_WP
    
        $Type = WhatIsIt -Entry $Item

        Switch ($Type) {

            SubnetMask {
                
                $Mask = $Item

                $Msg = "Dotted-decimal $Mask"
                Write-Verbose $Msg

                $Output.InputType = "Mask"

                $Address=[System.Net.IPaddress]"0.0.0.0"

                Try {
                    
                   If ($IsValidInput = [System.Net.IPaddress]::TryParse($Mask, [ref]$Address)) {
            
                        $PrefixArray = @()
                        $Prefix = 0
                        $ByteArray = $Item.Split(".")

                        #This loop converts the bytes to bits, add zeroes when necessary
                        For ($byteCount = 0; $byteCount-lt 4; $byteCount++)  {

                            $bitVariable = $ByteArray[$byteCount]
                            $bitVariable = [Convert]::ToString($bitVariable, 2)
            
                            if ($bitVariable.Length -lt 8) {
                                $NumOnes = $bitVariable.Length
                                $NumZeroes = (8 - $bitVariable.Length)

                                for($BitCount=0; $BitCount -lt $NumZeroes; $BitCount++) {
                                    $Temp = $Temp + "0"
                                }
              
                                $bitVariable = $Temp + $bitVariable
                            }
            
                            #This loop counts the bits in the prefix
                            For ($BitCount=0; $BitCount -lt 8; $BitCount++) {
                                If ($bitVariable[$BitCount] -eq "1") {
                                    $Prefix ++ 
                                }
                                $PrefixArray = $PrefixArray + ($bitVariable[$BitCount])
                            }
                        }

                        [switch]$Mark = $False

                        Foreach ($bit in $PrefixArray) {

                            If ($bit -eq "0") {
                                If (-not $Mark.IsPresent){
                                    $Mark = $True
                                }
                            }
            
                            If ($bit -eq "1") {

                                If ($Mark.IsPresent){
                                    $Msg = "Invalid subnet mask $Mask"
                                    $Host.UI.WriteErrorLine("ERROR: $Msg")
                                    If (-not $ReturnValidOnly.IsPresent) {
                                        $Results += $Output
                                    }
                                }    
                            }       
                        }

                        If ($Prefix -gt 0) {
                            $Output.Output = $Prefix
                            $Results += $Output
                        }
        
                    }
                    
                    # seems like this shouldn't be needed?
                    Else {
                        $Msg = "Invalid subnet mask $Mask"
                        $Host.UI.WriteErrorLine("ERROR: $Msg")
                        If (-not $ReturnValidOnly.IsPresent) {
                            $Results += $Output
                        }
                    }
                }
                Catch {
                    $Msg = "Invalid subnet mask $Mask"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg = "$Msg`n$ErrorDetails"}
                    $Host.UI.WriteErrorLine("ERROR: $Msg")
                    If (-not $ReturnValidOnly.IsPresent) {
                        $Results += $Output
                    }
                }

            } #end if subnet mask

            PrefixLength {
                
                $Prefix = $Item 

                $Msg = "CIDR prefix length $Prefix"
                Write-Verbose $Msg

                $Output.InputType = "Prefix"

                $BitArray = ""
                For ($BitCount = 0; $Prefix -ne "0"; $BitCount++) {
                    $BitArray += '1'
                    $Prefix = $Prefix - 1
                }
    
                #Fill in the rest with zeroes
                While ($BitCount -ne 32) {
                    $BitArray += '0'
                    $BitCount ++ 
                }
    
                #Convert the bit array into subnet mask
                $ClassAAddress = $BitArray.SubString(0,8)
                $ClassAAddress = [Convert]::ToUInt32($ClassAAddress, 2)
                $ClassBAddress = $BitArray.SubString(8,8)
                $ClassBAddress = [Convert]::ToUInt32($ClassBAddress, 2)
                $ClassCAddress = $BitArray.SubString(16,8)
                $ClassCAddress = [Convert]::ToUInt32($ClassCAddress, 2)
                $ClassDAddress = $BitArray.SubString(24,8)           
                $ClassDAddress = [Convert]::ToUInt32($ClassDAddress, 2)
 
                $Mask =  "$ClassAAddress.$ClassBAddress.$ClassCAddress.$ClassDAddress"
            
                $Output.Output = $Mask
                $Results += $Output
                
            } #end if prefix

            Error {
                $Msg = "Invalid input '$Item'"
                $Host.UI.WriteErrorLine($Msg)
                $Output.InputType = "-"
                If (-not $ReturnValidOnly.IsPresent) {
                    $Results += $Output
                }
            }
        
        } #end switch

        # This makes the verbose/console output line spacing funny and I care about that more than I should
        #Write-Output $Output
        
        # ...so I'm doing this awful thing
        #$Results += $Output

    } #end foreach 
 
}

End {
    
    Write-Progress -Activity $Activity -Completed

    $Host.UI.WriteLine()
    Write-Output $Results
}

} #end Convert-PKSubnetMask