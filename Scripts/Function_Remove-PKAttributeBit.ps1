#requires -Version 3
Function Remove-PKAttributeBit {
<#
.SYNOPSIS
    Removes one or more filesystem attribute bits from one or more files or folders (recursive)

.DESCRIPTION
    Removes one or more filesystem attribute bits from one or more files or folders (recursive)
    Accepts pipeline input
    Returns a PSobject

.NOTES        
    Name    : Function_Remove-PKAttributeBit.ps1
    Created : 2019-12-31
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2019-12-31 - Created script based on Ed Wilson's original

        # Future plans? 
        ...figure out how to apply multiple attribute changes at once
        ...change function to allow setting *or* removing *or* toggling

.LINK
    https://devblogs.microsoft.com/scripting/use-powershell-to-toggle-the-archive-bit-on-files/

.LINK
    https://devblogs.microsoft.com/scripting/use-a-powershell-cmdlet-to-work-with-file-attributes/

.PARAMETER Path
    Absolute path to file or folder (default is current directory)

.PARAMETER Attributes
    One or more filesystem attributes: Archive, Hidden, Normal, ReadOnly, System

.PARAMETER Quiet
    Hide all non-verbose console output

.EXAMPLE
    PS C:\> Remove-PKAttributeBit -path C:\Temp\Patching -Attribute hidden,readonly,kittens -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value                      
        ---           -----                      
        Path          C:\Temp\Patching           
        Attributes    {hidden, readonly, kittens}
        Verbose       True                       
        Quiet         False                      
        ScriptName    Remove-PKAttributeBit      
        ScriptVersion 1.0.0                      
        PipelineInput False                      

        WARNING: Attribute 'kittens' not in Archive,Hidden,Normal,ReadOnly,System

        BEGIN  : Remove filesystem object attributes 'Hidden', 'ReadOnly'

        VERBOSE: [C:\Temp\Patching\temp] Hidden attribute bit is not set
        VERBOSE: [C:\Temp\Patching\temp] ReadOnly attribute bit is not set
        VERBOSE: [C:\Temp\Patching\All Windows Laptops.csv] Remove Hidden attribute bit
        VERBOSE: [C:\Temp\Patching\All Windows Laptops.csv] Successfully removed attribute bit
        VERBOSE: [C:\Temp\Patching\All Windows Laptops.csv] ReadOnly attribute bit is not set
        VERBOSE: [C:\Temp\Patching\All Windows Servers.csv] Hidden attribute bit is not set
        VERBOSE: [C:\Temp\Patching\All Windows Servers.csv] ReadOnly attribute bit is not set
        VERBOSE: [C:\Temp\Patching\Aug2018_report.xlsx] Hidden attribute bit is not set
        VERBOSE: [C:\Temp\Patching\Aug2018_report.xlsx] Remove ReadOnly attribute bit
        VERBOSE: [C:\Temp\Patching\Aug2018_report.xlsx] Successfully removed attribute bit
        VERBOSE: [C:\Temp\Patching\devdatabases.xml] Remove Hidden attribute bit
        VERBOSE: [C:\Temp\Patching\devdatabases.xml] Successfully removed attribute bit
        VERBOSE: [C:\Temp\Patching\devdatabases.xml] ReadOnly attribute bit is not set
        VERBOSE: [C:\Temp\Patching\temp\contact list.txt] Hidden attribute bit is not set
        VERBOSE: [C:\Temp\Patching\temp\contact list.txt] Remove ReadOnly attribute bit
        VERBOSE: [C:\Temp\Patching\temp\contact list.txt] Operation cancelled by user


        ComputerName  : WORKSTATION14
        Item          : C:\Temp\Patching\temp
        AttributeName : Hidden
        ExistingValue : -
        NewValue      : -
        Messages      : Hidden attribute bit is not set

        ComputerName  : WORKSTATION14
        Item          : C:\Temp\Patching\temp
        AttributeName : ReadOnly
        ExistingValue : -
        NewValue      : -
        Messages      : ReadOnly attribute bit is not set

        ComputerName  : WORKSTATION14
        Item          : C:\Temp\Patching\All Windows Laptops.csv
        AttributeName : Hidden
        ExistingValue : Hidden
        NewValue      : 0
        Messages      : Successfully removed attribute bit

        ComputerName  : WORKSTATION14
        Item          : C:\Temp\Patching\All Windows Laptops.csv
        AttributeName : ReadOnly
        ExistingValue : -
        NewValue      : -
        Messages      : ReadOnly attribute bit is not set

        ComputerName  : WORKSTATION14
        Item          : C:\Temp\Patching\All Windows Servers.csv
        AttributeName : Hidden
        ExistingValue : -
        NewValue      : -
        Messages      : Hidden attribute bit is not set

        ComputerName  : WORKSTATION14
        Item          : C:\Temp\Patching\All Windows Servers.csv
        AttributeName : ReadOnly
        ExistingValue : -
        NewValue      : -
        Messages      : ReadOnly attribute bit is not set

        ComputerName  : WORKSTATION14
        Item          : C:\Temp\Patching\Aug2018_report.xlsx
        AttributeName : Hidden
        ExistingValue : -
        NewValue      : -
        Messages      : Hidden attribute bit is not set

        ComputerName  : WORKSTATION14
        Item          : C:\Temp\Patching\Aug2018_report.xlsx
        AttributeName : ReadOnly
        ExistingValue : ReadOnly
        NewValue      : 0
        Messages      : Successfully removed attribute bit

        ComputerName  : WORKSTATION14
        Item          : C:\Temp\Patching\devdatabases.xml
        AttributeName : Hidden
        ExistingValue : Hidden
        NewValue      : 0
        Messages      : Successfully removed attribute bit

        ComputerName  : WORKSTATION14
        Item          : C:\Temp\Patching\devdatabases.xml
        AttributeName : ReadOnly
        ExistingValue : -
        NewValue      : -
        Messages      : ReadOnly attribute bit is not set

        ComputerName  : WORKSTATION14
        Item          : C:\Temp\Patching\temp\contact list.txt
        AttributeName : Hidden
        ExistingValue : -
        NewValue      : -
        Messages      : Hidden attribute bit is not set

        ComputerName  : WORKSTATION14
        Item          : C:\Temp\Patching\temp\contact list.txt
        AttributeName : ReadOnly
        ExistingValue : ReadOnly
        NewValue      : -
        Messages      : Operation cancelled by user


        END  : Remove filesystem object attributes 'Hidden', 'ReadOnly'



    
#>

[Cmdletbinding(
    SupportsShouldProcess=$True,
    ConfirmImpact="High"
)]    
Param(
    [Parameter(
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "Absolute path to file or folder (default is current directory)"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Name","FullName")]
    $Path = $PWD,

    [Parameter(
        Mandatory = $True,
        HelpMessage = "One or more filesystem attributes: Archive, Hidden, Normal, ReadOnly, System"
    )]
    [ValidateNotNullOrEmpty()]
    $Attributes,

    [Parameter(
        HelpMessage = "Hide all non-verbose console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [Switch] $Quiet

)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    #$Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object {$CurrentParams.keys -notContains $_} |
        Where-Object {Test-Path variable:$_}| ForEach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference    = "Continue"

    #region Attributes
    
    #$AvailableAttributes = [enum]::GetValues([system.io.fileattributes])
    $FileSystemAttributes = "Archive","Hidden","Normal","ReadOnly","System"

    $AttributeArr = @()
    $Attributes | Foreach-Object {
        If ($_ -in $FileSystemAttributes) {
            $AttributeArr += [io.fileattributes]::$_ 
        }
        Else {
            Write-Warning "Attribute '$_' not in $($FileSystemAttributes -join(","))"
        }
    }
    If (-not $AttributeArr) {
        $Msg = "No valid attributes provided; script will exit"
        $Host.UI.WriteErrorLine($Msg)
        Break
    }
    Else {
        #Write-Verbose "Attributes: $($AttributeArr -join(", "))"
    }
    
    #endregion Attributes

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

    # Function to write an error as a string (no stacktrace)
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message)
        $Host.UI.WriteErrorLine("$Message")
    }

    #endregion Functions
    
    #region Splats

    # Splat for Set-ItemProperty
    $Param_Set = @{}
    $Param_Set = @{
        Path        = $Null
        Confirm     = $False
        Name        = "attributes"
        Value       = $Null
        ErrorAction = "Stop"
    }

    #endregion Splats

    # Console output
    $Activity = "Remove filesystem object attributes '$($AttributeArr -join("', '"))'"
    $Msg = "BEGIN  : $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title

}
Process {

    $TotalPaths = $Path.Count
    $CurrentPath = 0
    
    Foreach ($P in $Path) {
        
        If ($PathObj = Get-Item -Path $P) {            
            
            $Results = @()
            $Items = Get-ChildItem -Path $P -Recurse -Force
        
            $CurrentPath ++
            Write-Progress -Activity "Remove attribute bits" -ID 1 -CurrentOperation $Path -PercentComplete ($CurrentPath/$TotalPaths*100)

            $TotalItems = $Items.Count
            $currentItem = 0
        
            Foreach ($Item in $Items){
            
                $CurrentItem ++

                Foreach ($Attr in $AttributeArr) {
                
                    $CurrentValue = (Get-ItemProperty -Path $Item.fullname).attributes -band $Attr
                    $Output = [pscustomobject]@{
                        ComputerName  = $Env:ComputerName
                        Item          = $Item.FullName
                        AttributeName = $Attr
                        ExistingValue = $CurrentValue
                        NewValue      = $Null
                        Messages      = $Null
                    }
                
                    Write-Progress -Activity "Remove $Attr attribute bit" -CurrentOperation $Item.FullName -ID 2 -PercentComplete ($CurrentItem/$TotalItems*100)
                    If ($CurrentValue) {

                        $Msg = "Remove $Attr attribute bit"
                        Write-Verbose "[$($Item.FullName)] $Msg"
                        
                        If ($PSCmdlet.ShouldProcess($Item.FullName,"`n`n`t$Msg`n`n")) {
                            
                            $Param_Set.Path = $Item.FullName
                            $Param_Set.Value = ((Get-ItemProperty $Item.fullname).attributes -BXOR $Attr)
                            Set-ItemProperty @Param_Set 
                            
                            $NewValue = (Get-ItemProperty -Path $Item.fullname).attributes -band $Attr
                            
                            $Msg = "Successfully removed attribute bit"
                            Write-Verbose "[$($Item.FullName)] $Msg"

                            $Output.NewValue = $NewValue
                            $Output.Messages = $Msg

                        } #end if confirmed
                        Else {
                            $Msg = "Operation cancelled by user"
                            Write-Verbose "[$($Item.FullName)] $Msg"

                            $Output.NewValue = "-"
                            $Output.Messages = $Msg

                        } #end else if cancelled

                    } #end if current value
                    Else {
                        $Msg = "Attribute bit is not set"
                        Write-Verbose "[$($Item.FullName)] $Msg"
                        $Output.ExistingValue = $Output.NewValue = "-"
                        $Output.Messages = $Msg
                    }

                    $Results += $Output

                } #end for each attribute

            } #end Foreach item

            Write-Output $Results

        } #end if path
        Else {
            Write-Warning "Invalid path '$P'"
        }
  
    } # end for each path                   
} 
End {
    Write-Progress -Activity * -Complete
    
    $Msg = "END  : $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title
    
} 
} #end Remove-PKAttributeBit




