#requires -Version 4
Function Convert-PKCollectionToString {
<#
.SYNOPSIS
    Converts object properties containing collections (arrays) into flattened strings
    for safe and clean export to CSV, avoiding the [system]System.Object[] value

.DESCRIPTION
    This function processes pipeline input objects. For any property that is an
    array or collection (e.g., IPAddress, DNSServerSearchOrder), it joins the
    elements into a single, user-defined string format (Comma-Separated or Stacked)
    before exporting. This eliminates the 'System.Object[]' display in CSV files.

    Meant to be used in a pipeline, in between your input PSObject and Export-CSV.

.NOTES
    Name    : Function_Convert-PKCollectionToString.ps1
    Created : 2027-10-25
    Author  : Paula Kingsley
    Version : 01.01.0000
    History :
    
        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v01.00.0000 - 2025-10-27 - Created script based mostly on Boe Prox's Convert-OutputForCSV

.PARAMETER InputObject
    The object from the pipeline to process
    
.PARAMETER Delimiter
    The character used to separate elements in the output string
        - Comma (Default): Elements are joined by ', '
        - NewLine: Elements are joined by a newline character (`n`) for a 'stacked' appearance

.EXAMPLE
    # Example using CIM (modern WMI replacement)
    PS C:\> Get-CIMInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" |
        Select-Object PSComputerName, IPAddress, DNSServerSearchOrder |
            Convert-PKCollectionToString -Delimiter NewLine |
                Export-Csv -NoTypeInformation -Path "C:\Reports\NIC_Report.csv"
    
#>

[CmdletBinding()]
Param (
    [Parameter(
        Mandatory,
        ValueFromPipeline,
        HelpMessage = "The collection PSObject from the pipeline to process"
    )]
    [psobject]$InputObject,

    [Parameter(HelpMessage = "Choose how to delimit collection elements in the output string (comma or newline; default is newline)")]
    [ValidateSet('Comma', 'NewLine')]
    [string]$Delimiter = 'Comma'
)

Begin {
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"
    
    # Show our settings
    $ScriptName = $MyInvocation.MyCommand.Name
    #$Source = $PSCmdlet.ParameterSetName
    $CurrentParams = $PSBoundParameters

    $MyInvocation.MyCommand.Parameters.keys | Where-Object { $CurrentParams.keys -notContains $_ } | 
    Where-Object { Test-Path variable:$_ } | ForEach-Object {
        $CurrentParams.Add($_, (Get-Variable $_).value)
    }
    #$CurrentParams.Add("ParameterSetName", $Source)
    $CurrentParams.Add("ScriptName", $ScriptName)
    $CurrentParams.Add("ScriptVersion", $Version)
    Write-Verbose "PSBoundParameters: `n`t$(($CurrentParams.GetEnumerator() | Sort-Object) | Format-Table -AutoSize | out-string )"
    
    # Determine the join string based on the user's choice
    $JoinString = switch ($Delimiter) {
        'Comma' { ", " }
        'NewLine' { "`n" }
        default { ", " }
    }

    Write-Verbose "[BEGIN: $ScriptName] Converting input PSObject collection properties to strings for CSV output (delimiter set to $Delimiter)"
}

Process {
    # Check if the object is null (shouldn't happen but good for safety)
    If (-not $InputObject) { return }
    # This block only runs once per InputObject if the properties haven't been cached
    # In this refactor, we don't need to pre-cache properties, we process them dynamically.
    
    # 1. Get the list of properties to maintain order
    $PropertyNames = $InputObject.PSObject.Properties.Name
    
    # 2. Create an ordered hashtable to build the new object's properties
    $NewObjectProperties = [ordered]@{}

    Foreach ($PropName in $PropertyNames) {
        # Get the current value for the property
        $Value = $InputObject.$PropName
        
        # Check If the value is a collection (array, list, etc.) but not a simple string (which is a collection of chars, so we exclude it)
        If ($Value -is [System.Collections.ICollection] -and -not ($Value -is [string])) {
            
            # If it's a collection, join the elements into the desired string format
            $FlattenedValue = $Value -join $JoinString
            
            # Add the flattened value to the new object definition
            $NewObjectProperties.$PropName = $FlattenedValue
            
            Write-Verbose "Flattened property '$PropName' using '$Delimiter'."
        } 
        Else {
            # If it's a single value (string, int, date, etc.), ignore 
            $NewObjectProperties.$PropName = $Value
        }
    }
    
    # 3. Create the new PSCustomObject from the hashtable and output it to the pipeline
    [PSCustomObject]$NewObjectProperties
}

End {
    Write-Verbose "[END: $ScriptName] Finished processing"
}
} # end Convert-PKCollectionToString
