Function Get-PKDateTimeExamples {
    <# 
.SYNOPSIS
    Returns standard or unix format date/time formatting options with examples and descriptions

.DESCRIPTION
    Returns standard or unix format date/time formatting options with examples and descriptions
    By default, returns a PSObject, but output can also be to console as write-host, table, or list format.

.NOTES        
    Name    : Function_Get-PKDateTimeExamples.ps1
    Created : 2023-09-27
    Author  : Paula Kingsley
    Version : 01.01.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2023-09-27 - Created script because I can never remember these things!
        v01.01.0000 - 2023-09-28 - Added parameter to hide description from output

.LINK
    https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/get-date?view=powershell-7.3

.PARAMETER Uformat
    Return UFormat/Unix types (default is standard)

.PARAMETER HideDescription
    Hide the description, displaying only command and example output

.PARAMETER OutputType
    Return results as console output using table or list view, or Write-Host (if omitted, output is a PSObject)

.EXAMPLE
    PS C:\> Get-PKDateTimeExamples 
    Returns standard date/time formats as a PSObject

.EXAMPLE
    PS C:\> Get-PKDateTimeExamples -Uformat
    Returns unix-style date/time formats as a PSObject

.EXAMPLE
    PS C:\> Get-PKDateTimeExamples -OutputType AsTable
    Returns standard date/time formats to the console in table format

.EXAMPLE
    PS C:\> Get-PKDateTimeExamples -OutputType AsList
    Returns standard date/time formats to the console in list format

.EXAMPLE
    PS C:\> Get-PKDateTimeExamples -HideDescription -OutputType AsList
    Returns standard date/time formats to the console in list format (commands and examples only)

#>
    
    [CmdletBinding()]
    Param(
        [Parameter(
            HelpMessage = "Return UFormat/Unix types (default is standard)"
        )]    
        [switch]$Uformat,
        [Parameter(
            HelpMessage = "Hide the description, displaying only command and example output"
        )]    
        [switch]$HideDescription,
        [Parameter(
            HelpMessage = "Return results as console output using table or list view, or Write-Host (if omitted, output is a PSObject)"
        )]
        [ValidateSet("AsTable", "WriteHost", "AsList")]
        [string]$OutputType
    )

    Begin {
    
        # Current version (please keep up to date from comment block)
        [version]$Version = "01.00.0000"

        # Show our settings
        $ScriptName = $MyInvocation.MyCommand.Name

        $CurrentParams = $PSBoundParameters
        $MyInvocation.MyCommand.Parameters.keys | Where-Object { $CurrentParams.keys -notContains $_ } | 
        Where-Object { Test-Path variable:$_ } | ForEach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
        $CurrentParams.Add("ScriptName", $ScriptName)
        $CurrentParams.Add("ScriptVersion", $Version)
        Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

        $Display = "standard"
        If ($UFormat.IsPresent) { $Display = "unix/Uformat" }
        Switch ($OutputType) {
            AsList { $Activity = "Display $Display date/time formats as console output in list form" }
            AsTable { $Activity = "Display $Display date/time formats as console output in table form" }
            WriteHost { $Activity = "Display simplified $Display date/time formats as Write-Host console output" }
            Default { $Activity = "Display $Display date/time formats as PSObject" }
        }

        # Console output
        Write-Verbose "[BEGIN: $ScriptName] $Activity"
    
    }
    Process {        

        If (-not $Uformat.IsPresent) {
            $CSVInput = "Description;Command;Example
            Displays long date ;`Get-Date -Format D`;Get-Date -Format D
            Displays long date, short time;`Get-Date -Format f`;Get-Date -Format f
            Displays long date, long time;`Get-Date -Format F`;Get-Date -Format F
            Displays date suitable for filenames;`Get-Date -Format FileDate`;Get-Date -Format FileDate
            Displays long general date/time;`Get-Date -Format G`;Get-Date -Format G
            Displays short general date/time;`Get-Date -Format g`;Get-Date -Format g
            Displays month, day;`Get-Date -Format m`;Get-Date -Format m
            Displays roundtrip format (includes milliseconds and DateTimeKind property);`Get-Date -Format o`;Get-Date -Format o 
            Displays long general date/time in UTC;`Get-Date -Format U`;Get-Date -Format U
            Displays short general date/time in UTC;`Get-Date -Format u`;Get-Date -Format u
            Displays sortable format based on ISO 8601;`Get-Date -Format s`;Get-Date -Format s
            Displays long time format;`Get-Date -Format T`;Get-Date -Format T
            Displays short time format;`Get-Date -Format t`;Get-Date -Format t
            Displays month, year;`Get-Date -Format y`;Get-Date -Format y" | ConvertFrom-CSV -Delimiter ";" | Sort-Object Command
        }
    
        Else {
        
            $CSVInput = "Description;Command;Example
            Displays the day of the week name (full);`Get-Date -Uformat %A`;Get-Date -Uformat %A
            Displays the day of the week name (abbreviated);`Get-Date -Uformat %a`;Get-Date -Uformat %a
            Displays the month name (full);`Get-Date -Uformat %B`;Get-Date -Uformat %B
            Displays the month name  (abbreviated);`Get-Date -Uformat %b`;Get-Date -Uformat %b
            Displays the century number;`Get-Date -Uformat %C`;Get-Date -Uformat %C
            Displays the date and time (abbreviated);`Get-Date -Uformat %c`;Get-Date -Uformat %c
            Displays the date in mm/dd/yy format;`Get-Date -Uformat %D`;Get-Date -Uformat %D
            Displays the day of the month (2 digits);`Get-Date -Uformat %d`;Get-Date -Uformat %d
            Displays the day of the month (preceded by a space if only a single digit);`Get-Date -Uformat %e`;Get-Date -Uformat %e
            Displays the date in YYYY-mm-dd (aka %Y-%m-%d ) using the ISO 8601 date format;`Get-Date -Uformat %F`;Get-Date -Uformat %F
            Displays the ISO week date year (year containing Thursday of the week);`Get-Date -Uformat %G`;Get-Date -Uformat %G
            Displays the hour in 24-hour format;`Get-Date -Uformat %H`;Get-Date -Uformat %H
            Displays the hour in 12-hour format;`Get-Date -Uformat %I`;Get-Date -Uformat %I
            Displays the day of the year;`Get-Date -Uformat %j`;Get-Date -Uformat %j
            Displays the minute;`Get-Date -Uformat %M`;Get-Date -Uformat %M
            Displays the number of the month;`Get-Date -Uformat %m`;Get-Date -Uformat %m
            Displays AM or PM;`Get-Date -Uformat %p`;Get-Date -Uformat %p
            Displays the time in 24-hour format (no seconds);`Get-Date -Uformat %R`;Get-Date -Uformat %R
            Displays the time in 12-hour format;`Get-Date -Uformat %r`;Get-Date -Uformat %r
            Displays the seconds;`Get-Date -Uformat %S`;Get-Date -Uformat %S
            Displays the seconds elapsed since January 1 1970 00:00:00 (UTC);`Get-Date -Uformat %s`;Get-Date -Uformat %s
            Displays the time in 24-hour format;`Get-Date -Uformat %T`;Get-Date -Uformat %T
            Displays the numeric day of the week as 1-7 (NB: changed in PowerShell 7.2);`Get-Date -Uformat %u`;Get-Date -Uformat %u
            Displays the week of the year;`Get-Date -Uformat %V`;Get-Date -Uformat %V
            Displays the numeric day of the week as 0-6;`Get-Date -Uformat %w`;Get-Date -Uformat %w
            Displays the week of the year;`Get-Date -Uformat %W`;Get-Date -Uformat %W
            Displays the date (in standard format for locale);`Get-Date -Uformat %x`;Get-Date -Uformat %x
            Displays the year in 4-digit format;`Get-Date -Uformat %Y`;Get-Date -Uformat %Y
            Displays the year in 2-digit format;`Get-Date -Uformat %y`;Get-Date -Uformat %y
            Displays the time zone offset from Universal Displays the time Coordinate (UTC);`Get-Date -Uformat %Z`;Get-Date -Uformat %Z"  | 
                ConvertFrom-CSV -Delimiter ";"  | Sort-Object Command
        }

        $Results = @()
        $Results = Foreach ($Line in $CSVInput) {
        
            [PSCustomObject]@{
                Command     = $Line.Command
                Description = $Line.Description
                Example     = $(Invoke-Expression $Line.Command)
            }
        }
        
        If ($HideDescription.IsPresent) {
            $Results = $Results | Select-Object Command,Example
        }
        Switch ($OutputType) {
            AsTable {
                Write-Output $Results | Format-Table -AutoSize
            }
            AsList {
                Write-Output $Results | Format-List 
            }
            WriteHost {
                Foreach ($item in $Results) {
                    Write-Host "$($item.Command): " -ForegroundColor Yellow -NoNewline
                    $Item.Example
                }
            }
            Default {
                Write-Output $Results
            }
        }
    }
    End {
        Write-Verbose "[END: $ScriptName]"
    }
}


