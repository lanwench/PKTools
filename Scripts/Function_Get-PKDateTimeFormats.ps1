#requires -Version 3
function Get-PKDateTimeFormats {
<#
.SYNOPSIS
    Lists Powershell date/time format examples (standard/custom/all, or a legend)

.DESCRIPTION
    Lists Powershell date/time format examples (standard/custom/all, or a legend)
    Saves valuable minutes spent frantically googling for reminders
    Defaults to examples of all formats
    Returns a PSObject array
    
.NOTES        
    Name    : Function_Get-PKDateTimeFormats.ps1
    Version : 01.00.0000
    Author  : Paula Kingsley
    Created : 2019-04-23

    History:

        ** PLEASE KEEP $VERSION UPDATED IN BEGIN BLOCK **

        v01.00.0000 - 2019-04-23 - Created function

.EXAMPLE
    PS C:\> Get-PKDateTimeFormats -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key              Value                
        ---              -----                
        Verbose          True                 
        Legend           False                
        Type             All                  
        ParameterSetName Example              
        ScriptName       Get-PKDateTimeFormats
        ScriptVersion    1.0.0                

        BEGIN  : Display standard and custom date/time formats
        
        Format                                Example                                           Description                                          
        ------                                -------                                           -----------                                          
        %d/%M/yy                              23-4-19                                                                                                
        d                                     2019-04-23                                        ShortDatePattern                                     
        D                                     Tuesday, April 23, 2019                           LongDatePattern                                      
        d/M/y                                 23-4-19                                                                                                
        dd/MM/yyyy                            23-04-02019                                                                                            
        dd/MM/yyyy %g                         23-04-02019 A.D.                                                                                       
        dd/MM/yyyy gg                         23-04-02019 A.D.                                                                                       
        dddd dd MMM yyyy HH:mm:s.ffff gg      Tuesday 23 Apr 02019 15:56:28.6252 A.D.                                                                
        dddd dd MMMM yyyy HH:mm:s.ffff gg     Tuesday 23 April 02019 15:56:28.6262 A.D.                                                              
        dddd dd MMMM yyyy HH:mm:s.ffff zz gg  Tuesday 23 April 02019 15:56:28.6262 -07 A.D.                                                          
        dddd dd MMMM yyyy HH:mm:s.ffff zzz gg Tuesday 23 April 02019 15:56:28.6272 -07:00 A.D.                                                       
        dddd dd/MM/yyyy %h:m:s tt gg          Tuesday 23-04-02019 3:56:28 PM A.D.                                                                    
        dddd dd/MM/yyyy gg                    Tuesday 23-04-02019 A.D.                                                                               
        dddd dd/MM/yyyy HH:mm:s gg            Tuesday 23-04-02019 15:56:28 A.D.                                                                      
        dddd dd/MM/yyyy hh:mm:s tt gg         Tuesday 23-04-02019 03:56:28 PM A.D.                                                                   
        dddd dd/MM/yyyy HH:mm:s.ffff gg       Tuesday 23-04-02019 15:56:28.6242 A.D.                                                                 
        f                                     Tuesday, April 23, 2019 03:56 PM                  Full date and time (long date and short time)        
        F                                     Tuesday, April 23, 2019 03:56:28 PM               FullDateTimePattern (long date and long time)        
        G                                     2019-04-23 03:56:28 PM                            General (short date and long time)                   
        g                                     2019-04-23 03:56 PM                               General (short date and short time)                  
        m                                     April 23                                          MonthDayPattern                                      
        M                                     April 23                                          MonthDayPattern                                      
        O                                     2019-04-23T15:56:28.5702126-07:00                 Round-trip date/time pattern always uses the invar...
        o                                     2019-04-23T15:56:28.5692145-07:00                 Round-trip date/time pattern always uses the invar...
        R                                     Tue, 23 Apr 2019 15:56:28 GMT                     RFC1123Pattern always uses the invariant culture     
        r                                     Tue, 23 Apr 2019 15:56:28 GMT                     RFC1123Pattern always uses the invariant culture     
        s                                     2019-04-23T15:56:28                               SortableDateTimePattern always uses the invariant ...
        T                                     03:56:28 PM                                       LongTimePattern                                      
        t                                     03:56 PM                                          ShortTimePattern                                     
        u                                     2019-04-23 15:56:28Z                              UniversalSortableDateTimePattern                     
        U                                     2019-04-23 15:56:28Z                              Full date and time - universal time                  
        Y                                     April 2019                                        YearMonthPattern                                     
        y                                     April 2019                                        YearMonthPattern                                     

        END    : Display standard and custom date/time formats

.EXAMPLE
    PS C:\> Get-PKDateTimeFormats -Legend

        BEGIN  : Display date/time format legend

        Letter code Meaning                                             Example                  
        ----------- -------                                             -------                  
        c           Date and time                                       Fri Jun 16 10:31:27 2015 
        D           Date in mm/dd/yy format                             06/14/06                 
        x           Date in standard format for locale                  09/12/15 for English-US  
        C           Century                                             20 for 2015              
        Y, G        Year in 4-digit format                              2015                     
        y, g        Year in 2-digit format                              15                       
        b, h        Month name – abbreviated                            Jan                      
        B           Month name – full                                   January                  
        m           Month number                                        6                        
        W, U        Week of the year – zero based                       00-52                    
        V           Week of the year – one based                        01-53                    
        a           Day of the week – abbreviated name                  Mon                      
        A           Day of the week – full name                         Monday                   
        u, w        Day of the week – number                            Monday = 1               
        d           Day of the month – 2 digits                         5                        
        e           Day of the month – digit preceded by a space        5                        
        j           Day of the year                                     1-366                    
        p           AM or PM                                            PM                       
        r           Time in 12-hour format                              9:15:36 AM               
        R           Time in 24-hour format – no seconds                 17:45                    
        T, X        Time in 24 hour format                              17:45:52                 
        Z           Time zone offset from Universal Time Coordinate UTC 7                        
        H, k        Hour in 24-hour format                              17                       
        I, l        Hour in 12 hour format                              5                        
        M           Minutes                                             35                       
        S           Seconds                                             5                        
        s           Seconds elapsed since January 1, 1970               00:00:00 1150451174.95705
        n           newline character                                   n                        
        t           Tab character                                       t                        

        END    : Display date/time format legend


.EXAMPLE
    PS C:\> Get-PKDateTimeFormats -Type Custom -Quiet

        Format                                Example                                         
        ------                                -------                                         
        %d/%M/yy                              23-4-19                                         
        d/M/y                                 23-4-19                                         
        dd/MM/yyyy                            23-04-02019                                     
        dd/MM/yyyy %g                         23-04-02019 A.D.                                
        dd/MM/yyyy gg                         23-04-02019 A.D.                                
        dddd dd MMM yyyy HH:mm:s.ffff gg      Tuesday 23 Apr 02019 16:03:9.3717 A.D.          
        dddd dd MMMM yyyy HH:mm:s.ffff gg     Tuesday 23 April 02019 16:03:9.3727 A.D.        
        dddd dd MMMM yyyy HH:mm:s.ffff zz gg  Tuesday 23 April 02019 16:03:9.3737 -07 A.D.    
        dddd dd MMMM yyyy HH:mm:s.ffff zzz gg Tuesday 23 April 02019 16:03:9.3747 -07:00 A.D. 
        dddd dd/MM/yyyy %h:m:s tt gg          Tuesday 23-04-02019 4:3:9 PM A.D.               
        dddd dd/MM/yyyy gg                    Tuesday 23-04-02019 A.D.                        
        dddd dd/MM/yyyy HH:mm:s gg            Tuesday 23-04-02019 16:03:9 A.D.                
        dddd dd/MM/yyyy hh:mm:s tt gg         Tuesday 23-04-02019 04:03:9 PM A.D.             
        dddd dd/MM/yyyy HH:mm:s.ffff gg       Tuesday 23-04-02019 16:03:9.3707 A.D.    

#>

[CmdletBinding(
    DefaultParameterSetName = "Example"
)]
Param(
    
    [Parameter(
        ParameterSetName = "Legend",
        HelpMessage = "Return legend of letter codes/meanings"
    )]
    [switch]$Legend,

    [Parameter(
        ParameterSetName = "Example",
        HelpMessage="Return formatting examples: Standard, Custom or All (default is All)"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Standard","Custom","All")]
    [string]$Type = "All",

    [Parameter(
        HelpMessage = "Suppress all non-verbose/non-error console output"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("SuppressConsoleOutput")]
    [switch]$Quiet

)

Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"
    
    # Console output
    Switch ($Source) {
        Example {
            Switch ($Type) {
                Standard {$Activity = "Display standard date/time formats (usage: 'Get-Date -Format d')"}
                Custom   {$Activity = "Display custom date/time formats (usage: 'Get-Date -Format yyyy-M-d')"}
                All      {$Activity = "Display standard and custom date/time formats"}
            }
        }
        Legend {$Activity = "Display date/time format legend"}        
    }

    $Msg = "BEGIN  : $Activity"
    $BGColor = $host.UI.RawUI.BackgroundColor
    $FGColor = "Yellow"
    If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg")}
    Else {Write-Verbose $Msg}
    
}
Process {
    
    Switch ($Source) {
        Example {
            
            # Standard
            $StandardArr = @()
            $StandardHT = @(
                @{Format='d';Description='ShortDatePattern';Example=$("{0} " -f (get-date -Format d ))}
                @{Format='D';Description='LongDatePattern';Example=$("{0} " -f (get-date -Format D ))}
                @{Format='f';Description='Full date and time (long date and short time)';Example = $("{0} " -f (get-date -Format f ))}
                @{Format='F';Description='FullDateTimePattern (long date and long time)';Example = $("{0} " -f (get-date -Format F ))}
                @{Format='g';Description='General (short date and short time)';Example = $("{0} " -f (get-date -Format g ))}
                @{Format='G';Description='General (short date and long time)';Example = $("{0} " -f (get-date -Format G ))}
                @{Format='m';Description='MonthDayPattern';Example = $("{0} " -f (get-date -Format m ))}
                @{Format='M';Description='MonthDayPattern';Example = $("{0} " -f (get-date -Format M ))}
                @{Format='o';Description='Round-trip date/time pattern always uses the invariant culture';Example = $("{0} " -f (get-date -Format o ))}
                @{Format='O';Description='Round-trip date/time pattern always uses the invariant culture';Example = $("{0} " -f (get-date -Format O ))}
                @{Format='r';Description='RFC1123Pattern always uses the invariant culture';Example = $("{0} " -f (get-date -Format r ))}
                @{Format='R';Description='RFC1123Pattern always uses the invariant culture';Example = $("{0} " -f (get-date -Format R ))}
                @{Format='s';Description='SortableDateTimePattern always uses the invariant culture';Example = $("{0} " -f (get-date -Format s ))}
                @{Format='t';Description='ShortTimePattern';Example = $("{0} " -f (get-date -Format t ))}
                @{Format='T';Description='LongTimePattern';Example = $("{0} " -f (get-date -Format T ))}
                @{Format='u';Description='UniversalSortableDateTimePattern';Example = $("{0} " -f (get-date -Format u ))}
                @{Format='U';Description='Full date and time - universal time';Example = $("{0} " -f (get-date -Format u ))}
                @{Format='y';Description='YearMonthPattern';Example = $("{0} " -f (get-date -Format y ))}
                @{Format='Y';Description='YearMonthPattern';Example = $("{0} " -f (get-date -Format Y ))}
            )
            $StandardHT | Foreach-Object {$StandardArr += New-Object PSObject -Property $_} | Select Format,Example | Sort Format,Example
    
            # Custom
            $CustomArr = @()
            $CustomHT = @(
                @{Format='d/M/y';Example=$("{0} " -f (get-date -Format d/M/y ))}
                @{Format='%d/%M/yy';Example=$("{0} " -f (get-date -Format d/M/yy ))}
                @{Format='dd/MM/yyyy';Example=$("{0} " -f (get-date -Format dd/MM/yyyyy ))}
                @{Format='dd/MM/yyyy %g';Example=$("{0} " -f (get-date -Format 'dd/MM/yyyyy %g'))}
                @{Format='dd/MM/yyyy gg';Example=$("{0} " -f (get-date -Format 'dd/MM/yyyyy gg'))}
                @{Format='dddd dd/MM/yyyy gg';Example=$("{0} " -f (get-date -Format 'dddd dd/MM/yyyyy gg'))}
                @{Format='dddd dd/MM/yyyy %h:m:s tt gg';Example=$("{0} " -f (get-date -Format 'dddd dd/MM/yyyyy %h:m:s tt gg'))}
                @{Format='dddd dd/MM/yyyy hh:mm:s tt gg';Example=$("{0} " -f (get-date -Format 'dddd dd/MM/yyyyy hh:mm:s tt gg'))}
                @{Format='dddd dd/MM/yyyy HH:mm:s gg';Example=$("{0} " -f (get-date -Format 'dddd dd/MM/yyyyy HH:mm:s gg'))}
                @{Format='dddd dd/MM/yyyy HH:mm:s.ffff gg';Example=$("{0} " -f (get-date -Format 'dddd dd/MM/yyyyy HH:mm:s.ffff gg'))}
                @{Format='dddd dd MMM yyyy HH:mm:s.ffff gg';Example=$("{0} " -f (get-date -Format 'dddd dd MMM yyyyy HH:mm:s.ffff gg'))}
                @{Format='dddd dd MMMM yyyy HH:mm:s.ffff gg';Example=$("{0} " -f (get-date -Format 'dddd dd MMMM yyyyy HH:mm:s.ffff gg'))}
                @{Format='dddd dd MMMM yyyy HH:mm:s.ffff zz gg';Example=$("{0} " -f (get-date -Format 'dddd dd MMMM yyyyy HH:mm:s.ffff zz gg'))}
                @{Format='dddd dd MMMM yyyy HH:mm:s.ffff zzz gg';Example=$("{0} " -f (get-date -Format 'dddd dd MMMM yyyyy HH:mm:s.ffff zzz gg'))} 
            )
            $CustomHT | Foreach-Object {$CustomArr += New-Object PSObject -Property $_} | Select Format,Example | Sort Format,Example
            
            Switch ($Type) {
                Standard {$StandardArr | Select Format,Example,Description | Sort Format}
                Custom   {$CustomArr | Select Format,Example | Sort Format}
                All      {$StandardArr + $CustomArr | Select Format,Example,Description | Sort Format}
            }
        }
        Legend   {
            '"Letter code";"Meaning";"Example"
            "c";"Date and time";"Fri Jun 16 10:31:27 2015"
            "D";"Date in mm/dd/yy format";"06/14/06"
            "x";"Date in standard format for locale";"09/12/15 for English-US"
            "C";"Century";"20 for 2015"
            "Y, G";"Year in 4-digit format";"2015"
            "y, g";"Year in 2-digit format";"15"
            "b, h";"Month name – abbreviated";"Jan"
            "B";"Month name – full";"January"
            "m";"Month number";"6"
            "W, U";"Week of the year – zero based";"00-52"
            "V";"Week of the year – one based";"01-53"
            "a";"Day of the week – abbreviated name";"Mon"
            "A";"Day of the week – full name";"Monday"
            "u, w";"Day of the week – number";"Monday = 1"
            "d";"Day of the month – 2 digits";"5"
            "e";"Day of the month – digit preceded by a space";"5"
            "j";"Day of the year";"1-366"
            "p";"AM or PM";"PM"
            "r";"Time in 12-hour format";"9:15:36 AM"
            "R";"Time in 24-hour format – no seconds";"17:45"
            "T, X";"Time in 24 hour format";"17:45:52"
            "Z";"Time zone offset from Universal Time Coordinate UTC";"7"
            "H, k";"Hour in 24-hour format";"17"
            "I, l";"Hour in 12 hour format";"5"
            "M";"Minutes";"35"
            "S";"Seconds";"5"
            "s";"Seconds elapsed since January 1, 1970";"00:00:00 1150451174.95705"
            "n";"newline character";"n"
            "t";"Tab character";"t"' | ConvertFrom-CSV -Delimiter ";"
        }
    }

}
End {

    $Msg = "END    : $Activity"
    $FGColor = "Yellow"
    If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg")}
    Else {Write-Verbose $Msg}
}
} #end Get-PKDateTimeFormats