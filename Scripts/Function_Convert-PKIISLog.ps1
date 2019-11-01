#requires -Version 3
Function Convert-PKIISLog {
<#
.SYNOPSIS
    Parses an IIS log from a file (string or object) and returns a PSObject

.DESCRIPTION
    Parses an IIS log from a file (string or object) and returns a PSObject
    Accepts pipeline input
            
.NOTES
    Name    : Function_Convert-PKIISLog.ps1
    Author  : Paula Kingsley
    Created : 2019-11-01
    Version : 01.00.0000
    History:

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK

        v01.00.0000 - 2019-11-01 - Created script based on Nathan Hartley's code (see link)

.LINK
    https://social.technet.microsoft.com/Forums/scriptcenter/en-US/46bc6859-d9e3-47c3-b1a6-5132281df18b/howto-use-powershell-to-parse-iis-logs-files

.EXAMPLE
    PS C:\> $LogResults = Convert-PKIISLog -Logfile $File -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value                                                                                                      
        ---           -----                                                                                                      
        Logfile       {G:\Logfiles\20191101_MegacorpSMTP_ex191101.log}
        Verbose       True                                                                                                       
        Quiet         False                                                                                                      
        PipelineInput False                                                                                                      
        ScriptName    Convert-PKIISLog                                                                                           
        ScriptVersion 1.0.0                                                                                                      
        BEGIN: Convert IIS log file to PSObject

        [20191101_MegacorpSMTP_ex191101.log] Verify file path
        [20191101_MegacorpSMTP_ex191101.log] Verified file path
        [20191101_MegacorpSMTP_ex191101.log] Get logfile content
        [20191101_MegacorpSMTP_ex191101.log] Successfuly got logfile content
        [20191101_MegacorpSMTP_ex191101.log] Parse logfile content and output PSObject
        [20191101_MegacorpSMTP_ex191101.log] Successfully converted 11494 logfile entries to a PSObject

        END  : Convert IIS log file to PSObject
        
        PS C:\> $LogResults | Select -First 3

        date            : 2019-11-01
        time            : 00:01:01
        c-ip            : 10.32.12.4
        cs-username     : internaltesting@megacorp.com
        s-sitename      : SMTPSVC1
        s-computername  : WEBSERVER
        s-ip            : 10.32.8.61
        s-port          : 0
        cs-method       : EHLO
        cs-uri-stem     : -
        cs-uri-query    : +internaltesting@megacorp.com
        sc-status       : 250
        sc-win32-status : 0
        sc-bytes        : 196
        cs-bytes        : 39
        time-taken      : 0
        cs-version      : SMTP
        cs-host         : -
        cs(User-Agent)  : -
        cs(Cookie)      : -
        cs(Referer)     : -

        date            : 2019-11-01
        time            : 00:01:01
        c-ip            : 10.32.12.4
        cs-username     : internaltesting@megacorp.com
        s-sitename      : SMTPSVC1
        s-computername  : WEBSERVER
        s-ip            : 10.32.8.61
        s-port          : 0
        cs-method       : MAIL
        cs-uri-stem     : -
        cs-uri-query    : +FROM:<internaltesting@megacorp.com>
        sc-status       : 250
        sc-win32-status : 0
        sc-bytes        : 59
        cs-bytes        : 46
        time-taken      : 0
        cs-version      : SMTP
        cs-host         : -
        cs(User-Agent)  : -
        cs(Cookie)      : -
        cs(Referer)     : -

        date            : 2019-11-01
        time            : 00:01:01
        c-ip            : 10.32.12.4
        cs-username     : internaltesting@megacorp.com
        s-sitename      : SMTPSVC1
        s-computername  : WEBSERVER
        s-ip            : 10.32.8.61
        s-port          : 0
        cs-method       : RCPT
        cs-uri-stem     : -
        cs-uri-query    : +TO:<database-leads@megacorp.com>
        sc-status       : 250
        sc-win32-status : 0
        sc-bytes        : 37
        cs-bytes        : 34
        time-taken      : 0
        cs-version      : SMTP
        cs-host         : -
        cs(User-Agent)  : -
        cs(Cookie)      : -
        cs(Referer)     : -

.EXAMPLE
    PS C:\> Get-Item -Path \\sqldevbox.domain.local\c$\temp\smtplogcopy.log | Convert-PKIISLog -Quiet | 
            Select-Object date,time,s-computername,cs-username,cs-method,sc-status,cs-uri-query | Format-Table -Autosize

        date       time     s-computername cs-username                  cs-method sc-status cs-uri-query                                                                                              
        ----       ----     -------------- -----------                  --------- --------- ------------                                                                                              
        2019-11-01 00:01:01 SQLDEVBOX      sqldevbox.domain.local       EHLO      250       +sqldevbox.domain.local                                                                       
        2019-11-01 00:01:01 SQLDEVBOX      sqldevbox.domain.local       MAIL      250       +FROM:<SQLDEVBOX@domain.local>                                                                
        2019-11-01 00:01:01 SQLDEVBOX      sqldevbox.domain.local       RCPT      250       +TO:<joe.bloggs@company.com>                                                                            
        2019-11-01 00:01:01 SQLDEVBOX      sqldevbox.domain.local       DATA      250       +<01d59047$Blat.v3.2.4$71ce9260$3885101d168@smtp.intra>                                                   
        2019-11-01 00:01:01 SQLDEVBOX      sqldevbox.domain.local       QUIT      240       sqldevbox.domain.local                                                                        
        2019-11-01 00:01:01 SQLDEVBOX      OutboundConnectionResponse   -         0         220+***************************************************
        2019-11-01 00:01:01 SQLDEVBOX      OutboundConnectionCommand    EHLO      0         domain.local                                                                                      
        2019-11-01 00:01:01 SQLDEVBOX      OutboundConnectionResponse   -         0         250-corpexchange.domain.local+Hello+[192.168.56.7]                                                   
        2019-11-01 00:01:01 SQLDEVBOX      OutboundConnectionCommand    MAIL      0         FROM:<SQLDEVBOX@domain.local>                                                                 
        2019-11-01 00:01:01 SQLDEVBOX      OutboundConnectionResponse   -         0         250+2.1.0+Sender+OK                                                                                       
        2019-11-01 00:01:01 SQLDEVBOX      OutboundConnectionCommand    RCPT      0         TO:<joe.bloggs@company.com>                                                                             
        2019-11-01 00:01:03 SQLDEVBOX      OutboundConnectionResponse   -         0         550+5.1.1+User+unknown                                                                                    
        2019-11-01 00:01:03 SQLDEVBOX      OutboundConnectionCommand    RSET      0         -                                                                                                         
        2019-11-01 00:01:04 SQLDEVBOX      OutboundConnectionResponse   -         0         250+2.0.0+Resetting                                                                                       
        2019-11-01 00:01:04 SQLDEVBOX      OutboundConnectionCommand    QUIT      0         -                                                                                                         
        2019-11-01 00:01:04 SQLDEVBOX      OutboundConnectionResponse   -         0         221+2.0.0+Service+closing+transmission+channel                                                            
        2019-11-01 00:01:04 SQLDEVBOX      OutboundConnectionResponse   -         0         220+******************************************************
        2019-11-01 00:01:04 SQLDEVBOX      OutboundConnectionCommand    EHLO      0         domain.local                                                                                      
        2019-11-01 00:01:04 SQLDEVBOX      OutboundConnectionResponse   -         0         250-corpexchange.domain.local+Hello+[192.168.56.7]                                                   
        2019-11-01 00:01:04 SQLDEVBOX      OutboundConnectionCommand    MAIL      0         FROM:<>   

        <snip>


#>
[Cmdletbinding()]
Param(
    [Parameter(
        Position = 0,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "One or more IIS logfile absolute paths or file objects"
    )]
    [Alias("Name","FullName")]
    [ValidateNotNullOrEmpty()]
    [object[]]$Logfile,

    [Parameter(
        HelpMessage = "Suppress all non-verbose console output"
    )]
    [Alias("SuppressconsoleOutput")]
    [switch]$Quiet

)

Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here?
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    $CurrentParams = $PSBoundParameters
    
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path Variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    #$CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference = "Continue"

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
        Param([Parameter(ValueFromPipeline)]$Message,[switch]$Force)
        $Host.UI.WriteErrorLine("$Message")
    }

    #endregion Functions

    #region Splats

    # General-purpose splat
    $StdParams = @{}
    $StdParams = @{
        Verbose     = $False
        ErrorAction = "Stop"
    }

    # Splat for Write-Progress 
    $Activity = "Convert IIS log file to PSObject"
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        CurrentOperation = $Null
        #PercentComplete  = $Null
        Status           = "Working"
    }

    #endregion Splats

    # Console output
    "BEGIN: $Activity" | Write-MessageInfo -FGColor Yellow -Title
}
Process {

    Foreach ($File in $Logfile) {
        
        If ($File -is [string]) {}
        Elseif ($File -is [System.IO.FileInfo]){$File = $File.FullName}

        $Msg = "Verify file path"

        $Param_WP.Status = $File
        $Param_WP.CurrentOperation = $Msg
        Write-Progress @Param_WP
        
        $FileName = $File | Split-Path -Leaf -ErrorAction Stop
        "[$FileName] $Msg" | Write-MessageInfo -FGColor White

        Try {
            $Null = Test-Path $File -PathType Leaf -ErrorAction Stop
            $Msg = "Verified file path"
            "[$FileName] $Msg" | Write-MessageInfo -FGColor Green

            Try {
                
                $Msg = "Get logfile content"
                "[$FileName] $Msg" | Write-MessageInfo -FGColor White

                $Param_WP.CurrentOperation = $Msg
                Write-Progress @Param_WP

                $FileContent = Get-Content -Path $File -ErrorAction Stop 
                $Msg = "Successfuly got logfile content"
                "[$FileName] $Msg" | Write-MessageInfo -FGColor Green

                If ($FileContent -match "#Software: Microsoft Internet Information Services") {
                    
                    $Total = ($FileContent -as [array]).Count
                    $Current = 0
                    Try {
                    
                        $Msg = "Parse logfile content and output PSObject"
                        "[$FileName] $Msg" | Write-MessageInfo -FGColor White

                        $Param_WP.CurrentOperation = $Msg
                        Write-Progress @Param_WP

                        $Output = $FileContent | Foreach-Object {
                        
                            $Current ++
                            Write-Progress @Param_WP -PercentComplete ($Current/$Total*100)

                            $_ -replace '#Fields: ', ''
                            } | 
                            
                            Where-Object {$_ -notmatch '^#'} | ConvertFrom-CSV -Delimiter ' '
                            
                            $Msg = "Successfully converted $Total logfile entries to a PSObject"
                            "[$FileName] $Msg" | Write-MessageInfo -FGColor Green
                            Write-Progress -Activity $Activity -Completed
                            
                            Write-Output $Output
                
                    }
                    Catch {
                        $Msg = "Failed to parse content for file '$File'"
                        If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                        "[$FileName] $Msg" | Write-MessageError
                    }
                } 
                Else {
                    $Msg = "File '$File' does not appear to be an IIS logfile"
                    "[$FileName] $Msg" | Write-MessageError
                }
            }
            Catch {
                $Msg = "Failed to get content from file '$File'"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                "[$FileName] $Msg" | Write-MessageError
           }
        }
        Catch {
            $Msg = "Failed to verify file path '$File'"
            "[$FileName] $Msg" | Write-MessageError
        }

    } #end foreach 
    
}
End {
    
    Write-Progress -Activity $Activity -Completed
    "END  : $Activity" | Write-MessageInfo -FGColor Yellow -Title

}
} #end Convert-PKIISLog
