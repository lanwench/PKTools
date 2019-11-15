#requires -Version 3
Function Convert-PKEXchangeSMTPLog {
<#
.SYNOPSIS
    Parses an Exchange send or receive connector log from a file (string or object) and returns a PSObject        

.DESCRIPTION
    Parses an Exchange send or receive connector log from a file (string or object) and returns a PSObject    
    Accepts pipeline input
            
.NOTES
    Name    : Function_Convert-PKExchangeSMTPLog.ps1
    Author  : Paula Kingsley
    Created : 2019-11-01
    Version : 01.00.0000
    History:

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK

        v01.00.0000 - 2019-11-01 - Created script based on Nathan Hartley's code (see link)

.LINK
    https://social.technet.microsoft.com/Forums/scriptcenter/en-US/46bc6859-d9e3-47c3-b1a6-5132281df18b/howto-use-powershell-to-parse-iis-logs-files

.EXAMPLE
    PS C:\> $LogContents = Convert-PKExchangeSMTPLog -Logfile C:\Temp\2019-11-08_EXCHSRVER_SEND20191108-8.LOG -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value                                                                                                        
        ---           -----                                                                                                        
        Logfile       {C:\Temp\2019-11-08_EXCHSRVER_SEND20191108-8.LOG}
        Verbose       True                                                                                                         
        Quiet         False                                                                                                        
        PipelineInput False                                                                                                        
        ScriptName    Convert-PKEXchangeSMTPLog                                                                                    
        ScriptVersion 1.0.0                                                                                                        

        BEGIN: Convert EXchange Send/Receive connector log file to PSObject

        [2019-11-08_SCEXCH02_SEND20191108-8.LOG] Verify file path
        [2019-11-08_SCEXCH02_SEND20191108-8.LOG] Verified file path
        [2019-11-08_SCEXCH02_SEND20191108-8.LOG] Get logfile content
        [2019-11-08_SCEXCH02_SEND20191108-8.LOG] Successfuly got logfile content
        [2019-11-08_SCEXCH02_SEND20191108-8.LOG] Parse logfile content and output PSObject
        [2019-11-08_SCEXCH02_SEND20191108-8.LOG] Successfully converted 67154 logfile entries to a PSObject

        END  : Convert EXchange Send/Receive connector log file to PSObject

        PS C:\> $MyFile | Select-Object -First 10 | Format-Table -AutoSize

        connector-id   session-id       sequence-number local-endpoint     remote-endpoint   event data                              context               DateTime             
        ------------   ----------       --------------- --------------     ---------------   ----- ----                              -------               --------             
        SMTP-relay-out 08D751B8BF60FE06 21              192.168.30.5:57626 74.125.142.27:25  -                                       Local                 11/8/2019 10:04:39 AM
        SMTP-relay-out 08D751B8BF60FE0A 0                                  76.74.201.51:25   *                                       attempting to connect 11/8/2019 10:04:49 AM
        SMTP-relay-out 08D751B8BF60FE0A 1               192.168.30.5:57642 76.74.201.51:25   +                                                             11/8/2019 10:04:49 AM
        SMTP-relay-out 08D751B8BF60FE0A 2               192.168.30.5:57642 76.74.201.51:25   <     554 srv8.tinyco.com                                     11/8/2019 10:04:49 AM
        SMTP-relay-out 08D751B8BF60FE0B 0                                  192.206.151.51:25 *                                       attempting to connect 11/8/2019 10:04:49 AM
        SMTP-relay-out 08D751B8BF60FE0A 3               192.168.30.5:57642 76.74.201.51:25   >     QUIT                                                    11/8/2019 10:04:49 AM
        SMTP-relay-out 08D751B8BF60FE0A 4               192.168.30.5:57642 76.74.201.51:25   -                                       Remote                11/8/2019 10:04:49 AM
        SMTP-relay-out 08D751B8BF60FE0B 1               192.168.30.5:57643 192.206.151.51:25 +                                                             11/8/2019 10:04:49 AM
        SMTP-relay-out 08D751B8BF60FE0B 2               192.168.30.5:57643 192.206.151.51:25 <     554 srv8.tinyco.com                                     11/8/2019 10:04:49 AM
        SMTP-relay-out 08D751B8BF60FE0B 3               192.168.30.5:57643 192.206.151.51:25 >     QUIT                                                    11/8/2019 10:04:49 AM
        SMTP-relay-out 08D751B8BF60FE0B 4               192.168.30.5:57643 192.206.151.51:25 -                                       Remote                11/8/2019 10:04:49 AM
        SMTP-relay-out 08D751B8BF60FE0E 0                                  74.125.142.27:25  *                                       attempting to connect 11/8/2019 10:04:52 AM
        SMTP-relay-out 08D751B8BF60FE0E 1               192.168.30.5:57646 74.125.142.27:25  +                                                             11/8/2019 10:04:52 AM
        SMTP-relay-out 08D751B8BF60FE0E 2               192.168.30.5:57646 74.125.142.27:25  <     220 mx.google.com ESMTP                                 11/8/2019 10:04:52 AM
        SMTP-relay-out 08D751B8BF60FE0E 3               192.168.30.5:57646 74.125.142.27:25  >     EHLO outbound.megacorp.net                              11/8/2019 10:04:52 AM
        SMTP-relay-out 08D751B8BF60FE0E 4               192.168.30.5:57646 74.125.142.27:25  <     250-mx.google.com here                                  11/8/2019 10:04:52 AM
        SMTP-relay-out 08D751B8BF60FE0E 5               192.168.30.5:57646 74.125.142.27:25  <     250-SIZE 157286400                                      11/8/2019 10:04:52 AM
        SMTP-relay-out 08D751B8BF60FE0E 6               192.168.30.5:57646 74.125.142.27:25  <     250-8BITMIME                                            11/8/2019 10:04:52 AM
        SMTP-relay-out 08D751B8BF60FE0E 7               192.168.30.5:57646 74.125.142.27:25  <     250-STARTTLS                                            11/8/2019 10:04:52 AM
        SMTP-relay-out 08D751B8BF60FE0E 8               192.168.30.5:57646 74.125.142.27:25  <     250-ENHANCEDSTATUSCODES                                 11/8/2019 10:04:52 AM
        SMTP-relay-out 08D751B8BF60FE0E 9               192.168.30.5:57646 74.125.142.27:25  <     250-PIPELINING                                          11/8/2019 10:04:52 AM
        SMTP-relay-out 08D751B8BF60FE0E 10              192.168.30.5:57646 74.125.142.27:25  <     250-CHUNKING                                            11/8/2019 10:04:52 AM
        SMTP-relay-out 08D751B8BF60FE0E 11              192.168.30.5:57646 74.125.142.27:25  <     250 SMTPUTF8                                            11/8/2019 10:04:52 AM
        SMTP-relay-out 08D751B8BF60FE0E 12              192.168.30.5:57646 74.125.142.27:25  *     22068745                          sending message       11/8/2019 10:04:52 AM
        SMTP-relay-out 08D751B8BF60FE0E 13              192.168.30.5:57646 74.125.142.27:25  >     MAIL FROM:<alert@megacorp.net>                          11/8/2019 10:04:52 AM
        SMTP-relay-out 08D751B8BF60FE0E 15              192.168.30.5:57646 74.125.142.27:25  >     RCPT TO:<mmax@megacorp.net>                             11/8/2019 10:04:52 AM
        SMTP-relay-out 08D751B8BF60FE0E 16              192.168.30.5:57646 74.125.142.27:25  >     RCPT TO:<gmacluskie@megacorp.net>                       11/8/2019 10:04:52 AM
        SMTP-relay-out 08D751B8BF60FE0E 17              192.168.30.5:57646 74.125.142.27:25  >     RCPT TO:<ppalliser@megacorp.net>                        11/8/2019 10:04:52 AM                  
    

#>
[Cmdletbinding(DefaultParameterSetName = "FilePath")]
Param(

    [Parameter(
        ParameterSetName = "Filepath",
        Position = 0,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "One or more Exchange send/receive connector logfile paths, or file objects"
    )]
    [Alias("Name","FullName")]
    [ValidateNotNullOrEmpty()]
    [object[]]$Logfile,

    [Parameter(
        ParameterSetName = "Content",
        Position = 0,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "Content of one or more Exchange send/receive connector logfiles"
    )]
    [ValidateNotNullOrEmpty()]
    [object[]]$Log,

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
    $Activity = "Convert EXchange Send/Receive connector logfile to PSObject"
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

                If ($FileContent -match "#Software: Microsoft Exchange Server" -and $FileContent -match "#Log-type: SMTP") {
                    
                    $Total = ($FileContent -as [array]).Count
                    $Current = 0
                    Try {
                    
                        $Msg = "Parse logfile content and output PSObject"
                        "[$FileName] $Msg" | Write-MessageInfo -FGColor White

                        $Param_WP.CurrentOperation = $Msg
                        Write-Progress @Param_WP

                        #$Fields = (($FileContent | where {$_ -match "#Fields"}) -split ":")[1].trim() -split ","
                        $Header = ((($FileContent | Select-String "#Fields:") -Split ":")[1]).Trim() -Split ","

                        $Output = ($Filecontent | Select-String -Pattern "#" -NotMatch) | 
                            ConvertFrom-CSV -Header $Header | Select-Object * ,
                                @{Name="DateTime";Expression = {$_.'date-time' -as [datetime]}} | 
                                    Select * -ExcludeProperty Date-Time
    
                        $Msg = "Successfully converted $Total logfile entries to a PSObject"
                        "[$FileName] $Msg" | Write-MessageInfo -FGColor Green
                        Write-Progress -Activity $Activity -Completed
                            
                        Write-Output $Output #| Select-Object * ,@{Name="DateTime";Expression = {$_.'date-time' -as [datetime]}} | Select * -ExcludeProperty Date-Time
                
                    }
                    Catch {
                        $Msg = "Failed to parse content for file '$File'"
                        If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                        "[$FileName] $Msg" | Write-MessageError
                    }
                } 
                Else {
                    $Msg = "File '$File' does not appear to be an Exchange send/receive connector SMTP logfile"
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
