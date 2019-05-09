#requires -Version 3
Function ConvertTo-PKCSV {
<#
.SYNOPSIS
    Performs ConvertTo-CSV on an input object, with customizeable delimiter and options to remove header row/quotes

.DESCRIPTION
    Performs ConvertTo-CSV on an input object, with customizeable delimiter and options to remove header row/quotes
    Handy when you want to paste output into a spreadsheet or table and not remember all the necessary parameters or perform string manipulation

.NOTES
    Name    : Function_ConvertTo-PKCSV.ps1 
    Created : 2019-04-11
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK ** 

        v01.00.0000 - 2019-04-11 - Created script


.PARAMETER Object
    PSbject to convert to CSV

.PARAMETER Delimiter
    Delimiter (defaults to culture standard)

.PARAMETER NoQuotes
    Return CSV output without quotes

.PARAMETER NoHeader
    Return CSV output without first row

.PARAMETER Quiet
    Suppress non-verbose console output

.EXAMPLE
    PS C:\> Get-Service | Where-Object {$_.Status -ne "Running"} | Select-Object Status,StartType,Name,DisplayName | ConvertTo-PKCSV -Verbose | Select-Object -First 20
    # Returns CSV output not much differently from ConvertTo-CSV on its own, although it uses -NoTypeInformation by default for convenience

        VERBOSE: PSBoundParameters: 
	
        Key              Value          
        ---              -----          
        Verbose          True           
        Object                          
        Delimiter        ,              
        NoHeader         False          
        NoQuotes         False          
        Quiet            False          
        ParameterSetName                
        PipelineInput    True           
        ScriptName       ConvertTo-PKCSV
        ScriptVersion    1.0.0          

        BEGIN  : Convert object to CSV with the following options: Specify -NoTypeInformation, use default culture delimiter
        VERBOSE: 153 items in object
        "Status","StartType","Name","DisplayName"
        "Stopped","Manual","AJRouter","AllJoyn Router Service"
        "Stopped","Manual","ALG","Application Layer Gateway Service"
        "Stopped","Manual","AppIDSvc","Application Identity"
        "Stopped","Manual","AppMgmt","Application Management"
        "Stopped","Manual","AppReadiness","App Readiness"
        "Stopped","Disabled","AppVClient","Microsoft App-V Client"
        "Stopped","Manual","AppXSvc","AppX Deployment Service (AppXSVC)"
        "Stopped","Manual","aspnet_state","ASP.NET State Service"
        "Stopped","Manual","AssignedAccessManagerSvc","AssignedAccessManager Service"
        "Stopped","Manual","AxInstSV","ActiveX Installer (AxInstSV)"
        "Stopped","Manual","BcastDVRUserService_7ba27","GameDVR and Broadcast User Service_7ba27"
        "Stopped","Disabled","BDESVC","BitLocker Drive Encryption Service"
        "Stopped","Manual","BluetoothUserService_7ba27","Bluetooth User Support Service_7ba27"
        "Stopped","Manual","camsvc","Capability Access Manager Service"
        "Stopped","Manual","CaptureService_7ba27","CaptureService_7ba27"
        "Stopped","Manual","ClipSVC","Client License Service (ClipSVC)"
        "Stopped","Manual","COMSysApp","COM+ System Application"
        "Stopped","Manual","cphs","Intel(R) Content Protection HECI Service"
        "Stopped","Manual","CscService","Offline Files"
        END    : Convert object to CSV with the following options: Specify -NoTypeInformation, use default culture delimiter

.EXAMPLE
    PS C:\> Get-Service | Where-Object {$_.Status -ne "Running"} | Select-Object Status,StartType,Name,DisplayName | ConvertTo-PKCSV -NoHeader -NoQuotes | Select-Object -first 20
    # Returns CSV output without the quote marks or header row (useful when pasting into a spreadsheet, for example)

        WARNING: Parameter -NoHeader simply skips the first row from the CSV output; it cannot verify that it is a header!
        BEGIN  : Convert object to CSV with the following options: Specify -NoTypeInformation, use default culture delimiter, skip first row as header, remove quotation marks
        Stopped,Manual,AJRouter,AllJoyn Router Service
        Stopped,Manual,ALG,Application Layer Gateway Service
        Stopped,Manual,AppIDSvc,Application Identity
        Stopped,Manual,AppMgmt,Application Management
        Stopped,Manual,AppReadiness,App Readiness
        Stopped,Disabled,AppVClient,Microsoft App-V Client
        Stopped,Manual,AppXSvc,AppX Deployment Service (AppXSVC)
        Stopped,Manual,aspnet_state,ASP.NET State Service
        Stopped,Manual,AssignedAccessManagerSvc,AssignedAccessManager Service
        Stopped,Manual,AxInstSV,ActiveX Installer (AxInstSV)
        Stopped,Manual,BcastDVRUserService_7ba27,GameDVR and Broadcast User Service_7ba27
        Stopped,Disabled,BDESVC,BitLocker Drive Encryption Service
        Stopped,Manual,BluetoothUserService_7ba27,Bluetooth User Support Service_7ba27
        Stopped,Manual,camsvc,Capability Access Manager Service
        Stopped,Manual,CaptureService_7ba27,CaptureService_7ba27
        Stopped,Manual,ClipSVC,Client License Service (ClipSVC)
        Stopped,Manual,COMSysApp,COM+ System Application
        Stopped,Manual,cphs,Intel(R) Content Protection HECI Service
        Stopped,Manual,CscService,Offline Files
        Stopped,Automatic,dbupdate,Dropbox Update Service (dbupdate)
        END    : Convert object to CSV with the following options: Specify -NoTypeInformation, use default culture delimiter, skip first row as header, remove quotation marks

.EXAMPLE
    PS C:\> ConvertTo-PKCSV -Object $Arr -Delimiter ";" -NoQuotes -Quiet
    # Returns CSV output using a semicolon as delimiter due to commas in input object, removes quotes, and suppresses all console output

        Name;SamAccountName;EmailAddress;DistinguishedName
        Ana Espinosa Diaz;aespinosa;aespinosa@domain.com;CN=Ana Espinosa Diaz,OU=FTE,OU=AllUsers,DC=Domain,DC=Local
        Alison Krasney;akrasney;akrasney@domain.com;CN=Alison Krasney,OU=Contract,OU=AllUsers,DC=Domain,DC=Local
        Payal Gupta;pgupta;pgupta@domain.com;CN=Payal Gupta,OU=FTE,OU=AllUsers,DC=Domain,DC=Local
        Jason Robert Steinman;rsteinman;rsteinman@domain.com;CN=Robert Steinman,OU=FTE,OU=AllUsers,DC=Domain,DC=Local

#>
[CmdletBinding()]
Param(
    [Parameter(
        Mandatory = $True,
        ValueFromPipeline = $True,
        HelpMessage = "PSObject to convert to CSV"
    )]
    $Object,
    
    [Parameter(
        HelpMessage = "Delimiter (defaults to culture standard)"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateLength(1,1)]
    [string]$Delimiter,

    [Parameter(
        HelpMessage = "Return CSV output without first/header row"
    )]
    [Alias("RemoveHeader")]
    [switch]$NoHeader,

    [Parameter(
        HelpMessage = "Return CSV output without quotes"
    )]
    [Alias("RemoveQuotes")]
    [switch]$NoQuotes,

    [Parameter(
        HelpMessage = "Suppress non-verbose console output"
    )]
    [Switch]$Quiet
)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput

    # Display parameters
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("PipelineInput",$PipelineInput)    
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Make sure we know why we may be missing some data...
    If ($NoHeader.IsPresent -and -not ($Quiet.IsPresent)) {
        $Msg = "Parameter -NoHeader simply skips the first row from the CSV output; it cannot verify that it is a header!"
        Write-Warning $Msg
    }
    
    # Splat
    $Param_Convert = @{}
    $Param_Convert = @{
        NoTypeInformation = $True
        ErrorAction       = "Stop"
        Verbose           = $False
    }
    If ($CurrentParams.Delimiter) {
        $Param_Convert.Add("Delimiter",$Delimiter)
    }
    Else {
        $Delimiter =  (Get-Culture).TextInfo.ListSeparator
        $Param_Convert.Add("UseCulture",$True)
    }

    # Because we want the same behavior whether the input is via parameter or pipeline, we'll collect everything
    $Collection = @()

    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Actions = @()
    $Actions += "Convert object to CSV with the following options: Specify -NoTypeInformation" 
    If ($Delimiter -ne ',') {$Actions += "use nonstandard delimiter '$Delimiter'"}
    Else {$Actions += "use default culture delimiter"}
    If ($NoHeader.IsPresent) {$Actions += "skip first row as header"}
    If ($NoQuotes.IsPresent) {$Actions += "remove quotation marks"}
    $Activity = $Actions -join(", ")
    $Msg = "BEGIN  : $Activity"
    $FGColor = "Yellow"
    If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg}

}
Process {
    
    # We're going to collect everything in the pipeline into a single object so we can run the actual cmdlet in End and not get
    # each individual row output to CSV with the header repeated (!)
    If ($PipelineInput.IsPresent) {
        $Collection += $_
    }
    Else {
        $Collection += $Object
    }
}
End {
    
    $Msg =  "$($Collection.count) items in object"
    Write-Verbose $Msg

    Try {
        If ($Output = $Collection | ConvertTo-CSV @Param_Convert) {
            If ($NoQuotes.IsPresent) {$Output = $Output.Replace('"',$Null)}
            If ($Output) {
                If ($NoHeader.IsPresent) {
                    $Output = $Output | Select-Object -Skip 1
                    If ($Output) {
                        Write-Output $Output
                    }
                }
                Else {Write-Output $Output}
            }
            Else {
                $Msg = "No output was created with your selections; please verify input object type is compatible with ConvertTo-CSV"
                Write-Warning $Msg
            }
        }
    }
    Catch {
        $Msg = "Operation failed"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
        If (-not $Quiet.IsPresent) {$Host.UI.WriteErrorLine("ERROR  : $Msg")}
        Else {Write-Verbose $Msg}
    }
    Finally {
        $Msg = "END    : $Activity"
        $FGColor = "Yellow"
        If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
        Else {Write-Verbose $Msg}
    }
}
} #end ConvertTo-PKCSV

