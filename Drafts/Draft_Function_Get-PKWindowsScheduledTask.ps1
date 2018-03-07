Function Get-PKWindowsScheduledTask {
<#
.SYNOPSIS
   Returns scheduled tasks on a local or remote computer, and optionally sends an HTML report to a recipient

.DESCRIPTION
   Returns scheduled tasks on a local or remote computer, and optionally sends an HTML report to a recipient
   This script uses the Schedule.Service COM-object to query computer to gather a formatted 
   list including the Author, UserId, and description of the task 
   This information is parsed from the XML attributed to providing a more human-readable format 
   The function must be run as administrator


.NOTES        
    Name    : Get-PKWindowsScheduledTask.ps1
    Created : 2018-01-24
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2018-01-24 - Created script from Jaap Brasser / Jose Ortega's original


.INPUTS
    $computername => If you don't provide a name it will run on the localhost.
    $File => File with the names of the Computers, in case you want to run it on several computers

   Switches:
    $RootFolder => Added in original script in version 1.2: "Added the -RootFolder switch, this only gathers the tasks in the root folder instead of all subfolders. This is similar to the behavior of the script in the 1.1 version."
    $Objs => Added so you can get the objects and use them as you want (Filter, csv, json,HTML, etc.)
    $Html => Get HTML File from the original Object. $OBjs
    $Email => Send Emails (required some additional info, From, To, and server)

.LINK
    https://gallery.technet.microsoft.com/scriptcenter/Get-scheduled-tasks-from-1ba9b7fa


.EXAMPLE
    .\Get-ScheduledTask.ps1
   To run the script against the local Machine, assuming that your computer is called "SVR". this will create the file: TaskInfo-SRV-20170718.html

.EXAMPLE
    .\Get-ScheduledTask.ps1 -Computername MyRemoteMachine
    To run the script on a remote machine (Single remote machine)
    ** Requires an AD CS (a Domain)
    ** Run the script as admin
    Assuming that your remote computer is called "MyRemoteMachine." this will create the file: TaskInfo-MyRemoteMachine-20170718.html

.EXAMPLE
    .\Get-ScheduledTask.ps1 -File .\MyComputers.txt 
    ** Same requirements as example2.
   Assuming that your remote computers are called "Remote1" and "Remote2". This will create the files: TaskInfo-Remote1-20170718.html and TaskInfo-Remote2-20170718.html

.EXAMPLE
    .\Get-ScheduledTaskv3.ps1 -Objs | where{$_.Name -eq "AppleSoftwareUpdate"}

    You can filter by: NextRunTime, Author, Trigger, State, UserId, Actions, Name, LastRunTime, LastTaskResult, Description, NumberOfMissedRuns, Enabled, Path, ComputerName
.EXAMPLE
    .\Get-ScheduledTaskv3.ps1 -Objs | where{$_.Name -eq "AppleSoftwareUpdate"} | ConvertTo-Html | Out-File here.html

    Export to Html.

#>
[CmdletBinding(
        DefaultParameterSetName="SingleHost"
    )]
param(
    
    [Parameter(
        
    )]
    [Parameter(
        position=0,
        ParameterSetName='SingleHost',
        Mandatory=$false,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True
    )]
    [ValidateNotNullOrEmpty()]
    [string[]]$computername = "localhost",
    
    [Parameter(
        position=0,
        parameterSetName='MultiHost',
        Mandatory=$true,
        ValueFromPipeline=$True
    )]
    [ValidateNotNullOrEmpty()]
    $File,
    
    [Parameter(
        position=1,
        Mandatory=$False,
        HelpMessage = "Search from root folder (will include a lot of tasks)"
    )]
    [switch]$RootFolder,

    [Parameter(
        position=2,
        Mandatory=$False,
        HelpMessage = "I have no idea"
        
    )]
    [switch]$Objs,

    [Parameter(
        position=3,
        Mandatory=$False,
        HelpMessage = "Absolute path to logfile folder"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$LogPath = $Env:Temp,
    
    [Parameter(
        position=4,
        Mandatory = $False,
        HelpMessage = "Send email report on completion"
    )]
    [switch]$Email,

    [Parameter(
        position=4,
        Mandatory = $False,
        HelpMessage = "Sender SMTP address"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$From="paula.kingsley@nielsen.com",

    [Parameter(
        position=7,
        Mandatory = $False,
        HelpMessage = "Recipient SMTP address"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$To="paula.kingsley@nielsen.com",

    [Parameter(
        position=8,
        Mandatory = $False,
        HelpMessage = "SMTP relay server"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$SMTPServer="mailhost.globix-sc.gracenote.com",

    [Parameter(
        Mandatory=$False,
        HelpMessage="Valid credentials on target computer"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential] $Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(
        ParameterSetName = "Job",
        Mandatory = $False,
        HelpMessage = "Run as remote PSJob"
    )]
    [Switch] $AsJob,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Don't test WinRM connectivity before invoking command"
    )]
    [Switch] $SkipConnectionTest,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Hide all non-verbose/non-error console output"
    )]
    [Switch] $SuppressConsoleOutput

)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    $ScriptName = $MyInvocation.MyCommand.Name
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ParameterSetNAme",$Source)
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference    = "Continue"
    
    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    $Results = @()

    #region Variables

    $global:ScriptLocation = $(get-location).Path
    $global:DefaultLog = "$global:ScriptLocation\ScheduledTasks.log"
    $global:Path=$(get-location).Path
    $global:AllTheFiles=@()
    $global:AllTheObjs=@()


    #endregion Variables


    #region Internal functions

    function Write-Log{
        [CmdletBinding()]
        [OutputType([int])]
        Param(
            [Parameter(
                Mandatory=$true, 
                ValueFromPipelineByPropertyName=$true, 
                Position=0
            )] 
            [ValidateNotNullOrEmpty()] 
            [Alias("LogContent")] 
            [string]$Message,
            
            
            [Parameter(
                Mandatory=$false, 
                ValueFromPipelineByPropertyName=$true,
                Position=1
            )] 
            [Alias('LogPath')] 
            [string]$Path=$global:DefaultLog,

            [Parameter(
                Mandatory=$false, 
                ValueFromPipelineByPropertyName=$true,
                Position=2
            )] 
            [ValidateSet("Error","Warn","Info","Load","Execute")] 
            [string]$Level="Info",
            
            [Parameter(
                Mandatory=$False, 
                ValueFromPipelineByPropertyName=$true, 
                Position=3
            )] 
            [ValidateNotNullOrEmpty()] 
            [string]$ComputerName = $Computer,


            [Parameter(
                Mandatory=$false
            )] 
            [switch]$NoClobber
        )

         Process{
            $BGColor = $host.UI.RawUI.BackgroundColor

            if ((Test-Path $Path) -AND $NoClobber) {
                Write-Warning "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name."
                Return
                }

            # If attempting to write to a log file in a folder/path that doesn't exist
            # to create the file include path.
            elseif (!(Test-Path $Path)) {
                Write-Verbose "Creating $Path."
                $NewLogFile = New-Item $Path -Force -ItemType File -ErrorAction Stop
            }
            else {
                # Nothing to see here yet.
            }

            # Now do the logging and additional output based on $Level
            switch ($Level) {
                'Error' {
                    #Write-Warning $Message
                    $Host.UI.WriteErrorLine($Message)
                    Write-Output "[$ComputerName] [$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] ERROR: `t $Message" | Out-File -FilePath $Path -Append
                    }
                'Warn' {
                    Write-Warning $Message
                    Write-Output "[$ComputerName] [$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] WARNING: `t $Message" | Out-File -FilePath $Path -Append
                    }
                'Info' {
                    $FGColor = "Green"
                    $Host.UI.WriteLine($FGColor,$BGColor,$Message)
                    #Write-Host $Message -ForegroundColor Green
                    Write-Verbose $Message
                    Write-Output "[$ComputerName] [$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] INFO: `t $Message" | Out-File -FilePath $Path -Append
                    }
                'Load' {
                    $FGColor = "Magenta"
                    $Host.UI.WriteLine($FGColor,$BGColor,$Message)
                    #Write-Host $Message -ForegroundColor Magenta
                    Write-Verbose $Message
                    Write-Output "[$ComputerName] [$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] LOAD: `t $Message" | Out-File -FilePath $Path -Append
                    }
                'Execute' {
                    $FGColor = "Green"
                    $Host.UI.WriteLine($FGColor,$BGColor,$Message)
                    #Write-Host $Message -ForegroundColor Green
                    Write-Verbose $Message
                    Write-Output "[$ComputerName] [$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] EXEC: `t $Message" | Out-File -FilePath $Path -Append
                    }
                }
        }
    }

    function Get-AllTaskSubFolders {
        [cmdletbinding()]
        param (
            # Set to use $Schedule as default parameter so it automatically list all files
            # For current schedule object if it exists.
            $FolderRef = $Schedule.getfolder("\")
        )
        if ($FolderRef.Path -eq '\') {
            $FolderRef
        }
        if (-not $RootFolder) {
            $ArrFolders = @()
            if(($Folders = $folderRef.getfolders(1))) {
                $Folders | ForEach-Object {
                    $ArrFolders += $_
                    if($_.getfolders(1)) {
                        Get-AllTaskSubFolders -FolderRef $_
                    }
                }
            }
            $ArrFolders
        }
    }
    
    function Get-TaskTrigger {
        [cmdletbinding()]
        param (
            $Task
        )
        $Triggers = ([xml]$Task.xml).task.Triggers
        if ($Triggers) {
            $Triggers | Get-Member -MemberType Property | ForEach-Object {
                $Triggers.($_.Name)
            }
        }
    }
    
    function Get-TaskTimestampInfo {
        [cmdletbinding()]
        param (
            [Parameter(
                ValueFromPipeline=$True,
                position=0,
                mandatory=$true
            )]
            $AllFolders
        )
	    PROCESS{
		    foreach ($Folder in $AllFolders) {
			    if (($Tasks = $Folder.GetTasks(1))) {
				    $Tasks | Foreach-Object {
                         $xml = [xml]$_.xml;
					    New-Object -TypeName PSCustomObject -Property @{
					    'Name' = $_.name
					    'Path' = $_.path
					    'State' = switch ($_.State) {
						    0 {'Unknown'}
						    1 {'Disabled'}
						    2 {'Queued'}
						    3 {'Ready'}
						    4 {'Running'}
						    Default {'Unknown'}
					    }
					    'Enabled' = $_.enabled
					    'LastRunTime' = $_.lastruntime
					    'LastTaskResult' = $_.lasttaskresult
					    'NumberOfMissedRuns' = $_.numberofmissedruns
					    'NextRunTime' = $_.nextruntime
                        'Actions' = ($xml.Task.Actions.Exec | % { "$($_.Command) $($_.Arguments)" }) -join "`n"
					    'Author' =  ([xml]$_.xml).Task.RegistrationInfo.Author
					    'UserId' = ([xml]$_.xml).Task.Principals.Principal.UserID
					    'Description' = ([xml]$_.xml).Task.RegistrationInfo.Description
                        'Trigger' = Get-TaskTrigger -Task $_
                        'ComputerName' = $Schedule.TargetServer
					    }
				    }
			    }
		    }
	    }
	    END{
            return $Tasks
	    }
    }
    
    function Get-TaskInfoHTML {
        [cmdletbinding()]
        param (
            [Parameter(
                ValueFromPipeline=$True,
                position=0,mandatory=$true
            )]
            $AllFolders,
            
            [Parameter(
                ValueFromPipeline=$True,
                position=1,
                mandatory=$false
            )]
            $computer=$env:COMPUTERNAME
        
           )
        BEGIN{
            $tasks =Get-TaskTimestampInfo $AllFolders
            $global:TodayDate= [Datetime]::Now.ToString("yyyyMMdd") # (20170917)
            $global:Time=[Datetime]::Now.ToString("hhmmss") #  (001305)
            $CompName=$computer
            $FnHTML= "TaskInfo-$CompName-$global:TodayDate$global:Time.html"
            $global:AllTheFiles+="$global:Path\$FnHTML"
            $title= "Task Information from $computer on $global:TodayDate // $global:Time"
            $header= "<style type=""text/css"">{margin:0;padding:0}@import url(https://fonts.googleapis.com/css?family=Indie+Flower|Josefin+Sans|Orbitron:500|Yrsa);body{text-align:center;font-family:14px/1.4 'Indie Flower','Josefin Sans',Orbitron,sans-serif;font-family:'Indie Flower',cursive;font-family:'Josefin Sans',sans-serif;font-family:Orbitron,sans-serif}#page-wrap{margin:50px}tr:nth-of-type(odd){background:#eee}th{background:#EF5525;color:#fff;font-family:Orbitron,sans-serif}td,th{padding:6px;border:1px solid #ccc;text-align:center;font-size:large}table{width:90%;border-collapse:collapse;margin-left:auto;margin-right:auto;font-family:Yrsa,serif}</style>";
        }
	    END{
        $html5Text="<!DOCTYPE HTML>
    <html lang=""en-US"">
    <head>
	    <meta charset=""UTF-8"">
	    <title>" + $title + "</title>
        " + $header + "
    </head>
    <body>
    <h1>" + $title + "</h1>
    <table>
    <tr><th>Name</th><th>State</th><th>Enabled</th><th>Author</th><th>UserId</th><th>NextRunTime</th><th>LastRunTime</th><th>LastTaskResult</th><th># MissedRuns</th><th>Trigger</th><th>Actions</th><th>Description</th><th>Path</th></tr>
    ";

    foreach($task in $tasks){
        $nm=$task.Name
        $pt=$task.Path
        $st=$task.State
        $en=$task.Enabled
        $missed=$task.NumberOfMissed
        $des=$task.Description
        $aut=$task.Author
        $nt=$task.NextRunTime
        $lt=$task.LastRunTime
        $uid=$task.UserId
        $ltr=$task.LastTaskResult
        $act=$task.Actions
        $tri = $task.Trigger
        $html5Text+="<tr> <td>$nm</td><td>$st</td><td>$en</td><td>$aut</td><td>$uid</td><td>$nt</td><td>$lt</td><td>$ltr</td><td>$missed</td><td>$tri</td><td>$act</td><td>$des</td><td>$pt</td></tr>"
    }
	
    $html5Text+="
    </table>
    </body>
    </html>"

    $html5Text | Out-File "$global:Path\$FnHTML"    
    #    Write-Output  $task  | select * | Sort-Object LastRunTime -descending |  ConvertTo-html * -head $header -Title "Task Information from $computer" | Out-File "$path\$FnHTML"
	    }
    }



#endregion Internal functions







}

Process {

    $MSg = "Log file: $DefaultLog"
    Write-Verbose "[$(Get-Date -f G)]: $Msg"
    Write-Log -Level Info -Message $Msg -ComputerName $Env:ComputerName

    $Msg = "Start script $ScriptName"
    Write-Verbose "[$(Get-Date -f G)]: $Msg"
    Write-Log -Level Info -Message $Msg -ComputerName $Env:ComputerName

    $Msg = "Parameter set name: '$Source'"
    Write-Verbose "[$(Get-Date -f G)]: $Msg"
    Write-Log -Level Info -Message $Msg -ComputerName $Env:ComputerName

    $Msg = "Create COM object"
    Write-Verbose "[$(Get-Date -f G)]: $Msg"
    Write-Log -Level Info -Message $Msg -ComputerName $Env:ComputerName
    
    Try {
        $Schedule = New-Object -ComObject("Schedule.Service") -EA Stop
        $Msg = "Created Schedule.Service COM object"
        Write-Verbose "[$(Get-Date -f G)]: $Msg"
        Write-Log -Level Info $Msg -ComputerName $Env:ComputerName
    }
    Catch {
        $Msg = "Failed to create Schedule.Service COM object; script will now exit"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
        $Host.UI.WriteErrorLine("[$(Get-Date -f G)]: $Msg")
        Write-Log -Level Warn $Msg -ComputerName $Env:ComputerName
        $Msg = "End script $ScriptName"
        Write-Log -Level Info $Msg -ComputerName $Env:ComputerName
        Break
    }


    If ($Source -eq "Multihost") {
        $Msg = "Get computer names from file '$File'"
        Write-Verbose "[$(Get-Date -f G)]: $Msg"
        Write-Log -Level Info -Message $Msg -ComputerName $Env:ComputerName
        Try {
            [Array]$Computers = Get-Content $File    
        }
        Catch {
            $Msg = "Failed to get computer names from file '$File'"
            If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
            $Host.UI.WriteErrorLine("[$(Get-Date -f G)]: $Msg")
            Write-Log -Level Warn $Msg
            $Msg = "End script $ScriptName"
            Write-Log -Level Info $Msg
        }
    }
    Else {
        [array]$Computers = $ComputerName
    }


    Foreach ($Computer in $Computers) {
    
        $Msg = "Gather information for '$Computer'"
        Write-Verbose "[$(Get-Date -f G)]: $Msg"
        Write-Log -Level Info -Message $Msg -ComputerName $Computer

        Try {
                
            $Schedule.connect($computer) 

            #If local computer
            If ($Computer -in @("localhost",$Env:ComputerName,"127.0.0.1")) {
                    
                Try {
                    
                    # If obj or not
                    Switch ($Obj.IsPresent) {
                        $True {
                            $AllFolders = Get-AllTaskSubFolders -EA Stop
                            $tasks = Get-TaskTimestampInfo $AllFolders -EA Stop
                            $tasks  
                        }
                        $False {
                            $AllFolders = Get-AllTaskSubFolders -EA Stop
                            Get-TaskInfoHTML $AllFolders -EA Stop
                        }
                    } #end switch
                    
                    

                }
                Catch {
                    $Msg = "Operation failed"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                    $Host.UI.WriteErrorLine("[$(Get-Date -f G)]: $Msg on $Computer")
                    Write-Log -Level Error $Msg -ComputerName $Computer
                    Break
                
                }
            }
            Else {
                $AllFolders = Get-AllTaskSubFolders
                Get-TaskInfoHTML $AllFolders $computername
            }
        
            $Msg = "Retrieved scheduled task details"
            Write-Verbose "[$(Get-Date -f G)]: $Msg on $Computer"
            Write-Log -Level Info -Message $Msg -ComputerName $Computer
        }
        Catch {
            $Msg = "[$(Get-Date -f G)]: Operation failed on $Computer"
            If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
            $Host.UI.WriteErrorLine($Msg)
            Write-Log -Level Warn $Msg.Replace($Computer,$Null)
            $Msg = "[$(Get-Date -f G)]: End script $ScriptName"
        }
    
    }
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    }
















    Switch ($Source) {
    
        SingleHost {
                
            $Msg = "Gather information for '$ComputerName'"
            Write-Verbose "[$(Get-Date -f G)]: $Msg"
            Write-Log -Level Info -Message $Msg

            Try {
                
                $Schedule.connect($computername) 

                #If local computer
                If ($Computername -in @("localhost",$Env:ComputerName,"127.0.0.1")) {
                    
                    Try {
                    
                        # If obj or not
                        Switch ($Obj.IsPresent) {
                            $True {
                                $AllFolders = Get-AllTaskSubFolders -EA Stop
                                $tasks = Get-TaskTimestampInfo $AllFolders -EA Stop
                                $tasks  
                            }
                            $False {
                                $AllFolders = Get-AllTaskSubFolders -EA Stop
                                Get-TaskInfoHTML $AllFolders -EA Stop
                            }
                        } #end switch
                    
                        $Msg = "Retrieved scheduled task details on $ComputerName"
                        Write-Verbose "[$(Get-Date -f G)]: $Msg"
                        Write-Log -Level Info -Message $Msg

                    }
                    Catch {
                        $Msg = "Operation failed on $ComputerName"
                        If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
                        $Host.UI.WriteErrorLine("[$(Get-Date -f G)]: $Msg")
                        Write-Log -Level Error $Msg
                        $Msg = "End script $ScriptName"
                        Write-Log -Level Info $Msg
                        Break
                
                    }
                }
                Else {
                    #if it goes this long do a default action
                    Write-Log -Level Load "Gathering information for $computername"
                    $AllFolders = Get-AllTaskSubFolders
                    Get-TasksInfoHTML $AllFolders $computername
            
                }
            }
            Catch {
                $Msg = "[$(Get-Date -f G)]: Operation failed on $Computername"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
                $Host.UI.WriteErrorLine($Msg)
                Write-Log -Level Warn $Msg
                $Msg = "[$(Get-Date -f G)]: End script $ScriptName"
            }


            <#
#----------------------

try{
        #SoloComputername #default
        if($computername -eq "localhost" -and !($Objs)){
            Write-Log -Level Load "Gathering information for $computername"
            $Schedule.connect($ComputerName) 
            $AllFolders = Get-AllTaskSubFolders
            Get-TasksInfoHTML $AllFolders
        }
        
        
        if($computername -eq "localhost" -and $Objs){
            Write-Log -Level Load "Gathering information for $computername"
            $Schedule.connect($computername) 
            $AllFolders = Get-AllTaskSubFolders
            $tasks =Get-TasktsInfo $AllFolders
            $tasks
        }
        else{
            #if it goes this long do a default action
            Write-Log -Level Load "Gathering information for $computername"
            $Schedule.connect($ComputerName) 
            $AllFolders = Get-AllTaskSubFolders
            Get-TasksInfoHTML $AllFolders $computername
        }

    }
    catch{
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        Write-Log -Level Error -Message "Message: $ErrorMessage. Make sure that you're running the script as Administrator (in a elevated PS console)"
    }


            }
            catch{
                $ErrorMessage = $_.Exception.Message
                $FailedItem = $_.Exception.ItemName
                Write-Log -Level Error -Message "Message: $ErrorMessage. Make sure that you're running the script as Administrator (in a elevated PS console)"
            }
    #>
        }
        Multihost {
            
            $Msg = "Get computer names from file '$File'"
            Write-Verbose "[$(Get-Date -f G)]: $Msg"
            Write-Log -Level Info -Message $Msg
            $computers = Get-Content $File

            foreach ($computer in $computers){
        
                try{
                    $Msg = 
                    Write-Log -Level Load "Gather information for $computer"
                    if(!($Objs)){
                        #Write-Log -Level Load "Gathering information for $computer"
                        $Schedule.connect($computer) 
                        $AllFolders = Get-AllTaskSubFolders
                        Get-TaskInfoHTML $AllFolders $computer
                    }

                    if($Objs){
                        #Write-Log -Level Load "Gathering information for $computer"
                        $Schedule.connect($computername) 
                        $AllFolders = Get-AllTaskSubFolders
                        $tasks =Get-TaskTimestampInfo $AllFolders
                        $tasks
                    }
                }
                catch{
                    $ErrorMessage = $_.Exception.Message
                    $FailedItem = $_.Exception.ItemName
                    Write-Log -Level Error -Message "Message: $ErrorMessage. Make sure that you're running the script as Administrator (in a elevated PS console)"
                }
            }
        }
        
    
    
    
    
    
    
    
    
    
    }
    

if($file){
    Write-Log -Level Info "Getting information from the $file file"
    $computers = Get-Content $File

    foreach($computer in $computers){
        
        try{
            Write-Log -Level Load "Gathering information for $computer"
            if(!($Objs)){
                #Write-Log -Level Load "Gathering information for $computer"
                $Schedule.connect($computer) 
                $AllFolders = Get-AllTaskSubFolders
                Get-TaskInfoHTML $AllFolders $computer
            }

            if($Objs){
                #Write-Log -Level Load "Gathering information for $computer"
                $Schedule.connect($computername) 
                $AllFolders = Get-AllTaskSubFolders
                $tasks =Get-TaskTimestampInfo $AllFolders
                $tasks
            }
        }
        catch{
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Log -Level Error -Message "Message: $ErrorMessage. Make sure that you're running the script as Administrator (in a elevated PS console)"
        }
    }

}
else{
    
}


if($Email){
    try{
        if( !($from) -or !($to) -or !($SMTPServer)){
            Write-Log -Level Error -Message "When you use the Option -Email, you need to provide 3 more arguments: -from ""email@j0rt3g4.com"" -to ""to@domain.com"" -server ""smtpserver.domain.com"""
        }
        else{
            Send-MailMessage -From $from -To $to -SmtpServer $SMTPServer -Attachments $global:AllTheFiles -Subject "Report for $global:TodayDate-$global:Time"
        }
    }
    catch{
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        Write-Log -Level Error -Message "Message: $ErrorMessage. Make sure that you're running the script as Administrator (in a elevated PS console)"
    }
}


Write-Log -Level Info "********************************* FINISHED THE SCRIPT *********************************"
}