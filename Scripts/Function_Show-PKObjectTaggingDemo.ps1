Function Show-PKObjectTaggingDemo {
<#
.SYNOPSIS
    Demonstrates the differences between using Select-Object and Add-Member to add properties to an output object

.DESCRIPTION
    Demonstrates the differences between using Select-Object and Add-Member to add properties to an output object
    No parameters, no output

.NOTES
    Name    : Function_Show-PKObjectTaggingDemo.ps1
    Created : 2017-12-01
    Version : 01.00.0000
    Author  : Paula Kingsley
    History :
        
        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK

        v01.00.0000 - 2017-12-01 - Created script

.LINK
    http://community.idera.com/powershell/powertips/b/tips/posts/tagging-objects-efficiently

.EXAMPLE
    PS C:\> Show-PKObjectTaggingDemo


#>
[cmdletbinding()]
Param()

Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

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


    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "Demo: Tagging Objects Efficiently"
    $FGColor = "Yellow"
    $Host.UI.WriteLine($FGColor,$BGColor,$Msg)
    $Host.UI.WriteLine()

}

Process {
    
    $Host.UI.WriteLine()
    $Msg ="Occasionally you see scripts that use Select-Object to append information to existing objects.`nThis can look similar to the code below:"
    $HS = @'

        PS C:\> $Select = Get-Process |
                    Select-Object -Property *, Sender|
                        ForEach-Object { 
                            $_.Sender = $env:COMPUTERNAME
                            $_
                        }
        
        # or

        PS C:\> $Select = Get-Process |
                    Select-Object -Property *,@{N="Sender";E={$env:COMPUTERNAME}}
'@
    
    $FGColor = "Cyan"
    $Host.UI.WriteLine($FGColor,$BGColor,$Msg)
    $FGColor = "White"
    $Host.UI.WriteLine($FGColor,$BGColor,$HS)
    $Host.UI.WriteLine()


    $Msg = "It does work, but Select-Object creates a complete object copy, so this approach is slow and changes the object type.`nYou’ll notice that PowerShell no longer outputs the process objects with the usual table design because of this:"
    $HS = $HS = @'
        
        PS C:\>$Select

            Name                       : AGSService
            Id                         : 4296
            PriorityClass              : Normal
            FileVersion                : 4.4.0.652
            HandleCount                : 291
            WorkingSet                 : 3104768
            PagedMemorySize            : 4444160
            PrivateMemorySize          : 4444160
            VirtualMemorySize          : 91815936
            TotalProcessorTime         : 00:00:07.7500000
            SI                         : 0
            Handles                    : 291
            VM                         : 91815936
            WS                         : 3104768
            PM                         : 4444160
            NPM                        : 16760
            Path                       : C:\Program Files (x86)\Common Files\Adobe\AdobeGCClient\AGSService.exe
            Company                    : Adobe Systems, Incorporated
            CPU                        : 7.75
            ProductVersion             : 4.4.0.652 BuildVersion: 4.4; BuildDate: Wed Aug 23 2017 10:59:31
            Description                : Adobe Genuine Software Integrity Service
            Product                    : Adobe Genuine Software Integrity Service
            __NounName                 : Process
            BasePriority               : 8
            ExitCode                   : 
            HasExited                  : False
        
            <snip>

            Sender                     : WORKSTATION19        
'@
    
    $FGColor = "Cyan"
    $Host.UI.WriteLine($FGColor,$BGColor,$Msg)
    $FGColor = "White"
    $Host.UI.WriteLine($FGColor,$BGColor,$HS)
    $Host.UI.WriteLine()

    $Msg = "Add-Member is the premier cmdlet to add more info to existing objects because it does not copy objects and does not change object types.`nJust compare the output from this code:"
    $HS = @'

        PS C:\> $AddMember = Get-Process | 
                    Add-Member -MemberType NoteProperty -Name Sender -Value $env:COMPUTERNAME -PassThru

        PS C:\> $AddMember

                Handles  NPM(K)    PM(K)      WS(K)     CPU(s)     Id  SI ProcessName                                                                           
                -------  ------    -----      -----     ------     --  -- -----------                                                                           
                    291      16     4272       3084       7.75   4296   0 AGSService                                                                            
                    226      21     3784       4556      52.48   4332   0 AppleMobileDeviceService                                                              
                    479      21    13100      18604      39.59  13236   1 ApplicationFrameHost                                                                  
                    299      16     3040       2724       0.33   4240   0 armsvc                                                                                
                    629      55   136752      15628      98.56   1392   1 chrome       

                    <snip>
        
'@

    $FGColor = "Cyan"
    $Host.UI.WriteLine($FGColor,$BGColor,$Msg)
    $FGColor = "White"
    $Host.UI.WriteLine($FGColor,$BGColor,$HS)


    $Msg = "The object type is unchanged, and PowerShell continues to use the default output layout for processes.`nThis, coincidentally is why the new 'Sender' property initially is not initially visible.`nIt is present, though:"
    $HS = @'

        PS C:\> Get-Process |
                    Add-Member -MemberType NoteProperty -Name Sender -Value $env:COMPUTERNAME -PassThru |
                        Select-Object -Property Name, Id, Sender 
                        
                Name                        Id Sender         
                ----                        -- ------         
                AGSService                4296 WORKSTATION19
                AppleMobileDeviceService  4332 WORKSTATION19
                ApplicationFrameHost     13236 WORKSTATION19
                armsvc                    4240 WORKSTATION19
                audiodg                  22644 WORKSTATION19

                <snip>
'@
    
    $FGColor = "Cyan"
    $Host.UI.WriteLine($FGColor,$BGColor,$Msg)
    $FGColor = "White"
    $Host.UI.WriteLine($FGColor,$BGColor,$HS)
    $Host.UI.WriteLine()


    $Msg = "Compare the property names from Select-Object with those from Add-Member:"
    $HS = @'
    
        # Member properties using Select-Object
        PS C:\> ($Select | Get-Member) | Format-Wide Name -Column 3

            Equals                               GetHashCode                          GetType                             
            ToString                             BasePriority                         Company                             
            Container                            CPU                                  Description                         
            EnableRaisingEvents                  ExitCode                             ExitTime                            
            FileVersion                          Handle                               HandleCount                         
            Handles                              HasExited                            Id                                  
            MachineName                          MainModule                           MainWindowHandle                    
            MainWindowTitle                      MaxWorkingSet                        MinWorkingSet                       
            Modules                              Name                                 NonpagedSystemMemorySize            
            NonpagedSystemMemorySize64           NPM                                  PagedMemorySize                     
            PagedMemorySize64                    PagedSystemMemorySize                PagedSystemMemorySize64             
            Path                                 PeakPagedMemorySize                  PeakPagedMemorySize64               
            PeakVirtualMemorySize                PeakVirtualMemorySize64              PeakWorkingSet                      
            PeakWorkingSet64                     PM                                   PriorityBoostEnabled                
            PriorityClass                        PrivateMemorySize                    PrivateMemorySize64                 
            PrivilegedProcessorTime              ProcessName                          ProcessorAffinity                   
            Product                              ProductVersion                       Responding                          
            SafeHandle                           Sender                               SessionId                           
            SI                                   Site                                 StandardError                       
            StandardInput                        StandardOutput                       StartInfo                           
            StartTime                            SynchronizingObject                  Threads                             
            TotalProcessorTime                   UserProcessorTime                    VirtualMemorySize                   
            VirtualMemorySize64                  VM                                   WorkingSet                          
            WorkingSet64                         WS                                   __NounName 


        
        # Member properties using Add-Member
        PS C:\> ($AddMember | Get-Member) | Format-Wide Name -Column 3

            Handles                              Name                                 NPM                                 
            PM                                   SI                                   VM                                  
            WS                                   Disposed                             ErrorDataReceived                   
            Exited                               OutputDataReceived                   BeginErrorReadLine                  
            BeginOutputReadLine                  CancelErrorRead                      CancelOutputRead                    
            Close                                CloseMainWindow                      CreateObjRef                        
            Dispose                              Equals                               GetHashCode                         
            GetLifetimeService                   GetType                              InitializeLifetimeService           
            Kill                                 Refresh                              Start                               
            ToString                             WaitForExit                          WaitForInputIdle                    
            Sender                               __NounName                           BasePriority                        
            Container                            EnableRaisingEvents                  ExitCode                            
            ExitTime                             Handle                               HandleCount                         
            HasExited                            Id                                   MachineName                         
            MainModule                           MainWindowHandle                     MainWindowTitle                     
            MaxWorkingSet                        MinWorkingSet                        Modules                             
            NonpagedSystemMemorySize             NonpagedSystemMemorySize64           PagedMemorySize                     
            PagedMemorySize64                    PagedSystemMemorySize                PagedSystemMemorySize64             
            PeakPagedMemorySize                  PeakPagedMemorySize64                PeakVirtualMemorySize               
            PeakVirtualMemorySize64              PeakWorkingSet                       PeakWorkingSet64                    
            PriorityBoostEnabled                 PriorityClass                        PrivateMemorySize                   
            PrivateMemorySize64                  PrivilegedProcessorTime              ProcessName                         
            ProcessorAffinity                    Responding                           SafeHandle                          
            SessionId                            Site                                 StandardError                       
            StandardInput                        StandardOutput                       StartInfo                           
            StartTime                            SynchronizingObject                  Threads                             
            TotalProcessorTime                   UserProcessorTime                    VirtualMemorySize                   
            VirtualMemorySize64                  WorkingSet                           WorkingSet64                        
            PSConfiguration                      PSResources                          Company                             
            CPU                                  Description                          FileVersion                         
            Path                                 Product                              ProductVersion                      

'@

    $FGColor = "Cyan"
    $Host.UI.WriteLine($FGColor,$BGColor,$Msg)
    $FGColor = "White"
    $Host.UI.WriteLine($FGColor,$BGColor,$HS)
    $Host.UI.WriteLine()


    $Msg = "And compare the processing time!"
    $HS = @'

        PS C:\ $SelectTime = Measure-Command -Expression {Get-Process |Select-Object -Property *, Sender|ForEach-Object {  $_.Sender = $env:COMPUTERNAME;$_}}
        PS C:\> $AddMemberTime = Measure-Command -Expression {Get-Process | Add-Member -MemberType NoteProperty -Name Sender -Value $env:COMPUTERNAME -PassThru}
        
        
        PS C:\> $SelectTime

            Days              : 0
            Hours             : 0
            Minutes           : 0
            Seconds           : 6
            Milliseconds      : 28
            Ticks             : 60281826
            TotalDays         : 6.97706319444444E-05
            TotalHours        : 0.00167449516666667
            TotalMinutes      : 0.10046971
            TotalSeconds      : 6.0281826
            TotalMilliseconds : 6028.1826
    
        PS C:\> $ $AddMemberTime

            Days              : 0
            Hours             : 0
            Minutes           : 0
            Seconds           : 0
            Milliseconds      : 14
            Ticks             : 146004
            TotalDays         : 1.68986111111111E-07
            TotalHours        : 4.05566666666667E-06
            TotalMinutes      : 0.00024334
            TotalSeconds      : 0.0146004
            TotalMilliseconds : 14.6004    
    
'@

    
    $FGColor = "Cyan"
    $Host.UI.WriteLine($FGColor,$BGColor,$Msg)
    $FGColor = "White"
    $Host.UI.WriteLine($FGColor,$BGColor,$HS)
    $Host.UI.WriteLine()

}

}
