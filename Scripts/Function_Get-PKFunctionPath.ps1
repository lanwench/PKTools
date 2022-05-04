#requires -Version 4
Function Get-PKFunctionPath {
<#
.SYNOPSIS
    Returns the underlying filename(s) for a PowerShell alias or function, either by command name or module name

.DESCRIPTION
    Returns the underlying filename(s) for a PowerShell alias or function, either by command name or module name
    Gets all files in module (or searches for module containing function), then uses Get-Childitem/Select-String & regular expression 
    Returns a PSObject

.NOTES        
    Name    : Function_Get-PKFunctionPath.ps1
    Created : 2022-04-04
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2022-04-04 - Created script 

.PARAMETER CommandName
    One or more PowerShell function or alias names

.PARAMETER Module
    One or more PowerShell modules names or objects

.PARAMETER Detailed
    Return detailed output, including line numbers, module version, and pattern searched for

.EXAMPLE
    PS C:\> Get-PKFunctionPath -Module kittens -Verbose
   
        Get file paths for functions inside a module 

        VERBOSE: PSBoundParameters: 
	
        Key              Value             
        ---              -----             
        Module           {kittens}           
        Verbose          True              
        CommandName                        
        Detailed         False             
        ScriptName       Get-PKFunctionPath
        ScriptVersion    1.0.0             
        PipelineInput    False             
        ParameterSetName ByModule          

        VERBOSE: [BEGIN: Get-PKFunctionPath] Return underlying file names/paths for PowerShell commands within modules
        
        VERBOSE: [kittens] Get the module object
        VERBOSE: [kittens] Get all commands in module
        VERBOSE: [kittens] Alias 'Get-Zoomy' points to underlying function 'Get-Catnip'
        VERBOSE: [kittens] Get functions in module
        VERBOSE: [kittens] 11 unique function name(s) found in module
        VERBOSE: [kittens] Get *.ps1 & *.psm1 files inside module
        VERBOSE: [kittens] 12 *.ps1 & *.psm1 file(s) found in module path
        VERBOSE: [kittens] Search files for function names
        VERBOSE: [kittens] Get-Catnip: Found match in 'C:\MyModules\Kittens\Scripts\Function_Get-Catnip.ps1'
        VERBOSE: [kittens] Get-Yarn: Found match in 'C:\MyModules\Kittens\Scripts\Function_Get-Yarn.ps1'
        VERBOSE: [kittens] Reset-PaperBag: Found match in 'C:\MyModules\Kittens\Scripts\Function_Reset-PaperBag.ps1'

        VERBOSE: [END: Get-PKFunctionPath] Return underlying file names/paths for PowerShell commands within modules

        Module Command            Filepath                                                  Messages                                                                                
        ------ -------            --------                                                  --------                                                                                
        kittens  Get-Zoomy        -                                                          Alias points to underlying function 'Get-Catnip'                                
        kittens  Get-Catnip       C:\MyModules\Kittens\Scripts\Function_Get-Catnip.ps1      Pattern '^[ \t]*Function[ \t]+Get-Catnip[ \t]*(\(.+\))?[ \t]*\{[ \t]*$' found in ...
        kittens  Get-Yarn         C:\MyModules\Kittens\Scripts\Function_Get-Yarn.ps1        Pattern '^[ \t]*Function[ \t]+Get-Yarn[ \t]*(\(.+\))?[ \t]*\{[ \t]*$' found i...
        kittens  Reset-PaperBag   C:\MyModules\Kittens\Scripts\Function_Reset-PaperBag.ps1  Pattern '^[ \t]*Function[ \t]+Reset-PaperBag[ \t]*(\(.+\))?[ \t]*\{[ \t]*$' fo...
        
.EXAMPLE
    PS C:\> Get-PKFunctionPath -Module ProfileHelpers -Detailed

        Get file paths and detailed information about the functions inside a module

        Module        : ProfileHelpers
        ModuleVersion : 1.0.0
        ModulePath    : C:\MyModules\ProfileHelpers\ProfileHelpers.psd1
        Command       : Open
        Type          : Alias
        Filepath      : 
        LineNumber    : 
        Messages      : Alias points to underlying function 'Get-OpenFiles'

        Module        : ProfileHelpers
        ModuleVersion : 1.0.0
        ModulePath    : C:\MyModules\ProfileHelpers\ProfileHelpers.psd1
        Command       : Get-OpenFiles
        Type          : Function
        Filepath      : C:\MyModules\ProfileHelpers\Scripts\Function_AllFunctions.ps1
        LineNumber    : 79
        Messages      : Pattern '^[ \t]*Function[ \t]+Get-OpenFiles[ \t]*(\(.+\))?[ \t]*\{[ \t]*$' found in 1/2 module file(s)

        Module        : ProfileHelpers
        ModuleVersion : 1.0.0
        ModulePath    : C:\MyModules\ProfileHelpers\ProfileHelpers.psd1
        Command       : Get-AccountLockout
        Type          : Function
        Filepath      : C:\MyModules\ProfileHelpers\Scripts\Function_AllFunctions.ps1
        LineNumber    : 288
        Messages      : Pattern '^[ \t]*Function[ \t]+Get-AccountLockout[ \t]*(\(.+\))?[ \t]*\{[ \t]*$' found in 1/2 module file(s)

.EXAMPLE
    PS C:\> Get-PKFunctionPath -CommandName Get-Yarn -Detailed -Verbose

        Get file paths and detailed information about a function or alias in a module

        VERBOSE: PSBoundParameters: 
	
        Key              Value             
        ---              -----             
        CommandName      {Get-Yarn}
        Detailed         True              
        Verbose          True              
        Module                             
        ScriptName       Get-PKFunctionPath
        ScriptVersion    1.0.0             
        PipelineInput    False             
        ParameterSetName ByCommand         

        VERBOSE: [BEGIN: Get-PKFunctionPath] Return underlying file path for PowerShell commands
        VERBOSE: [Get-Yarn] Get command object
        VERBOSE: [Get-Yarn] Found command object 'Get-Yarn'
        VERBOSE: [Get-Yarn] Command object is a function
        VERBOSE: [Get-Yarn] Get parent module object
        VERBOSE: [Get-Yarn] Found parent module 'kittens'
        VERBOSE: [Get-Yarn] Get all ps* files in module 'kittens'
        VERBOSE: [Get-Yarn] Search 12 file(s) for '^Function Get-Yarn'
        VERBOSE: [Get-Yarn] Found 1 match(es) in 'C:\MyModules\kittens\Scripts\Function_Get-Yarn.ps1'

        Input         : Get-Yarn
        IsFound       : True
        InputType     : Function
        Function      : Get-Yarn
        Module        : kittens
        ModuleVersion : 1.18.0
        ModulePath    : C:\MyModules\kittens\kittens.psd1
        Filepath      : C:\MyModules\kittens\Scripts\Function_Get-Yarn.ps1
        LineNumber    : 2
        Messages      : Found match for pattern '^Function Get-Yarn\s+\{' in 1/12 file(s)

        VERBOSE: [END: Get-PKFunctionPath] Return underlying file path for PowerShell commands

.EXAMPLE 
    PS C:\> Get-PKFunctionPath -CommandName Reset-MyModules -Verbose

        Get file paths for a function or alias in a module

        VERBOSE: PSBoundParameters: 
	
        Key              Value                                 
        ---              -----                                 
        CommandName      {Reset-MyModules}
        Verbose          True                                  
        Module                                                 
        Detailed         False                                 
        ScriptName       Get-PKFunctionPath                    
        ScriptVersion    1.0.0                                 
        PipelineInput    False                                 
        ParameterSetName ByCommand                             

        VERBOSE: [BEGIN: Get-PKFunctionPath] Return underlying file names/paths for PowerShell commands
        VERBOSE: [Reset-MyModules] Get command object
        VERBOSE: [Reset-MyModules] Found command object 'Reset-MyModules'
        VERBOSE: [Reset-MyModules] Command object is a function
        VERBOSE: [Reset-MyModules] Get parent module object
        VERBOSE: [Reset-MyModules] Found parent module 'WinTools'
        VERBOSE: [Reset-MyModules] Get all ps* files in module 'WinTools'
        VERBOSE: [Reset-MyModules] Search 105 file(s) for '^Function Reset-MyModules'
        VERBOSE: [Reset-MyModules] Found 2 match(es) in 'C:\Temp\WinTools\Scripts\Function_Reset-MyModules.ps1', 'C:\Temp\WinTools\SampleCode.ps1'

        Function                Module    Filepath                                                                           
        --------                ------    --------                                                                           
        Reset-MyModules         WinTools  {C:\Temp\WinTools\Scripts\Function_Reset-MyModules.ps1,C:\Temp\WinTools\SampleCode.ps1}

        VERBOSE: [END: Get-PKFunctionPath] Return underlying file names/paths for PowerShell commands

.EXAMPLE
    PS C:\> Get-PKFunctionPath -CommandName Get-FilteredUsers,Get-ManagerDetails -Detailed

        Get file paths and detailed information about two functions or aliases in a module

        Input         : Get-FilteredUsers
        IsFound       : True
        InputType     : Function
        Function      : Get-FilteredUsers
        Module        : ADTools
        ModuleVersion : 1.19.0
        ModulePath    : D:\Modules\Custom\ADTools\ADTools.psd1
        Filepath      : D:\Modules\Custom\ADTools\scripts\Function_Get-FilteredUsers.ps1
        LineNumber    : 3
        Messages      : Found match for pattern '^[ \t]*Function[ \t]+Get-FilteredUsers[ \t]*(\(.+\))?[ \t]*\{[ \t]*$' in 1/28 module file(s)

        Input         : Get-ManagerDetails
        IsFound       : True
        InputType     : Function
        Function      : Get-ManagerDetails
        Module        : ADTools
        ModuleVersion : 1.19.0
        ModulePath    : D:\Modules\Custom\ADTools\ADTools.psd1
        Filepath      : D:\Modules\Custom\ADTools\scripts\Function_Get-ManagerDetails.ps1
        LineNumber    : 3
        Messages      : Found match for pattern '^[ \t]*Function[ \t]+Get-ManagerDetails[ \t]*(\(.+\))?[ \t]*\{[ \t]*$' in 1/28 module file(s)

#>
    [CmdletBinding(DefaultParameterSetName = "ByCommand")]
    Param(
        [Parameter(
            Position = 0,
            ParameterSetName = "ByCommand",
            Mandatory,
            HelpMessage = "One or more PowerShell function or alias names"
        )]
        [string[]]$CommandName,
        
        [Parameter(
            ParameterSetName = "ByModule",
            Mandatory,
            HelpMessage = "One or more PowerShell modules names or objects"
        )]
        [object[]]$Module,

        [Parameter(
            HelpMessage = "Return detailed output, including line numbers, module version, and pattern searched for"
        )]
        [switch]$Detailed
    )

Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$Scriptname)
    $CurrentParams.Add("ScriptVersion",$Version)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ParameterSetName",$Source)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | Out-String )"

    $Activity = Switch ($PSCmdlet.ParameterSetName) {
        ByCommand {"Return underlying file names/paths for PowerShell commands"}
        ByModule {"Return underlying file names/paths for PowerShell commands within modules"}
    }
    
    # Inner function to return the script name (filename) of the code which invoked this function.
    function Get-ScriptName {
        [PSCustomObject]@{
            ScriptName = $MyInvocation.ScriptName | Split-Path -Leaf
            LineNumber = $MyInvocation.ScriptLineNumber
        }
    }
    
    # Inner function to look up a command and resolve aliases to functions
    Function Get-UnderlyingFunction {
        [CmdletBinding()]
        Param(
            [Parameter(ValueFromPipeline)][string[]]$Command
        )
        Begin {}
        Process {
            Foreach ($C in $Command) {
                $CmdObject = Get-Command -Name $C -ErrorAction Stop
                Switch ($CmdObject.CommandType) {
                    Alias {
                        Write-Verbose "Command '$($CmdObject.Name)' is an alias to '$($CmdObject.ResolvedCommandName)'"
                        Get-UnderlyingFunction $CmdObject.ResolvedCommandName -Verbose
                    }
                    Function {$CmdObject}
                    Default {Write-Warning "Invalid command type '$($CmdObject.CommandType)'; please enter a function or alias"}
                }
            }
        }
    }

    $Msg = "[BEGIN: $ScriptName] $Activity"
    Write-Verbose $Msg

}
Process {

    Switch ($Source) {

        ByModule {
            
            If ($Detailed.IsPresent) {$Select = "Module,ModuleVersion,ModulePath,Command,Type,Filepath,LineNumber,Messages" -split(",")}
            Else {$Select = "Module,Command,Filepath" -split(",")}
            
            $Results = @()
            $ModuleObjects = @()

            $Module = ($Module | Select-Object -Unique)
            $TotalModules = $Module.Count
            $CurrentModule = 0

            Foreach ($Name in $Module) {
                
                $CurrentModule ++
                $Msg = "Get the module object"
                Write-Verbose "[$Name] $Msg"
                Write-Progress -Id 1 -Activity $Activity -CurrentOperation $Msg -Status $Name -PercentComplete ($CurrentModule/$TotalModules*100)

                Try {
                    $ModuleObj = Get-Module $Name -ListAvailable -ErrorAction Stop -Verbose:$False

                    Try {
                        $Msg = "Get all commands in module"
                        Write-Verbose "[$Name] $Msg"
                        $ModuleCommands = Get-Command -Module $ModuleObj.Name -ErrorAction Stop -Verbose:$False

                        If (-not $ModuleCommands) {
                            $Msg = "No exported commands found in module"
                            $Results += [PSCustomObject]@{
                                Module        = $ModuleObj.Name
                                ModuleVersion = $ModuleObj.Version.ToString()
                                ModulePath    = $ModuleObj.Path
                                Command       = "ERROR"
                                Type          = "ERROR"
                                Filepath      = "ERROR"
                                LineNumber    = "ERROR"
                                Messages      = $Msg
                            }
                        }
                        If ($ModuleCommands) {

                            # We will also check for aliases
                            If ([object[]]$Aliases = (Get-Alias | Where-Object {$_.Source -eq $ModuleObj.Name})) {
                                Foreach ($Alias in $Aliases) {
                                    $Msg = "Alias '$($Alias.Name)' points to underlying function '$($Alias.ResolvedCommand)'" 
                                    Write-Warning "[$Name] $Msg"
                                    $Msg = "Alias points to underlying function '$($Alias.ResolvedCommand)'" 
                                    $Results += [PSCustomObject]@{
                                        Module        = $ModuleObj.Name
                                        ModuleVersion = $ModuleObj.Version.ToString()
                                        ModulePath    = $ModuleObj.Path
                                        Command       = $Alias.Name
                                        Type          = "Alias"
                                        Filepath      = "-"
                                        LineNumber    = "-"
                                        Messages      = $Msg
                                    }
                                } 
                            } #end if aliases

                            $Msg = "Get functions in module"
                            Write-Verbose "[$Name] $Msg"
                            Write-Progress -Id 1 -Activity $Activity -CurrentOperation $Msg -Status $Name -PercentComplete ($CurrentModule/$TotalModules*100)
                            Try {
                            
                                If ($Functions = $ModuleCommands.Name | Get-UnderlyingFunction -Verbose:$False | Select-Object -Unique) {
                                    
                                    $TotalFunctions = $Functions.Count
                                    $CurrentFunction = 0  

                                    $Msg = "$Totalfunctions unique function name(s) found in module"
                                    Write-Verbose "[$Name] $Msg"

                                    $AllFiles = @()
                                    $Msg = "Get *.ps1 & *.psm1 files inside module"
                                    Write-Verbose "[$Name] $Msg"
                                    Write-Progress -Id 1 -Activity $Activity -CurrentOperation $Msg -Status $Name -PercentComplete ($CurrentModule/$TotalModules*100)

                                    Try {
                                        [object[]]$AllFiles = ($ModuleObj.Path | Split-Path -Parent) | Get-ChildItem -File -Recurse -Include *.ps1,*.psm1 -ErrorAction Stop
                                
                                        $Msg = "$($AllFiles.Count) *.ps1 & *.psm1 file(s) found in module path"
                                        Write-Verbose "[$Name] $Msg"
                    
                                        $Msg = "Search files for function names"
                                        Write-Verbose "[$Name] $Msg"
                                        Write-Progress -Id 1 -Activity $Activity -CurrentOperation $Msg -Status $Name -PercentComplete ($CurrentModule/$TotalModules*100)

                                        Foreach ($Function in ($Functions | Sort Name)) { 
                        
                                            $CurrentFunction ++
                                            Write-Progress -id 2 -Activity $Activity -Status $Function.Name -CurrentOperation $Msg -PercentComplete ($CurrentFunction/$TotalFunctions * 100)

                                            # Regex is fun! TM
                                            $Pattern = "^[ \t]*Function[ \t]+$($Function.Name)[ \t]*(\(.+\))?[ \t]*\{[ \t]*$"

                                            Try {
                                                If ($FoundIt = ($AllFiles | Sort-Object FullName | Select-String -pattern $Pattern -ErrorAction SilentlyContinue)) { 
                                                    
                                                    $Msg = "Found match in '$($FoundIt.Path -join("', '"))'"
                                                    Write-Verbose "[$Name] $($Function.Name): $Msg"

                                                    $Msg = "Pattern '$Pattern' found in $($FoundIt.Path.Count)/$($AllFiles.Count) module file(s)"
                                                    #Write-Verbose "[$Name] $Msg"
                                                    $Results += [PSCustomObject]@{
                                                        Module        = $ModuleObj.Name
                                                        ModuleVersion = $ModuleObj.Version.ToString()
                                                        ModulePath    = $ModuleObj.Path
                                                        Command       = $Function.Name
                                                        Type          = $Function.CommandType
                                                        Filepath      = $FoundIt.Path
                                                        LineNumber    = $FoundIt.LineNumber
                                                        Messages      = $Msg
                                                    }
                                                }
                                                Else {
                                                    $Msg = "No match found in $($AllFiles.Count) module file(s)"
                                                    Write-Warning "[$Name] $($Function.Name): $Msg"

                                                    $Msg = "Pattern '$Pattern' was not found in $($AllFiles.Count) module file(s)"
                                                    $Results += [PSCustomObject]@{
                                                        Module        = $ModuleObj.Name
                                                        ModuleVersion = $ModuleObj.Version.ToString()
                                                        ModulePath    = $ModuleObj.Path
                                                        Command       = $Function.Name
                                                        Type          = $Function.CommandType
                                                        Filepath      = "ERROR"
                                                        LineNumber    = "ERROR"
                                                        Messages      = $Msg
                                                    }
                                                }
                                            }
                                            Catch {
                                                $Msg = "Error searching file content for '$Pattern'"
                                                If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                                                Write-Warning "[$Name] $($Function.Name): $Msg"
                                                $Results += [PSCustomObject]@{
                                                    Module        = $ModuleObj.Name
                                                    ModuleVersion = $ModuleObj.Version.ToString()
                                                    ModulePath    = $ModuleObj.Path
                                                    Command       = $Function.Name
                                                    Type          = $Function.CommandType
                                                    Filepath      = "ERROR"
                                                    LineNumber    = "ERROR"
                                                    Messages      = $Msg
                                                }
                                            }
                                        } # end for each function to match   
                                    } #end try to get child items
                                    Catch {
                                        $Msg = "Unable to get underlying files in '$($ModuleObj.Path | Split-Path -Parent)'"
                                        If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                                        Write-Warning "[$Name] $Msg"
                                        $Results += [PSCustomObject]@{
                                            Module        = $ModuleObj.Name
                                            ModuleVersion = $ModuleObj.Version.ToString()
                                            ModulePath    = $ModuleObj.Path
                                            Command       = "ERROR"
                                            Type          = "ERROR"
                                            Filepath      = "ERROR"
                                            LineNumbers   = "ERROR"
                                            Messages      = $Msg
                                        }
                                    }
                                } #end if functions

                            } #end try getting underlying functions
                            Catch {}
                        } # end if module has commands to export
                    }
                    Catch {
                        $Msg = "Failed to get commands in module"
                        If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                        Write-Warning "[$Name] $Msg"
                        $Results += [PSCustomObject]@{
                            Module        = $Name
                            ModuleVersion = $ModuleObj.Version.ToString()
                            ModulePath    = $ModuleObj.Path
                            Command       = "ERROR"
                            Type          = "ERROR"
                            Filepath      = "ERROR"
                            LineNumber    = "ERROR"
                            Messages      = $Msg
                        }
                    }

                } #end try to get module object
                Catch {
                    $Msg = "No matching module object(s) found"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                    Write-Warning "[$Name] $Msg"
                    $Results += [PSCustomObject]@{
                        Module        = $Name
                        ModulePath    = "ERROR"
                        ModuleVersion = "ERROR"
                        Command       = "ERROR"
                        Type          = "ERROR"
                        Filepath      = "ERROR"
                        LineNumber    = "ERROR"
                        Messages      = $Msg
                    }
                }

            } #end foreach module name
            Write-Output ($Results | Select-Object $Select)
        }
        ByCommand {
            
            If ($Detailed.IsPresent) {$Select = "Input,IsFound,InputType,Function,Module,ModuleVersion,ModulePath,Filepath,LineNumber,Messages" -split(",")}
            Else {$Select = "Input,Function,Module,Filepath" -split(",")}

            $Results = @()
            $Errors = @()
            $Found = @()
            $AllFiles = @()   

            Foreach ($Name in $CommandName) {
                
                $Msg = "Get command object"
                Write-Verbose "[$Name] $Msg"
                
                # Get the underlying command object
                Try {
                    If ($CommandObj = Get-Command $Name -ErrorAction Stop) {
                        Switch ($CommandObj.CommandType) {
                            Alias{
                                $Msg = "This command is an alias resolving to '$($CommandObj.ResolvedCommandName)'"
                                Write-Warning "[$Name] $Msg"
                                If ($ResolvedAlias = (Get-Command $CommandObj.ResolvedCommandName -ErrorAction SilentlyContinue)) {
                                    $Msg = "Found target function '$($ResolvedAlias.Name)'"
                                    Write-Verbose "[$Name] $Msg"
                                    $Found += [PSCustomObject]@{
                                        Input         = $Name
                                        IsFound       = $False
                                        InputType     = "Alias"
                                        Function      = $ResolvedAlias.Name
                                        Module        = $ResolvedAlias.ModuleName
                                        ModuleVersion = "ERROR"
                                        ModulePath    = "ERROR"
                                        Filepath      = "ERROR"
                                        LineNumber    = "ERROR"
                                        Messages      = $Msg
                                    }
                                }
                                Else {
                                    $Msg = "Failed to locate underlying command object for '$($CommandObj.ResolvedCommandName)'"
                                    Write-Warning "[$Name] $Msg"
                                    $Results += [PSCustomObject]@{
                                        Input         = $Name
                                        IsFound       = $False
                                        InputType     = $CommandObj.CommandType
                                        Function      = $CommandObj.Name
                                        Module        = "ERROR"
                                        ModuleVersion = "ERROR"
                                        ModulePath    = "ERROR"
                                        Filepath      = "ERROR"
                                        LineNumber    = "ERROR"
                                        Messages      = $Msg
                                    }
                                }
                            }
                            Function {
                                $Msg = "Found target function"
                                Write-Verbose "[$Name] $Msg"
                                $Found += [PSCustomObject]@{
                                    Input         = $Name
                                    IsFound       = $False
                                    InputType     = $CommandObj.CommandType
                                    Function      = $CommandObj.Name
                                    Module        = $CommandObj.ModuleName
                                    ModulePath    = "ERROR"
                                    ModuleVersion = "ERROR"
                                    Filepath      = "ERROR"
                                    LineNumber    = "ERROR"
                                    Messages      = $Msg
                                }
                            }
                            Default {
                                $Msg = "Input must be an alias or a function"
                                Write-Warning "[$Name] $Msg"
                                $Results += [PSCustomObject]@{
                                    Input         = $Name
                                    IsFound       = $False
                                    InputType     = $CommandObj.CommandType
                                    Function      = $CommandObj.Name
                                    Module        = "ERROR"
                                    ModulePath    = "ERROR"
                                    ModuleVersion = "ERROR"
                                    Filepath      = "ERROR"
                                    LineNumber    = "ERROR"
                                    Messages      = $Msg
                                }
                            }
                        } #end switch

                    } #end if command object found
                }
                Catch {
                    $Msg = "Failed to find matching command object(s)"
                    Write-Warning "[$Name] $Msg"
                    $Results += [PSCustomObject]@{
                        Input         = $Name
                        IsFound       = $False
                        InputType     = "ERROR"
                        Function      = "ERROR"
                        Module        = "ERROR"
                        ModuleVersion = "ERROR"
                        ModulePath    = "ERROR"
                        Filepath      = "ERROR"
                        LineNumber    = "ERROR"
                        Messages      = $Msg
                    }
                }

            } #end foreach command name


            # We're doing this separately although yes yes it looks silly
            # Want to change this so we get all the unique modules FIRST, collected
            # in a psojbect with name, version, + files

            $CurrentFunction = 0
            $TotalFunctions = $Found.Count
            $Activity2 = "Get parent module object"
            Foreach ($Function in $Found) {
                
                $CurrentFunction ++

                If (-not $Function.Module) {
                    $Msg = "Filename search is available only for functions within PowerShell modules"
                    Write-Warning "[$($Function.Input)] $Msg"
                    $Function.Messages = $Msg
                    $Results += $Function
                }
                Else {
                    $Msg = "Get parent module object"
                    Write-Verbose "[$($Function.Input)] $Msg"
                    Write-Progress -id 2 -Activity $Activity2 -Status $Function.Input -CurrentOperation $Msg -PercentComplete ($CurrentFunction/$TotalFunctions * 100)

                    Try {
                        $ModuleObj = Get-Module $Function.Module -ListAvailable -ErrorAction Stop -Verbose:$False
                        $Msg = "Found parent module '$($ModuleObj.Name)'"
                        Write-Verbose "[$($Function.Input)] $Msg"

                        $Function.Module = $ModuleObj.Name
                        $Function.ModulePath = $ModuleObj.Path
                        $Function.ModuleVersion = $ModuleObj.Version.ToString()

                        $AllFiles = @()
                        $Msg = "Get all ps* files in module '$($ModuleObj.Name)'"
                        Write-Verbose "[$($Function.Input)] $Msg"
                        $Activity2 = "Search file content for pattern"
                        Try {
                            [object[]]$AllFiles = @()
                            $AllFiles = ($ModuleObj.Path | Split-Path -Parent) | Get-ChildItem -File -Recurse -Include *.ps1,*.psm1 -ErrorAction SilentlyContinue

                            $Msg = "Search $($AllFiles.Count) file(s) for '^Function $($Function.Function)'"
                            Write-Verbose "[$($Function.Input)] $Msg"
                            Write-Progress -id 2 -Activity $Activity2 -Status $Function.Input -CurrentOperation $Msg -PercentComplete ($CurrentFunction/$TotalFunctions * 100)

                            # Regex is fun! TM
                            $Pattern = "^[ \t]*Function[ \t]+$($Function.Function)[ \t]*(\(.+\))?[ \t]*\{[ \t]*$"

                            Try {
                                If ($FoundIt = ($AllFiles | Sort-Object FullName | Select-String -pattern $Pattern -ErrorAction SilentlyContinue)) { 
                                    
                                    $Msg = "Found $($Foundit.Count) match(es) in '$($FoundIt.Path)'"
                                    Write-Verbose "[$($Function.Input)] $Msg"
                                    $Msg = "Found match for pattern '$Pattern' in $($FoundIt.Path.Count)/$($AllFiles.Count) module file(s)"
                                    $Function.IsFound = $True
                                    $Function.FilePath = $FoundIt.Path
                                    $Function.LineNumber = $Foundit.LineNumber
                                    $Function.Messages = $Msg
                                    $Results += $Function
                                }
                                Else {
                                    $Msg = "No match found for pattern '^Function $($Function.Function)' in $($AllFiles.Count) module file(s)"
                                    Write-Warning "[$($Function.Input)] $Msg"
                                    $Function.IsFound = $False
                                    $Function.FilePath = $Null = $Function.LineNumber = $Null
                                    $Function.Messages = $Msg
                                    $Results += $Function

                                }
                            } # end try select-string
                            Catch {
                                $Msg = "Error searching file content"
                                If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                                Write-Warning "[$($Function.Input)] $Msg"
                                $Function.Messages = $Msg
                                $Results += $Function
                            }

                        } # end try get-childitem
                        Catch {
                            $Msg = "Failed to get child items in $($ModuleObj.Path | Split-Path -Parent)"
                            If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                            Write-Warning "[$($Function.Input)] $Msg"
                            $Function.Messages = $Msg
                            $Results += $Function
                        }
                    }
                    Catch {
                        $Msg = "Failed to get module object(s)"
                        If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                        Write-Warning "[$Function] $Msg"
                        $Function.Messages = $Msg
                        $Results += $Function
                    }
                } #end if function in module            
            }
                    
            Write-Output ($Results | Select-Object $Select)

        } # end switch ByCommand

    } # End for switch for parametersetname    
}
End {
    $Msg = "[END: $ScriptName] $Activity"
    Write-Verbose $Msg
    $Null = Write-Progress -Activity * -Completed
}
} #end Get-PKFunctionPath
