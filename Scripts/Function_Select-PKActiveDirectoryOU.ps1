#requires -Version 3
Function Select-PKActiveDirectoryOU{
<#
.SYNOPSIS
    Uses Windows Forms to generate a selectable tree view of Active Directory containers/organizational units

.DESCRIPTION
    Uses Windows Forms to generate a selectable tree view of Active Directory containers/organizational units
    Allows for selection by Active Directory forest or domain (defaults to current user's)
    Returns a string

.NOTES
    Name    : Function_Select-PKActiveDirectoryOU.ps1
    Author  : Paula Kingsley
    Created : 2019-03-04
    Version : 01.00.0000
    History :

        *** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00.0000 - 2019-03-04 - Created script based on Rene Horn's original 

.PARAMETER ADDomain
    Active Directory domain (default is current user's)

.PARAMETER ADForest
    Active Directory forest (default is current user's)

.PARAMETER Server
    Active Directory domain controller (default is first available)

.PARAMETER Credential
    Valid AD credential (default is current user)

.PARAMETER SuppressConsoleOutput
    Suppress non-verbose/non-error console output        

.LINK
    https://gist.github.com/supercheetah/b68023f3254dfc9a6497

.LINK
    https://itmicah.wordpress.com/2013/10/29/active-directory-ou-picker-in-powershell/

.EXAMPLE
    PS C:\> Select-PKActiveDirectoryOU -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                   
        ---                   -----                   
        Verbose               True                    
        ADDomain              domain.local
        ADForest                                      
        Server                                        
        Credential                                    
        SuppressConsoleOutput False                   
        Searchbase                                    
        ParameterSetName      Domain                  
        ScriptName            Select-PKActiveDirectoryOU
        ScriptVersion         1.0.0                   

        BEGIN  : Browse Active Directory forest for container/Organizational Unit selection

        VERBOSE: Get Active Directory domain object
        VERBOSE: Found Active Directory domain 'domain'
        VERBOSE: Get Active Directory forest object
        VERBOSE: Found forest object 'parentco.local'
        VERBOSE: Query AD forest for domains and hostnames
        VERBOSE: [parentco.local] domain.local
        VERBOSE: [parentco.local] [domain.local] OAKWINDOMP001.domain.local
        VERBOSE: [parentco.local] acquisition.net
        WARNING: [parentco.local] Failed to find domain controller for 'acquisition.net'
        VERBOSE: Please make a selection from the dialog box (may not be in front)
        
        OU=Exchange Admin Groups,OU=Global,DC=domain,DC=org

        END    : Script complete

.EXAMPLE    
    PS C:\> Get-ADDomain -Identity newcorp.lan | Select-PKActiveDirectoryOU -Server dc04.newcorp.lan -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                         
        ---                   -----                         
        Server                dc04.newcorp.lan
        Verbose               True                          
        ADDomain              domain.local      
        ADForest                                            
        Credential                                          
        SuppressConsoleOutput False                         
        Searchbase                                          
        ParameterSetName      Domain                        
        ScriptName            Select-PKActiveDirectoryOU      
        ScriptVersion         1.0.0                         

        BEGIN  : Browse Active Directory forest for container/Organizational Unit selection

        VERBOSE: Get Active Directory domain object
        VERBOSE: Found Active Directory domain 'newcorp'
        VERBOSE: Get Active Directory forest object
        VERBOSE: Found forest object 'newcorp.lan'
        VERBOSE: Query AD forest for domains and hostnames
        VERBOSE: [newcorp.lan] newcorp.lan
        VERBOSE: [newcorp.lan] [DC=newcorp,DC=lan] dc03.newcorp.lan
        VERBOSE: Please make a selection from the dialog box (may not be in front)

        OU=Security Groups,OU=All Groups,OU=newcorp,OU=Production,DC=newcorp,DC=lan
        
        END    : Script complete


#>
[Cmdletbinding(
    DefaultParameterSetName = "Domain"
)]
Param(
    
    [Parameter(
        ParameterSetName = "Domain",
        HelpMessage = "Active Directory domain (default is current user's)",
        ValueFromPipeline = $True
    )]
    [ValidateNotNullOrEmpty()]
    [string] $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name,

    [Parameter(
        ParameterSetName = "Forest",
        HelpMessage = "Active Directory forest (default is current user's)"
        #ValueFromPipeline = $True
    )]
    [ValidateNotNullOrEmpty()]
    [string] $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Name,

    [Parameter(
        HelpMessage = "Active Directory domain controller (default is first available)"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$Server,

    [Parameter(
        HelpMessage = "Valid AD credential (default is current user)"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential] $Credential,

    [Parameter(
        HelpMessage = "Suppress non-verbose/non-error console output"
    )]
    [switch] $SuppressConsoleOutput
)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Parameterset
    $Source = $PSCmdlet.ParameterSetName
    
    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Searchbase = $Searchbase
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Switch ($Source) {
        Forest {$CurrentParams.ADDomain = $Null}
        Domain {$CurrentParams.ADForest = $Null}
    }
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preference
    $ErrorActionPreference = "Stop"
    $ProgressPreference = "Continue"

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    #region Test for AD module

    If (-not ($Null = Get-Module ActiveDirectory -ListAvailable -ErrorAction SilentlyContinue)) {
        $Msg = "This function requires the ActiveDirectory module"
        $Host.UI.WriteErrorLine("ERROR  : $Msg")
        Break
    }

    #endregion Test for AD module

    # Load assemblies
    [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

    # Hashtable for Domain Contollers
    $dc_hash = @{}

    #region Inner functions

    # Function to get node info
    function Get-NodeInfo($sender, $dn_textbox){
        
        $selected_node = $sender.Node
        $dn_textbox.Text = $selected_node.Name

    } #end Get-NodeInfo

    # Function to get OU
    function Add-ChildNodes($sender){
        [cmdletbinding()]
        $ProgressPreference = "Continue"
        

        $Param = @{}
        $Param = @{
            Server      = $Null
            Filter      = 'ObjectClass -eq "organizationalUnit" -or ObjectClass -eq "container"'
            SearchBase  = $Null
            SearchScope = "OneLevel"
            ErrorAction = "SilentlyContinue"
        }
        If ($Credential) {    
            $Param.Add("Credential",$Credential)
        }

        $expanded_node = $sender.Node
        if ($expanded_node.Name -eq "root") {
            return
        }

        $expanded_node.Nodes.Clear() | Out-Null
        $dc_hostname = $dc_hash[$($expanded_node.Name -replace '(OU=[^,]+,)*((DC=\w+,?)+)','$2')]

        $Msg = "Get Active Directory containers and organizational units"
        Write-Progress -Activity $Msg -Id 1

        $Param.Server = $dc_hostname
        $Param.SearchBase = $expanded_node.Name
        #$child_OUs = Get-ADObject -Server $dc_hostname -Filter 'ObjectClass -eq "organizationalUnit" -or ObjectClass -eq "container"' -SearchScope OneLevel -SearchBase $expanded_node.Name
        #$child_OUs = Get-ADObject -Server $dc_hostname -Filter 'ObjectClass -eq "organizationalUnit" -or ObjectClass -eq "container"' -SearchScope OneLevel -SearchBase $expanded_node.Name
        $child_OUs = Get-ADObject @Param
        
        if($child_OUs -eq $null) {
            $sender.Cancel = $true
        } else {
            foreach($ou in $child_OUs) {
                $ou_node = New-Object Windows.Forms.TreeNode
                $ou_node.Text = $ou.Name
                $ou_node.Name = $ou.DistinguishedName
                $ou_node.Nodes.Add('') | Out-Null
                $expanded_node.Nodes.Add($ou_node) | Out-Null
            }
        }
        Write-Progress -Activity $Msg -Completed
    } #end Add-ChildNodes

    # Function to get domain controllers
    function Add-ForestNodes{
        [Cmdletbinding()]
        Param (
            $ForestObj, 
            [ref]$dc_hash,
            $Credential
        )
        $Msg = "Query AD forest for domains and hostnames"
        $ProgressPreference = "Continue"
        Write-Verbose $Msg
        $Activity = $Msg
        
        $Param = @{}
        $Param = @{
            Server      = $Null
            ErrorAction = "SilentlyContinue"
        }
        If ($Credential) {    
            $Param.Add("Credential",$Credential)
        }

        $ad_root_node = New-Object Windows.Forms.TreeNode
        $ad_root_node.Text = $ForestObj.RootDomain
        $ad_root_node.Name = "root"
        $ad_root_node.Expand()

        $Current = 1
        $Total = $ForestObj.Domains.Count
        foreach ($ad_domain in $ForestObj.Domains) {
            $Msg = "[$($ForestObj.RootDomain)] $AD_Domain"
            Write-Verbose "$Msg"
            Write-Progress -Activity $Activity -Status $ad_domain -PercentComplete ($Current++ / $Total * 100) -Id 1
            Try {
                $Param.Server = $ad_domain
                $dc = Get-ADDomainController @Param
                
                $Msg = "[$($ForestObj.RootDomain)] [$ADDomain] $($DC.HostName)"
                Write-Verbose $Msg
                $dn = $dc.DefaultPartition
                $dc_hash.Value.Add($dn, $dc.Hostname)
                $dc_node = New-Object Windows.Forms.TreeNode
                $dc_node.Name = $dn
                $dc_node.Text = $dc.Domain
                $dc_node.Nodes.Add("") | Out-Null
                $ad_root_node.Nodes.Add($dc_node) | Out-Null
            }
            Catch {
                $Msg = "[$($ForestObj.RootDomain)] Failed to find domain controller for '$ad_domain'"
                Write-Warning $Msg
            }
        } 
        Write-Progress -Activity $Activity -Completed
        return $ad_root_node

    } #end Add-ForestNodes

    #endregion Inner functions

    # Console output
    $Activity = "Browse Active Directory forest for container/Organizational Unit selection"
    $BGColor = $Host.UI.RawUI.BackgroundColor
    $Msg = "BEGIN  : $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg`n")}
    Else {Write-Verbose $Msg}

}
Process {
    
    # Initialize objects
    $Selection    = $null
    $ForestObj    = $Null
    $DomainObj    = $Null
    $main_dlg_box = $Null
    $dn_text_box  = $Null

    # Get the AD forest object
    # Can't figure out how to use .NET to get the identical properties returned using either Get-ADForest
    # or [System.DirectoryServices.ActiveDirectory.Forest]

    Try {
        Switch ($Source) {
            Domain {

                $Msg = "Get Active Directory domain object"
                Write-Verbose $Msg
                Write-Progress -Activity $Activity -CurrentOperation $Msg

                # Get the domain
                $Param_GetDomainObj = @{}
                $Param_GetDomainObj = @{
                    Identity    = $ADDomain
                    ErrorAction = "SilentlyContinue"
                    Verbose     = $False
                }
                If ($Credential) {
                    $Param_GetDomainObj.Add("Credential",$Credential)
                }
                If ($CurrentParams.Server) {
                    $Param_GetDomainObj.Add("Server",$Server)
                }
                If ($DomainObj = Get-ADDomain @Param_GetDomainObj) {
                    
                    $Msg = "Found Active Directory domain '$($DomainObj.Name)'"
                    Write-Verbose $Msg

                    $Msg = "Get Active Directory forest object"
                    Write-Verbose $Msg
                    Write-Progress -Activity $Activity -CurrentOperation $Msg

                    # Get the forest for this domain
                    $Param_GetForestObj = @{}
                    $Param_GetForestObj = @{
                        ErrorAction = "Stop"
                        Verbose     = $False
                    }
                    If ($CurrentParams.Credential) {
                        $Param_GetForestObj.Add("Credential",$Credential)
                    }
                    If ($CurrentParams.Server) {
                        $Param_GetforestObj.Add("Server",$Server)
                    }
                    $ForestObj = $DomainObj | Get-ADForest @Param_GetForestObj
                }
            }
            Forest {
                
                $Msg = "Get Active Directory forest object"
                Write-Verbose $Msg
                Write-Progress -Activity $Activity -CurrentOperation $Msg

                # Get the forest
                $Param_GetForestObj = @{}
                $Param_GetForestObj = @{
                    ErrorAction = "SilentlyContinue"
                    Verbose     = $False
                }
                If ($Credential) {
                    $Param_GetForestObj.Add("Credential",$Credential)
                }
                If ($CurrentParams.Server) {
                    $Param_GetDomainObj.Add("Server",$Server)
                }
                # Can't seem to search by name ..
                $ForestObj = Get-ADForest @Param_GetForestObj | Where-Object {$_.Name -eq $ADForest}
            }
        }
    }
    Catch {
        $Msg = "Failed to get Active Directory forest object"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
        $Host.UI.WriteErrorLine("ERROR  : $Msg")
    }
    
    If ($ForestObj) {

        $Msg = "Found forest object '$($ForestObj.Name)'"
        Write-Verbose $Msg

        # Create main dialog box
        $main_dlg_box = New-Object System.Windows.Forms.Form
        $main_dlg_box.ClientSize = New-Object System.Drawing.Size(420,600)
        $main_dlg_box.MaximizeBox = $false
        $main_dlg_box.MinimizeBox = $false
        $main_dlg_box.FormBorderStyle = 'FixedSingle'

        # Add title
        $Title = "Please make a selection"
        $main_dlg_box.Text = $Title

        # Make sure box is centered & always on top
        $main_dlg_box.StartPosition = "CenterParent" # CenterScreen, Manual, WindowsDefaultLocation, WindowsDefaultBounds, CenterParent
        $main_dlg_box.TopMost = $True
        
        # Set opacity
        $main_dlg_box.Opacity = 8.5  # 1.0 is fully opaque; 0.0 is invisible

        # Set size/scroll - leaving in for doc, but don't use
        #$main_dlg_box.AutoScroll = $True
        #$main_dlg_box.AutoSize = $True
        #$main_dlg_box.AutoSizeMode = "GrowAndShrink"

        # widget size and location variables
        $ctrl_width_col = $main_dlg_box.ClientSize.Width/20
        $ctrl_height_row = $main_dlg_box.ClientSize.Height/15
        $max_ctrl_width = $main_dlg_box.ClientSize.Width - $ctrl_width_col*2
        $max_ctrl_height = $main_dlg_box.ClientSize.Height - $ctrl_height_row
        $right_edge_x = $max_ctrl_width
        $left_edge_x = $ctrl_width_col
        $bottom_edge_y = $max_ctrl_height
        $top_edge_y = $ctrl_height_row

        # setup text box showing the distinguished name of the currently selected node
        $dn_text_box = New-Object System.Windows.Forms.TextBox
        # cannot set the height for a single line text box, that's controlled by the font being used
        $dn_text_box.Width = (14 * $ctrl_width_col)
        $dn_text_box.Location = New-Object System.Drawing.Point($left_edge_x, ($bottom_edge_y - $dn_text_box.Height))
        $main_dlg_box.Controls.Add($dn_text_box)
        # /text box for dN

        # setup OK button
        $ok_button = New-Object System.Windows.Forms.Button
        $ok_button.Size = New-Object System.Drawing.Size(($ctrl_width_col * 2), $dn_text_box.Height)
        $ok_button.Location = New-Object System.Drawing.Point(($right_edge_x - $ok_button.Width), ($bottom_edge_y - $ok_button.Height))
        $ok_button.Text = "OK"
        $ok_button.DialogResult = 'OK'
        $main_dlg_box.Controls.Add($ok_button)
        # /Ok button

        # setup Cancel button (the size/position are still wonky)
        $cancel_button = New-Object System.Windows.Forms.Button
        $cancel_button.Size = New-Object System.Drawing.Size(($ctrl_width_col * 2), $dn_text_box.Height)
        $cancel_button.Location = New-Object System.Drawing.Point((($right_edge_x + $ok_button.width) - ($cancel_button.Width)), (($bottom_edge_y) - ($cancel_button.Height)))
        $cancel_button.Text = "Nope"
        $cancel_button.DialogResult = 'Cancel'
        $main_dlg_box.Controls.Add($cancel_button)
        # /Cancel button

        # setup tree selector showing the domains
        $ad_tree_view = New-Object System.Windows.Forms.TreeView
        $ad_tree_view.Size = New-Object System.Drawing.Size($max_ctrl_width, ($max_ctrl_height - $dn_text_box.Height - $ctrl_height_row*1.5))
        $ad_tree_view.Location = New-Object System.Drawing.Point($left_edge_x, $top_edge_y)
        $ad_tree_view.Nodes.Add($(Add-ForestNodes $ForestObj ([ref]$dc_hash))) | Out-Null
        $ad_tree_view.Add_BeforeExpand({Add-ChildNodes $_})
        $ad_tree_view.Add_AfterSelect({Get-NodeInfo $_ $dn_text_box})
        $main_dlg_box.Controls.Add($ad_tree_view)
        # /tree selector

        # Display dialog box
        $Msg = "Please make a selection from the dialog box (may not be in front)"
        Write-Verbose $Msg
        Write-Progress -Activity $Activity -CurrentOperation $Msg

        $Select = $main_dlg_box.ShowDialog() #| Out-Null

        Switch ($Select) {
            OK     {
                If ($dn_text_box.Text) {
                    
                    $dn_text_box.Text
                }
                Else {
                    $Msg = "No selection made"
                    $FGColor = "Red"
                    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"`n$Msg")}
                    Else {Write-Verbose $Msg}
                }
            }
            Cancel {
                $Msg = "No selection made"
                $FGColor = "Red"
                If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"`n$Msg")}
                Else {Write-Verbose $Msg}
            }
        }
    }
    Else {
        $Msg = "Failed to get Active Directory forest object"
        $Host.UI.WriteErrorLine("ERROR  : $Msg")
    }
}

End {

    $Msg = "END    : Script complete"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"`n$Msg")}
    Else {Write-Verbose $Msg}

    Write-Progress -Activity * -Complete
    $main_dlg_box.Dispose()
    $dn_text_box.Dispose()
}
}