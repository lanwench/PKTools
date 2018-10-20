
Function Get-PKFullUsername {
[CmdletBinding()]
Param([switch]$CurrentOnly)
Process {

    If ($CurrentOnly.IsPresent) {
        $Msg = "Get current user's name on '$Env:ComputerName'"
    }
    Else {
        $Msg = "Get user names on '$Env:ComputerName'"
    }
    Write-Verbose $Msg
    
    Try {
        
        [array]$Results = @()

        # Get all user profiles
        $SIDLocalUsers = Get-WmiObject Win32_UserProfile -EA Stop | select-Object Localpath,SID
        
        If ($CurrentOnly.IsPresent) {
            
            $UserName = (Get-WMIObject -class Win32_ComputerSystem -Property UserName -ErrorAction Stop).UserName
            $UserOnly = $UserName.Split("\")[1]

            $Msg = $UserName
            Write-Verbose $Msg

            Foreach ($Profile in $SIDLocalUsers) {
            
                # Match profile to current user
                If ($Profile.localpath -like "*$UserOnly"){

                    # Look up path in registry/Group Policy by SID
                    $SID = $Profile.sid
                    $DN = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\$SID" -ErrorAction SilentlyContinue | Select -ExpandProperty Distinguished-Name)
                    
                    If ($DN -match "(CN=)(.*?),.*") {
                        $FullName = $Matches[2]
                        $Output = New-Object PSObject -Property ([ordered] @{
                            ComputerName      = $Env:ComputerName
                            Username          = $UserName
                            FullName          = $FullName
                            ProfilePath       = $Profile.LocalPath
                            SID               = $SID
                            DistinguishedName = $($DN)
                            Messages          = $Null
                        })
                        $Results += $Output
                    }
                    <#
                    Else {
                        $Msg = "Failed to find registry match for username '$UserName'"
                        $Host.UI.WriteErrorLine("ERROR: $Msg")

                        $Output = New-Object PSObject -Property ([ordered] @{
                            ComputerName      = $Env:ComputerName
                            Username          = $UserName
                            FullName          = "Error"
                            ProfilePath       = $Profile.LocalPath
                            SID               = $SID
                            DistinguishedName = "Error"
                            Messages          = $Msg
                        })
                    }
                    #> 
                } #end if found in profile path
            } #end for each profile

            If ($Results.Count -eq 0) {
                $Msg = "No fullname data found"
                $Host.UI.WriteErrorLine("ERROR: $Msg")
                $Output = New-Object PSObject -Property ([ordered] @{
                    ComputerName      = $Env:ComputerName
                    Username          = $UserName
                    FullName          = $FullName
                    ProfilePath       = $Profile.LocalPath
                    SID               = $SID
                    DistinguishedName = $($DN)
                    Messages          = $Msg
                })
                $Results += $Output
            }

        } #end if current user only

        Else {
            
            Foreach ($Profile in $SIDLocalUsers) {
                
                $Msg = $Profile.LocalPath
                Write-Verbose $Msg

                # Look up path in registry/Group Policy by SID
                $SID = $Profile.sid
                $DN = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\$SID" -ErrorAction SilentlyContinue | Select -ExpandProperty Distinguished-Name)
                
                # Look up the username by SID
                $Username = (Get-WMIObject -Class Win32_UserAccount -Filter "sid = '$SID'" -ErrorAction Stop).name
                    
                If ($DN -match "(CN=)(.*?),.*") {
                    $FullName = $Matches[2]
                    $Output = New-Object PSObject -Property ([ordered] @{
                        ComputerName      = $Env:ComputerName
                        Username          = $UserName
                        FullName          = $FullName
                        ProfilePath       = $Profile.LocalPath
                        SID               = $SID
                        DistinguishedName = $($DN)
                        Messages          = $Null
                    })
                    $Results += $Output
                }
                Else {
                    $Msg = "Failed to find user registry match for '$($Profile.LocalPath)'"
                    $Host.UI.WriteErrorLine("ERROR: $Msg")

                    $Output = New-Object PSObject -Property ([ordered] @{
                        ComputerName      = $Env:ComputerName
                        Username          = "Error"
                        FullName          = "Error"
                        ProfilePath       = $Profile.LocalPath
                        SID               = "Error"
                        DistinguishedName = "Error"
                        Messages          = $Msg
                    })
                    $Results += $Output
                } 
            } #end if found in profile path
        } #end if geting for all profiles
        
    }
    Catch {
        
        $Msg = "Operation failed"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
        $Host.UI.WriteErrorLine("ERROR: $Msg")
        
        $Output = New-Object PSObject -Property ([ordered] @{
            ComputerName      = $Env:ComputerName
            Username          = $UserName
            FullName          = "Error"
            ProfilePath       = $Profile.LocalPath
            DistinguishedName = "Error"
            Messages          = $Msg
        })
        $Results += $Output
    }

    Write-Output $Results
}

}