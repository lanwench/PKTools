Function Test-PKADSite {
[CmdletBinding ()]
Param($SiteName,$Forest)
    Try {
        $Msg = "Get all sites in forest '$Forest'"
        Write-Verbose $Msg
        $ForestObj = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest", $Forest)
        [array]$ADSites = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ForestObj).sites
            
        $Msg = "Validate site name '$SiteName'"
        Write-Verbose $Msg
        If ($ADSites.Name -contains $SiteName) {$True}
        Else {$False}
    }
    Catch {$Host.UI.WriteErrorLine($_.Exception.Message)}
}



Function GetSites {
    [CmdletBinding ()]
    Param($Forest = $ADConfirm.Forest,$Credential)
                
        If ($Credential) {
            Try {
                $ForestObj = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest", $Forest,$Credential.UserName,$Credential.GetNetworkCredential().Password)
            }
            Catch {
                $Msg = "Failed to connect to forest '$Forest' as '$($Credential.UserName)'"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
                $Host.UI.WriteErrorLine("ERROR  : [Prerequisites] $Msg")
            }
        }
        Else {
            Try {
                $ForestObj = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest", $Forest)
            }
            Catch {
                $Msg = "Failed to connect to forest '$Forest' as '$Env:UserDomain\$env:USERNAME'"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
                $Host.UI.WriteErrorLine("ERROR  : [Prerequisites] $Msg")
            }
        }
                    
        If ($ForestObj) {
            Try {
                [array]$ADSites = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ForestObj).sites
                Write-Output ($ADSites.Name | Sort)
            }
            Catch {
                $Msg = "Failed to get sites in forest '$Forest'"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
                $Host.UI.WriteErrorLine("ERROR  : [Prerequisites] $Msg")
            }
        }
    }