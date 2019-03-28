function Get-NestedGroupMember{
<#

.LINK
    http://community.idera.com/powershell/powertips/b/tips/posts/finding-nested-active-directory-memberships-part-2
#>
param(
    [Parameter(Mandatory,ValueFromPipeline)]
    [string]
    $distinguishedName
)
 
process{
        
    $DomainController = $env:logonserver.Replace("\\","")
    $Domain = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DomainController")
    $Searcher = New-Object System.DirectoryServices.DirectorySearcher($Domain)
    $Searcher.PageSize = 1000
    $Searcher.SearchScope = "subtree"
    $Searcher.Filter = "(&(objectClass=group)(member:1.2.840.113556.1.4.1941:=$distinguishedName))"
    # attention: property names are case-sensitive!
    $colProplist = "name","distinguishedName"
    foreach ($i in $colPropList){$Searcher.PropertiesToLoad.Add($i) | Out-Null}
    $all = $Searcher.FindAll()
 
    $all.Properties | ForEach-Object {
      [PSCustomObject]@{
        # attention: property names on right side are case-sensitive!
        Name = $_.name[0]
        DN = $_.distinguishedname[0]
    } }
  }
}
 
# make sure you specify a valid distinguishedname for a user below 
#Get-NestedGroupMember -distinguishedName 'CN=Joe Bloggs,OU=Admins,DC=domain,DC=local'




<# 
# If you have the ActiveDirectory module
function Get-NestedGroupMember
{
    param
    (
        [Parameter(Mandatory,ValueFromPipeline)]
        [string]
        $Identity
    )
 
    process
    {
        $user = Get-ADUser -Identity $Identity
        $userdn = $user.DistinguishedName
        $strFilter = "(member:1.2.840.113556.1.4.1941:=$userdn)"
        Get-ADGroup -LDAPFilter $strFilter -ResultPageSize 1000
    }
}
 
 
Get-NestedGroupMember -Identity $env:username |
  Select-Object -Property Name, DistinguishedName

  #>