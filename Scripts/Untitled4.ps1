$Computer = "ops-pktest-1"


$Computer = "ops-sgtemp-1"


$Computer = "scdc03"


$Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher -ErrorVariable ErrProcessNewObjectSearcher @StdParams
$Searcher.SizeLimit = $SizeLimit
$searcher.PropertiesToLoad.AddRange(('name','canonicalname','serviceprincipalname','useraccountcontrol','lastlogonTimeStamp','distinguishedname','description','operatingsystem','location','whencreated','dnshostname'))

$Searcher.SearchRoot = $DomainDN
If ($DomainDN -notmatch "LDAP://") {$Searcher.SearchRoot = "LDAP://$DomainDN"}
        
$DomainDNSRoot = ($DomainDN -replace('DC=',$Null) -replace(',','.')).split('/')[-1]

$Searcher.Filter = "(&(objectCategory=Computer)(name=$Computer))"

$Obj = $Searcher.FindAll()

$Obj.Properties