
function Test-Department{
<#
.SYNOPSIS
    Demonstration of dynamic dynamic parameters

.DESCRIPTION
    Demonstration of dynamic dynamic parameters

.LINK
    http://community.idera.com/database-tools/powershell/powertips/b/tips/posts/using-dynamic-parameters

#>
[CmdletBinding()]

param(

    [Parameter(Mandatory=$true)]
    [ValidateSet('Microsoft','Amazon','Google','Facebook')]
    $Company

)
 
dynamicparam {

    # this hashtable defines the departments available in each company
    $data = @{
        Microsoft = 'CEO', 'Marketing', 'Delivery'
        Google = 'Marketing', 'Delivery'
        Amazon = 'CEO', 'IT', 'Carpool'
        Facebook = 'CEO', 'Facility', 'Carpool'
    }


    # check to see whether the user already chose a company
    if ($Company){

        # yes, so create a new dynamic parameter
        $paramDictionary = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameterDictionary
        $attributeCollection = New-Object -TypeName System.Collections.ObjectModel.Collection[System.Attribute]

        # define the parameter attribute
        $attribute = New-Object System.Management.Automation.ParameterAttribute
        $attribute.Mandatory = $false
        $attributeCollection.Add($attribute)

        # create the appropriate ValidateSet attribute, listing the legal values for
        # this dynamic parameter

        $attribute = New-Object System.Management.Automation.ValidateSetAttribute($data.$Company)
        $attributeCollection.Add($attribute)

        # compose the dynamic -Department parameter
        $Name = 'Department'
        $dynParam = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameter($Name,[string], $attributeCollection)
        $paramDictionary.Add($Name, $dynParam)

        # return the collection of dynamic parameters
        $paramDictionary
        
    } #end if company

} #end dynamic parameter

end{

    # take the dynamic parameters from $PSBoundParameters
    $Department = $PSBoundParameters.Department
    "Chosen department for $Company : $Department"
}

}
          