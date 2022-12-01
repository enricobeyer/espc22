#requires -Module Microsoft.Graph.Groups
function Search-MGGroup {
    param(
      $Name,
      #Only returns results that are Microsoft Teams
      [Switch]$Team
    )
    if ($Team) {
      $filter = "resourceProvisioningOptions/Any(x:x eq 'Team')"
    }
    if ($Name) {
      $search = "displayname:$Name"
    }
    Get-MgGroup -Filter $filter -Search $search -ConsistencyLevel eventual
  }