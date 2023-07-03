<#
.SYNOPSIS
    This script will in-/exclude a user from a policy assignment by handling group memberships. 

.DESCRIPTION
    This script will make sure a user is member of a given policy assignment group (inclusion) OR is member of a given policy exclusion group (exclusion).
    The user will be member of exactly one of the groups. 


  .INPUTS
  RunbookCustomization: {
        "Parameters": {
            "UserName": {
                "Hide": true
            },
            "PolicyAssignmentGroupName": {
                "Hide": true
            },
            "PolicyExclusionGroupName": {
                "Hide": true
            },
            "exclude": {
                "DisplayName": "Assign or exclude from policy assignment",
                "SelectSimple": {
                    "Assign policy": false,
                    "Exclude/exempt from policy": true
                }
            },
            "CallerName": {
                "Hide": true
            }
        }
    }
#>

#Requires -Modules @{ModuleName = "RealmJoin.RunbookHelper"; ModuleVersion = "0.8.1" }

param(
    [Parameter(Mandatory = $true)]
    [ValidateScript( { Use-RJInterface -Type Graph -Entity User -DisplayName "User" } )]
    [String] $UserName,
    [string] $PolicyAssignmentGroupName = "cfg - Enforce Device Compliance",
    [string] $PolicyExclusionGroupName = "cfg - Enforce Device Compliance - Exclusion",
    [bool] $exclude = $false,
    # CallerName is tracked purely for auditing purposes
    [Parameter(Mandatory = $true)]
    [string] $CallerName
)

Write-RjRbLog -Message "Caller: '$CallerName'" -Verbose

Connect-RjRbGraph

# Get the user
$targetUser = Invoke-RjRbRestMethodGraph -Resource "/users" -OdFilter "userPrincipalName eq '$UserName'" -ErrorAction SilentlyContinue
if (-not $targetUser) {
    throw ("User '$UserName' not found.")
}

# Get the policy assignment group
$policyAssignmentGroup = Invoke-RjRbRestMethodGraph -Resource "/groups" -OdFilter "displayName eq '$PolicyAssignmentGroupName'" -ErrorAction SilentlyContinue -UriQueryRaw '$expand=members'
if (-not $policyAssignmentGroup) {
    throw ("Policy assignment group '$PolicyAssignmentGroupName' not found.")
}
if (([array]$policyAssignmentGroup).count -ne 1) {
    throw ("Policy assignment group '$PolicyAssignmentGroupName' is not unique.")
}

# Get the policy exclusion group
$policyExclusionGroup = Invoke-RjRbRestMethodGraph -Resource "/groups" -OdFilter "displayName eq '$PolicyExclusionGroupName'" -ErrorAction SilentlyContinue -UriQueryRaw '$expand=members'
if (-not $policyExclusionGroup) {
    throw ("Policy exclusion group '$PolicyExclusionGroupName' not found.")
}
if (([array]$policyExclusionGroup).count -ne 1) {
    throw ("Policy exclusion group '$PolicyExclusionGroupName' is not unique.")
}

# Check if the user is member of the policy assignment group and the policy exclusion group
$isMemberOfAssignmentGroup = (($policyAssignmentGroup.members | Where-Object { $_.id -eq $targetUser.id }) -ne $null)
$isMemberOfExclusionGroup = (($policyExclusionGroup.members | Where-Object { $_.id -eq $targetUser.id }) -ne $null)

# Prepare a body for group add requests
$groupAddBody = @{
    "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($targetUser.id)"
} 

if ($exclude) {
    # Make sure user is member of the policy exclusion group and not member of the policy assignment group
    if (-not $isMemberOfExclusionGroup) {
        "## Adding user $UserName to policy exclusion group '$PolicyExclusionGroupName'"
        Invoke-RjRbRestMethodGraph -Resource "/groups/$($policyExclusionGroup.id)/members/`$ref" -Method POST -Body $groupAddBody | Out-Null
    }
    if ($isMemberOfAssignmentGroup) {
        "## Removing user $UserName from policy assignment group '$PolicyAssignmentGroupName'"
        Invoke-RjRbRestMethodGraph -Resource "/groups/$($policyAssignmentGroup.id)/members/$($targetUser.id)/`$ref" -Method DELETE | Out-Null
    }
} else {
    # Make sure user is member of the policy assignment group and not member of the policy exclusion group
    if (-not $isMemberOfAssignmentGroup) {
        "## Adding user $UserName to policy assignment group '$PolicyAssignmentGroupName'"
        Invoke-RjRbRestMethodGraph -Resource "/groups/$($policyAssignmentGroup.id)/members/`$ref" -Method POST -Body $groupAddBody | Out-Null
    }
    if ($isMemberOfExclusionGroup) {
        "## Removing user $UserName from policy exclusion group '$PolicyExclusionGroupName'"
        Invoke-RjRbRestMethodGraph -Resource "/groups/$($policyExclusionGroup.id)/members/$($targetUser.id)/`$ref" -Method DELETE | Out-Null
    }
}

"## Done"