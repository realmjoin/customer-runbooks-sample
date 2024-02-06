<#
  .SYNOPSIS
    Check for Autopilot devices not yet onboarded to Intune. Add these to an exclusion group.

  .DESCRIPTION
    Check for Autopilot devices not yet onboarded to Intune. Add these to an exclusion group.

  .NOTES
  Permissions
   MS Graph (API): 
#>

#Requires -Modules @{ModuleName = "RealmJoin.RunbookHelper"; ModuleVersion = "0.8.3" }, ExchangeOnlineManagement

param(
  # EntraID Group Object ID of exclusion group for Defender Compliance.
  [Parameter(Mandatory = $true)]
  [string]$exclusionGroupId,
  [int] $maxAgeInDays = 10,
  # CallerName is tracked purely for auditing purposes
  [Parameter(Mandatory = $true)]
  [string]$CallerName    
)

Connect-RjRbGraph

Write-RjRbLog -Message "Caller: '$CallerName'" -Verbose

$InGraceDevices = @()

## Get all Autopiolt Devices
$APDevices = Invoke-RjRbRestMethodGraph -Resource "/deviceManagement/windowsAutopilotDeviceIdentities" -FollowPaging

## Search for the Autopilot Devices not yet present in Intune
$APDevices | Where-Object { $_.enrollmentState -ne "enrolled" } | ForEach-Object {
  $device = Invoke-RjRbRestMethodGraph -Resource "/devices" -OdFilter "deviceId eq '$($_.azureActiveDirectoryDeviceId)'"
  if ($device) {
    $InGraceDevices += $device
  }
}

## Search for Intune devices younger than maxAgeInDays
$intuneDevices = Invoke-RjRbRestMethodGraph -Resource "/deviceManagement/managedDevices" -FollowPaging
$intuneDevices | Where-Object { ($_.enrolledDateTime -ge (Get-Date).AddDays(-$maxAgeInDays)) -and ($_.operatingSystem -eq "Windows") } | ForEach-Object {
  $device = Invoke-RjRbRestMethodGraph -Resource "/devices" -OdFilter "deviceId eq '$($_.azureADDeviceId)'"
  if ($device) {
    $InGraceDevices += $device
  }
}

## Get current members of the exclusion group
$exclusionGroupMembers = Invoke-RjRbRestMethodGraph -Resource "/groups/$exclusionGroupId/members" -FollowPaging

## Remove members from the group that are not in the InGraceDevices list
$exclusionGroupMembers | Where-Object { $InGraceDevices.id -notcontains $_.id } | ForEach-Object {
  Invoke-RjRbRestMethodGraph -Resource "/groups/$exclusionGroupId/members/$($_.id)/`$ref" -Method Delete
}

## Add members to the group that are in the InGraceDevices list and not yet members of the group
$InGraceDevices | Where-Object { $exclusionGroupMembers.id -notcontains $_.id } | ForEach-Object {
  $body = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($_.id)" }
  Invoke-RjRbRestMethodGraph -Resource "/groups/$exclusionGroupId/members/`$ref" -Method Post -Body $body
}

