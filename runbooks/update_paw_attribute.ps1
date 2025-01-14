<#
  .SYNOPSIS
  Update all PAW Device's ExtensionAttributes.

  .DESCRIPTION
  Fetches all AAD/Autopilot Devices and keeps their ExtensionAttributes up to datex.

  .NOTES
  Permissions: 
  MS Graph (API):
  - Device.ReadWrite.All

#>

#Requires -Modules Microsoft.Graph.Authentication

# Handle paging (throttling is already handled by MGGraph)
# Also - directly returns "values" array
function Invoke-MgGraphRequestAll {
    param (
        [string]$Uri
    )
    $result = Invoke-MgGraphRequest -Uri $Uri -Method GET
    $results = $result.value
    while ($sresult.'@odata.nextLink') {
        $nextLink = $sresult.'@odata.nextLink'
        $result = Invoke-MgGraphRequest -Uri $nextLink -Method GET
        $results += $result.value
    }
    return $results
}

Connect-MgGraph -Identity

## Get all Devices
$allDevices = Invoke-MgGraphRequestAll -Uri "https://graph.microsoft.com/v1.0/devices"

## Search for the PAW Devices 
foreach ($pawDevice in $allDevices) {
    ## testing purposes, correct to  afterwards
    if ($pawDevice.physicalIds -contains "[OrderId]:PAW") {

        ## print check
        "Physical IDs found for PAW Device: " + $pawDevice.id + ", Display Name: " + $pawDevice.displayName
        
        ## Create request body json
        $body = @{
            "extensionAttributes" = @{
                "extensionAttribute1" = "PAW"
            }
        }
        
        ## Call Graph API with Patch command to update the extensionAttribute
        Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/devices/$($pawDevice.id)" -Method PATCH -Body $body -ContentType "application/json"

        ## Query Graph API to check if device's extensionAttribute has been set correctly
        #$checkDevice = Invoke-RjRbRestMethodGraph -Resource "/devices/$($pawDevice.id)" -OdSelect "extensionAttributes" -Method Get
        $checkDevice = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/devices/$($pawDevice.id)?`$select=extensionAttributes" -Method GET
        
        ## Print to Terminal
        "## extensionAttribute1 of Device $($pawDevice.displayName) with the DeviceID $($pawDevice.id) has been set to: "
        $checkDevice.extensionAttributes.extensionAttribute1
    }
}
