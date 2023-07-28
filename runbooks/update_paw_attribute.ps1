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

#Requires -Modules @{ModuleName = "RealmJoin.RunbookHelper"; ModuleVersion = "0.8.1" }

Connect-RjRbGraph


## Get all Devices
$allDevices = Invoke-RjRbRestMethodGraph -Resource "/devices" -FollowPaging

## Search for the PAW Devices 
foreach ($pawDevice in $allDevices) {
    ## testing purposes, correct to  afterwards
    if ($pawDevice.physicalIds -eq "[OrderId]:PAW") {

        ## print check
        "Physical IDs found for PAW Device: " + $pawDevice.id + ", Display Name: " + $pawDevice.displayName
        
        ## Create request body json
        $body = @{
            "extensionAttributes" = @{
                "extensionAttribute1" = "PAW"
            }
        }
        
        ## Call Graph API with Patch command to update the extensionAttribute
        Invoke-RjRbRestMethodGraph -Resource "/devices/$($pawDevice.id)" -Body $body -ContentType "application/json" -Method Patch

        ## Query Graph API to check if device's extensionAttribute has been set correctly
        $checkDevice = Invoke-RjRbRestMethodGraph -Resource "/devices/$($pawDevice.id)" -OdSelect "extensionAttributes" -Method Get
        
        ## Print to Terminal
        "## extensionAttribute1 of Device $($pawDevice.displayName) with the DeviceID $($pawDevice.id) has been set to: "
        $checkDevice.extensionAttributes.extensionAttribute1
    }
}
