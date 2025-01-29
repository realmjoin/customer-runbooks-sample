<#
  .SYNOPSIS
  Import windows devices into Windows Autopilot from a list (CSV).

  .DESCRIPTION
  Import windows devices into Windows Autopilot from a list (CSV).
  The list is given as a URL / SAS-Link.
  The list is expected to have the following columns: Device Serial Number,Windows Product ID,Hardware Hash
  The list is expected to be separated by a comma.

  .NOTES
  Permissions: 
  MS Graph (API):
  - DeviceManagementServiceConfig.ReadWrite.All

#>

#Requires -Modules Microsoft.Graph.Authentication

param(
    # Path to the CSV file containing the devices
    [Parameter(Mandatory = $true)]
    [string] $FullPath,
    # Wait for the import to finish per device
    [bool] $Wait = $false
)

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

#Connect-MgGraph -Identity

## Read the CSV file
$allDevicesFromCSV = Import-Csv -Path $FullPath -Delimiter ","

## Get all devices (OData query is not supported for this endpoint)
$allDevices = Invoke-MgGraphRequestAll -Uri "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities"
$allDeviceInImport = Invoke-MgGraphRequestAll -Uri "https://graph.microsoft.com/v1.0/deviceManagement/importedWindowsAutopilotDeviceIdentities/"

# Check and import each device
foreach ($device in $allDevicesFromCSV) {
    $SerialNumber = $device.'Device Serial Number'
    $HardwareIdentifier = $device.'Hardware Hash'
    #$WindowsProductID = $device.'Windows Product ID'

    ## Check if the device is already imported
    if ((-not ($allDevices.serialNumber -contains $SerialNumber)) -and (-not ($allDeviceInImport.serialNumber -contains $SerialNumber))) {
        "## Importing device $SerialNumber"
        $body = @{
            serialNumber       = $SerialNumber 
            hardwareIdentifier = $HardwareIdentifier
            #windowsProductID   = $WindowsProductID
        }

        ## MS removed the ability to assign users directly via Autopilot
        ##if ($AssignedUser) {
        ##    ## Find assigned user's name
        ##    $username = (Invoke-RjRbRestMethodGraph -Resource "/users/$AssignedUser").UserPrincipalName
        ##    $body += @{ assignedUserPrincipalName = $username }
        ##}

        ##if ($groupTag) {
        ##    $body += @{ groupTag = $GroupTag }
        ##}

        $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/deviceManagement/importedWindowsAutopilotDeviceIdentities" -Method POST -Body $body
        "## Import of device $SerialNumber started."

        # Track the import's progress
        if ($Wait) {
            while (($result.state.deviceImportStatus -ne "complete") -and ($result.state.deviceImportStatus -ne "error")) {
                "## ."
                Start-Sleep -Seconds 20
                #$result = Invoke-RjRbRestMethodGraph -Resource "/deviceManagement/importedWindowsAutopilotDeviceIdentities/$($result.id)" -Method Get
                $result = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/deviceManagement/importedWindowsAutopilotDeviceIdentities/$($result.id)" -Method GET
            }
            if ($result.state.deviceImportStatus -eq "complete") {
                "## Import of device $SerialNumber is successful."
            }
            else {
                write-error ($result.state)
                throw "Import of device $SerialNumber failed."
            }
        }

    }
    else {
        "## Device $SerialNumber is already imported."
    }

}


