<#
  .SYNOPSIS
  Export a list of TPM Hardware Specs of all Managed Devices.

  .DESCRIPTION
  Exports a List of the Managed Device Name, Managed Device ID, TPM Version, TPM Manufacturer Version, TPM Manufacturer Name, SystemManagementBIOSVersion of all of the Managed Devices.

  .NOTES
  Permissions
   MS Graph (API): 
   - DeviceManagementManagedDevices.Read.All
   - Device.Read.All
#>

#Requires -Modules @{ModuleName = "RealmJoin.RunbookHelper"; ModuleVersion = "0.6.0" }   

Connect-RjRbGraph 

$devices = Invoke-RjRbRestMethodGraph -Resource "/deviceManagement/managedDevices" -OdSelect "id" -Beta -FollowPaging

"Device Name;Device ID;TPM Manufacturer;TPM Version;TPM Specification Version;system Management BIOS Version"

foreach ($device in $devices) {
  $deviceData = Invoke-RjRbRestMethodGraph -Resource "/deviceManagement/managedDevices/$($device.id)" -OdSelect "id,deviceName,hardwareInformation" -Beta
  $deviceData.deviceName + ";" + $deviceData.id + ";" + $deviceData.hardwareInformation.tpmManufacturer + ";" + $deviceData.hardwareInformation.tpmVersion + ";" + $deviceData.hardwareInformation.tpmSpecificationVersion + ";" + $deviceData.hardwareInformation.systemManagementBIOSVersion
}