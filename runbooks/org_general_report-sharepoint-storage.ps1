<#
  .SYNOPSIS
  Report on the storage usage of the SharePoint Online tenant.

  .DESCRIPTION
  Fetches the storage usage of the SharePoint Online tenant and sends a report via email (optional).
#>

## This runbook requires to be used in a PS7 environment, as it uses the PnP.PowerShell module.
## No 'Requires" Statement, as RJ Portal can not handle it yet for PS7

param(
    [Parameter(Mandatory = $true)]    
    [string]$SPOAdminUrl,
    [string]$sendAlertFrom,
    [string]$sendAlertTo,
    ## CallerName is tracked purely for auditing purposes
    [Parameter(Mandatory = $true)]
    [String] $CallerName
)

Write-RjRbLog -Message "Caller: '$CallerName'" -Verbose

Connect-PnPOnline -Url $SPOAdminUrl -ManagedIdentity

"## Fetching SharePoint Online tenant storage usage..."
""

$tenantStorageCapacity = Get-PnPTenant | Select-Object -ExpandProperty StorageQuota;

"Tenant Storage Capacity: " + $tenantStorageCapacity + " MB"
"Tenant Storage Capacity: " + $tenantStorageCapacity / 1024 + " GB"
"Tenant Storage Capacity: " + $tenantStorageCapacity / 1024 / 1024 + " TB"
""

$tenantStorageMultiGeo = Invoke-PnPSPRestMethod -Url ($SPOAdminUrl + "/_api/StorageQuotas(geoLocation='LOCAL')?api-version=1.3.1")
$tenantStorageMultiGeoAvailable = $tenantStorageMultiGeo.GeoAvailableStorageMB;

"GEO Tenant Storage Location: " + $tenantStorageMultiGeo.GeoLocation
"GEO Tenant Storage Capacity: " + $tenantStorageMultiGeo.TenantStorageMB + " MB"
"GEO Tenant Storage Capacity: " + $tenantStorageMultiGeo.TenantStorageMB / 1024 + " GB"
"GEO Tenant Storage Capacity: " + $tenantStorageMultiGeo.TenantStorageMB / 1024 / 1024 + " TB"
""

"GEO Tenant Storage Available: " + $tenantStorageMultiGeo.GeoAvailableStorageMB + " MB"
"GEO Tenant Storage Available: " + $tenantStorageMultiGeo.GeoAvailableStorageMB / 1024 + " GB"
"GEO Tenant Storage Available: " + $tenantStorageMultiGeo.GeoAvailableStorageMB / 1024 / 1024 + " TB"
""

$sites = Get-PnPTenantSite -IncludeOneDriveSites | Select Url, StorageUsageCurrent, @{Name = 'BaseTemplate'; Expression = { $_.Template.Split('#')[0] } } | Group-Object BaseTemplate

$sitesOnedrive = $sites | Where-Object { $_.Name -eq 'SPSPERS' }
$sitesOnedriveStats = $sitesOnedrive.Group | Measure-Object StorageUsageCurrent -Sum
$sitesOneDriveStorageUsageCurrent = $sitesOnedriveStats.Sum;

"OneDrive Count: " + $sitesOnedriveStats.Count
"OneDrive StorageUsageCurrent: " + $sitesOnedriveStats.Sum + " MB"
"OneDrive StorageUsageCurrent: " + $sitesOnedriveStats.Sum / 1024 + " GB"
"OneDrive StorageUsageCurrent: " + $sitesOnedriveStats.Sum / 1024 / 1024 + " TB"
""

$sitesOnlyStats = $sites | Where-Object { $_.Name -ne 'SPSPERS' } | Select-Object -ExpandProperty Group | Measure-Object -Property StorageUsageCurrent -Sum
$sitesOnlyStorageUsageCurrent = $sitesOnlyStats.Sum;

"Sites Count: " + $sitesOnlyStats.Count
"Sites StorageUsageCurrent: " + $sitesOnlyStats.Sum + " MB"
"Sites StorageUsageCurrent: " + $sitesOnlyStats.Sum / 1024 + " GB"
"Sites StorageUsageCurrent: " + $sitesOnlyStats.Sum / 1024 / 1024 + " TB"
""

"Multi Geo Avilability Reporting: " + $tenantStorageMultiGeoAvailable
"SitesOnly Calculation: " + [int]($tenantStorageCapacity - $sitesOnlyStorageUsageCurrent)
""

if (-not $sendAlertFrom -or -not $sendAlertTo) {
    "## No email addresses provided. Skipping email report."
    return
}

Connect-RjRbGraph

$HTMLBody = "<h2>SPO Tenant Report</h2>"
$HTMLBody += "<br/>"
$HTMLBody += "<pre>Tenant Storage Capacity: " + $tenantStorageCapacity + " MB</pre>"
$HTMLBody += "<pre>Tenant Storage Capacity: " + $tenantStorageCapacity / 1024 + " GB</pre>"
$HTMLBody += "<pre>Tenant Storage Capacity: " + $tenantStorageCapacity / 1024 / 1024 + " TB</pre>"
$HTMLBody += "<br/>"
$HTMLBody += "<pre>GEO Tenant Storage Location: " + $tenantStorageMultiGeo.GeoLocation + "</pre>"
$HTMLBody += "<pre>GEO Tenant Storage Capacity: " + $tenantStorageMultiGeo.TenantStorageMB + " MB</pre>"
$HTMLBody += "<pre>GEO Tenant Storage Capacity: " + $tenantStorageMultiGeo.TenantStorageMB / 1024 + " GB</pre>"
$HTMLBody += "<pre>GEO Tenant Storage Capacity: " + $tenantStorageMultiGeo.TenantStorageMB / 1024 / 1024 + " TB</pre>"
$HTMLBody += "<br/>"
$HTMLBody += "<pre>GEO Tenant Storage Available: " + $tenantStorageMultiGeo.GeoAvailableStorageMB + " MB</pre>"
$HTMLBody += "<pre>GEO Tenant Storage Available: " + $tenantStorageMultiGeo.GeoAvailableStorageMB / 1024 + " GB</pre>"
$HTMLBody += "<pre>GEO Tenant Storage Available: " + $tenantStorageMultiGeo.GeoAvailableStorageMB / 1024 / 1024 + " TB</pre>"
$HTMLBody += "<br/>"
$HTMLBody += "<pre>OneDrive Count: " + $sitesOnedriveStats.Count + "</pre>"
$HTMLBody += "<pre>OneDrive StorageUsageCurrent: " + $sitesOnedriveStats.Sum + " MB</pre>"
$HTMLBody += "<pre>OneDrive StorageUsageCurrent: " + $sitesOnedriveStats.Sum / 1024 + " GB</pre>"
$HTMLBody += "<pre>OneDrive StorageUsageCurrent: " + $sitesOnedriveStats.Sum / 1024 / 1024 + " TB</pre>"
$HTMLBody += "<br/>"
$HTMLBody += "<pre>Sites Count: " + $sitesOnlyStats.Count + "</pre>"
$HTMLBody += "<pre>Sites StorageUsageCurrent: " + $sitesOnlyStats.Sum + " MB</pre>"
$HTMLBody += "<pre>Sites StorageUsageCurrent: " + $sitesOnlyStats.Sum / 1024 + " GB</pre>"
$HTMLBody += "<pre>Sites StorageUsageCurrent: " + $sitesOnlyStats.Sum / 1024 / 1024 + " TB</pre>"
$HTMLBody += "<br/>"
$HTMLBody += "<pre>Multi Geo Avilability Reporting: " + $tenantStorageMultiGeoAvailable + "</pre>"
$HTMLBody += "<pre>SitesOnly Calculation: " + [int]($tenantStorageCapacity - $sitesOnlyStorageUsageCurrent) + "</pre>"

$message = @{
    subject      = "[Automated Report] SPO Tenant Report"
    body         = @{
        contentType = "HTML"
        content     = $HTMLBody
    }
    toRecipients = @(
        @{
            emailAddress = @{
                address = $sendAlertTo
            }
        }
    )
}

"Sending report to '$sendAlertTo'..."
Invoke-RjRbRestMethodGraph -Resource "/users/$sendAlertFrom/sendMail" -Method POST -Body @{ message = $message } -ContentType "application/json" | Out-Null
"Report sent to '$sendAlertTo'."