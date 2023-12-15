<#
  .SYNOPSIS
  Get a list of all shared mailboxes and their attributes.

  .DESCRIPTION
  Fetches all existing shared mailboxes and their respective attributes as JSON.
#>

#Requires -Modules @{ModuleName = "RealmJoin.RunbookHelper"; ModuleVersion = "0.8.3" }, ExchangeOnlineManagement


param (
  [Parameter(Mandatory = $true)]
  [string] $Username,
  ## CallerName is tracked purely for auditing purposes
  [Parameter(Mandatory = $true)]
  [string] $CallerName
)

Write-RjRbLog -Message "Caller: '$CallerName'" -Verbose

try {
  Connect-RjRbExchangeOnline | Out-Null

  #Write-Host "## Grabbing information for the provided mailbox..."
  
  $mailboxData = Get-EXOMailbox -Identity $Username -Properties DisplayName, UserPrincipalName, EmailAddresses, RecipientTypeDetails, PrimarySmtpAddress, ExchangeObjectId, WhenCreated, IsMailboxEnabled, DistinguishedName, OrganizationalUnit, ExchangeVersion
  $orgUnitRoot = Get-Mailbox -Identity $Username | Select-Object OrganizationalUnitRoot

  if($null -eq $mailboxData || $null -eq $orgUnitRoot){
    throw "## Mailbox does not exist. Check spelling and try again."
  } 

  $mailboxAttributes = New-Object -TypeName PSObject -Property @{
    ExchangeObjectId       = $mailboxData.ExchangeObjectId
    WhenCreated            = $mailboxData.WhenCreated
    IsMailboxEnabled       = $mailboxData.IsMailboxEnabled
    OrganizationalUnitRoot = $orgUnitRoot.OrganizationalUnitRoot
    OrganizationalUnit     = $mailboxData.OrganizationalUnit
    DisplayName            = $mailboxData.DisplayName
    UserPrincipalName      = $mailboxData.UserPrincipalName
    EmailAddresses         = $mailboxData.EmailAddresses
    RecipientTypeDetails   = $mailboxData.RecipientTypeDetails
    PrimarySmtpAddress     = $mailboxData.PrimarySmtpAddress
    DistinguishedName      = $mailboxData.DistinguishedName
    ExchangeVersion        = $mailboxData.ExchangeVersion
  }

  $mailboxAttributes | ConvertTo-JSON

  #Write-Host "## Saving all data to a .json file..."
  #$mailboxAttributes | ConvertTo-JSON | Out-File -FilePath ".\$($Username)_Attributes.json"
  #Write-Host "## File saved under '$($Username)_Attributes.json'. "

}
finally {
  Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Continue | Out-Null
}