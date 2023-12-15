<#
  .SYNOPSIS
  Get a list of all shared mailboxes and their attributes.

  .DESCRIPTION
  Fetches all existing shared mailboxes and their respective attributes as JSON.
#>

#Requires -Modules @{ModuleName = "RealmJoin.RunbookHelper"; ModuleVersion = "0.8.3" }, ExchangeOnlineManagement


param (
  [Parameter(Mandatory = $true)]
  [string] $MailboxUPN,
  ## CallerName is tracked purely for auditing purposes
  [Parameter(Mandatory = $true)]
  [TypeName] $CallerName
)

Write-RjRbLog -Message "Caller: '$CallerName'" -Verbose

try {
  Connect-RjRbExchangeOnline

  Write-Host "## Grabbing information for the provided mailbox..."
  
  $mailboxData = Get-EXOMailbox -Identity $MailboxUPN -Properties DisplayName, UserPrincipalName, EmailAddresses, RecipientTypeDetails, PrimarySmtpAddress, ExchangeObjectId, WhenCreated, IsMailboxEnabled, DistinguishedName, OrganizationalUnit, ExchangeVersion
  $orgUnitRoot = Get-Mailbox -Identity $MailboxUPN | Select-Object OrganizationalUnitRoot

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

  Write-Host "## Saving all data to a .json file..."
  $mailboxAttributes | ConvertTo-JSON | Out-File -FilePath ".\$($MailboxUPN)_Attributes.json"
  Write-Host "## File saved under '$($MailboxUPN)_Attributes.json'. "

}
finally {
  Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Continue | Out-Null
}