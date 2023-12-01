<#
  .SYNOPSIS
  Get a list of all shared mailboxes and their attributes.

  .DESCRIPTION
  Fetches all existing shared mailboxes and their respective attributes as JSON.
#>

#Requires -Modules @{ModuleName = "RealmJoin.RunbookHelper"; ModuleVersion = "0.8.3" }, ExchangeOnlineManagement


param (
  ## CallerName is tracked purely for auditing purposes
  [Parameter(Mandatory = $true)]
  [TypeName] $CallerName
)

Write-RjRbLog -Message "Caller: '$CallerName'" -Verbose

try {
  Connect-ExchangeOnline

  Write-Host "## Grabbing all shared mailboxes..."

  $AllSharedMailboxes = @()

  ## Grab a list of all shared mailboxes with their UPN via Get-EXOMailbox 
  $sharedMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize unlimited | Select-Object  userPrincipalName

  ## Go thru each mailbox and call Get-Mailbox and then Get-EXOMailbox to grab respective attributes in two separate PSObjects then slap them in one PSCustomObj
  foreach ($sharedmailbox in $sharedMailboxes) {
    ## grab the specific 5 attributes, for every shared mailbox, only available with the old cmdlet and save them in an array
    $MailboxProperties = Get-Mailbox -Identity $_.UserPrincipalName | Select-Object ExchangeObjectId, WhenCreated, IsMailboxEnabled, OrganizationalUnitRoot, OrganizationalUnit
    ## do the same but for the newer cmdlet
    $EXOMailboxProperties = Get-EXOMailbox -Identity $_.UserPrincipalName | Select-Object DisplayName, UserPrincipalName, EmailAddresses, RecipientTypeDetails, PrimarySmtpAddress, DistinguishedName, ExchangeVersion
  }

  ## in order to combine all of the attributes respectively for each mailbox do a simple index position array for loop 
  $AllSharedMailboxes = for($i = 0; $i -lt $MailboxProperties.Count; $i++){
    $obj1 = $MailboxProperties[$i]
    $obj2 = $EXOMailboxProperties[$i]

    $CombinedProperties = New-Object -TypeName PSObject -Property @{
      ExchangeObjectId       = $obj1.ExchangeObjectId
      WhenCreated            = $obj1.WhenCreated
      IsMailboxEnabled       = $obj1.IsMailboxEnabled
      OrganizationalUnitRoot = $obj1.OrganizationalUnitRoot
      OrganizationalUnit     = $obj1.OrganizationalUnit
      DisplayName            = $obj2.DisplayName
      UserPrincipalName      = $obj2.UserPrincipalName
      EmailAddresses         = $obj2.EmailAddresses
      RecipientTypeDetails   = $obj2.RecipientTypeDetails
      PrimarySmtpAddress     = $obj2.PrimarySmtpAddress
      DistinguishedName      = $obj2.DistinguishedName
      ExchangeVersion        = $obj2.ExchangeVersion
    }

    $CombinedProperties
  }
  ## Convert the array into a JSON
  Write-Host "## Saving all data to a .json file..."
  $AllSharedMailboxes | ConvertTo-JSON | Out-File -FilePath ".\AllSharedMailboxes.json"
  Write-Host "## File saved under 'AllSharedMailboxes.json'. "
}
finally {
  Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Continue | Out-Null
}