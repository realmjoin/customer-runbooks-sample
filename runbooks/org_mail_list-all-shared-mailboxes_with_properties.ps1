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
  [String] $CallerName
)

Write-RjRbLog -Message "Caller: '$CallerName'" -Verbose

try {
  Connect-RjRbExchangeOnline | Out-Null

  #Write-Host "## Grabbing all shared mailboxes..."

  $AllSharedMailboxes = New-Object System.Collections.Generic.LinkedList[System.Object]

  ## Grab a list of all shared mailboxes with their UPN via Get-EXOMailbox 
  $sharedMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize unlimited -Properties userPrincipalName | Select-Object  userPrincipalName

  $sharedMailboxes | ForEach-Object {
    ## grab the specific 5 attributes, for every shared mailbox, only available with the old cmdlet and save them in an array
    $MailboxProperties = Get-Mailbox -Identity $_.UserPrincipalName | Select-Object ExchangeObjectId, WhenCreated, IsMailboxEnabled, OrganizationalUnitRoot, OrganizationalUnit
    ## do the same but for the newer cmdlet
    $EXOMailboxProperties = Get-EXOMailbox -Identity $_.UserPrincipalName | Select-Object DisplayName, UserPrincipalName, EmailAddresses, RecipientTypeDetails, PrimarySmtpAddress, DistinguishedName, ExchangeVersion

    ## combine all of the attributes 
    $obj = New-Object -TypeName PSObject -Property @{
      ExchangeObjectId       = $MailboxProperties.ExchangeObjectId
      WhenCreated            = $MailboxProperties.WhenCreated
      IsMailboxEnabled       = $MailboxProperties.IsMailboxEnabled
      OrganizationalUnitRoot = $MailboxProperties.OrganizationalUnitRoot
      OrganizationalUnit     = $MailboxProperties.OrganizationalUnit
      DisplayName            = $EXOMailboxProperties.DisplayName
      UserPrincipalName      = $EXOMailboxProperties.UserPrincipalName
      EmailAddresses         = $EXOMailboxProperties.EmailAddresses
      RecipientTypeDetails   = $EXOMailboxProperties.RecipientTypeDetails
      PrimarySmtpAddress     = $EXOMailboxProperties.PrimarySmtpAddress
      DistinguishedName      = $EXOMailboxProperties.DistinguishedName
      ExchangeVersion        = $EXOMailboxProperties.ExchangeVersion
    }
    $AllSharedMailboxes.Add(
      $obj
    )
  }


  ## Convert the array into a JSON
  $AllSharedMailboxes | ConvertTo-JSON 
  
  #Write-Host "## Saving all data to a .json file..."
  #$AllSharedMailboxes | ConvertTo-JSON | Out-File -FilePath ".\AllSharedMailboxes.json"
  #Write-Host "## File saved under 'AllSharedMailboxes.json'. "
}
finally {
  Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Continue | Out-Null
}