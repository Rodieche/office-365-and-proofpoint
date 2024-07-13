# Connecting to Exchange Online
- Check the requirements for Exchange Online (Microsoft 365),
- Run ps Windows PowerShell.
- Install module:
```powershell
Install-Module -Name ExchangeOnlineManagement -Force
```
- Check your Execution policy settings:
```powershell
Get-ExecutionPolicy
```
-By default, the execution policy is set to Restricted. To successfully connect to Exchange Online with PowerShell, it is recommended to set the policy to RemoteSigned:
```powershell
Set-ExecutionPolicy RemoteSigned
```
- Connect to Exchange Online (Microsoft 365) using the Connect-ExchangeOnline cmdlet. You will be asked to sign in with your Microsoft 365 administrator credentials:
```powershell
Connect-ExchangeOnline
```
```powershell
$mailboxes | foreach { Get-User $_.UserPrincipalName | select FirstName, LastName, DisplayName, WindowsEmailAddress } | export-csv -NoTypeInformation .\Mailboxes.csv -Delimiter ";" -Encoding unicode
```
>NO TESTEADO
- Try if mailbox type is returned
```powershell
$mailboxes | foreach { 
    $mailbox = Get-Mailbox $_.UserPrincipalName
    $user = Get-User $_.UserPrincipalName 
    $type = if ($mailbox.RecipientTypeDetails -eq 'SharedMailbox') { 'SharedMailbox' } else { 'UserMailbox' }
    [PSCustomObject]@{
        FirstName = $user.FirstName
        LastName = $user.LastName
        DisplayName = $user.DisplayName
        WindowsEmailAddress = $user.WindowsEmailAddress
        MailboxType = $type
    }
} | export-csv -NoTypeInformation .\Mailboxes.csv -Delimiter ";" -Encoding unicode
```
- To disconnect, type:
```powershell
Disconnect-ExchangeOnline
```