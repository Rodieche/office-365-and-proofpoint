const Powershell = require('powershell');

const setExecutionPolicyCmd = 'Set-ExecutionPolicy RemoteSigned -RunAs';
const InstallExchangeOnlineModuleCmd = 'Install-Module -Name ExchangeOnlineManagement -Force -RunAs';
const ConnectExchangeCmd = 'Connect-ExchangeOnline';
const getMailBoxesExchangeCmd = '$mailboxes | foreach { Get-User $_.UserPrincipalName | select FirstName, LastName, DisplayName, WindowsEmailAddress } | export-csv -NoTypeInformation .\Mailboxes.csv -Delimiter ";" -Encoding unicode';
const DisconnectExchangeCmd = 'Disconnect-ExchangeOnline';

const setExecutionPolicy = new Powershell(setExecutionPolicyCmd);
setExecutionPolicy.on('error', err => {
    console.log(err);
});

setExecutionPolicy.on('output', data => {
    console.log(data);
});

module.exports = {
    setExecutionPolicy
}