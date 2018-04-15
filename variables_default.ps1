$ExcelPath = "$(Get-Item $PSScriptRoot)\work"
$LOG = "$ExcelPath\Log\Log.txt"
$TestCMDs = "$ExcelPath\Log\CMDs.txt"

$User = "Username"
$PW =  ConvertTo-SecureString -String "Password" -AsPlainText -Force
$UserCredential = New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $User, $PW 
$SPOnline=path-to-sharepoint.com