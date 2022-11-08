$AdminUserName = Get-ChildItem -Path ENV:OFFICEUSER -ErrorAction stop
$AdminPassword = Get-ChildItem -path Env:OFFICEPASSWD -ErrorAction stop
if ($AdminUserName.value -eq "null" -or $AdminPassword.value -eq "null") {
    Write-Error "office user or pasword is null"
    Exit -1
}
$SecurePassword = ConvertTo-SecureString $AdminPassword.value -AsPlainText -Force -ErrorAction stop
$Credential = New-Object System.Management.Automation.PSCredential -argumentlist $AdminUserName.value, $SecurePassword -ErrorAction stop
# if you are not in china just remove -AzureEnvironmentName AzureChinaCloud
Connect-AzureAD -AzureEnvironmentName AzureChinaCloud -Credential $Credential -ErrorAction stop
Connect-MsolService -AzureEnvironment AzureChinaCloud -Credential $Credential -ErrorAction stop