$message = Get-ADOrganizationalUnit -Filter * | select DistinguishedName | Out-GridView -Title "Select an OU and click OK" -OutputMode Single
$message.DistinguishedName | clip