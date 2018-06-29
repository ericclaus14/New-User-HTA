Param(
	[Parameter(Mandatory=$true)][String]$firstName,
	[Parameter(Mandatory=$true)][String]$lastName,
	[String]$userName=$firstName + $lastName,
	[Parameter(Mandatory=$true)][String]$pass,
    [String]$OU="causers"
)

Import-Module activedirectory
<#
$firstName = "first"
$lastName = "last"
$userName = "sam2"
$pass = "Password1234!!"
#>
$securePass = (ConvertTo-SecureString -AsPlainText $pass -Force)
$path = "OU=$OU,DC=ad,DC=collegedaleacademy,DC=com"
$name = "$firstName $lastName"

Try{
    New-ADUser -Name $name -GivenName $firstName -SamAccountName $userName -Surname $lastName -AccountPassword $securePass -Enabled $true
    $message = "The new user has been created."
    $exitCode = "0"
}
Catch{
    $message = $_.Exception.Message
    $exitCode = "1"
}
Finally{
    # Send result to HTA via clipboard
    $message | clip
    Exit($exitCode)
}