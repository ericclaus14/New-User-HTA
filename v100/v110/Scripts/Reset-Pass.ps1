Param (
    [Parameter(Mandatory=$true)][string]$sam,
    [Parameter(Mandatory=$true)][string]$pass
)

Import-Module activedirectory

$securePass = (ConvertTo-SecureString -AsPlainText $pass -Force)

Try{
    Set-ADAccountPassword $sam -NewPassword $securePass
    $message = "The password for $sam has been reset."
}
Catch{
    $message = $_.Exception.Message
}
Finally{
    $message | clip
}