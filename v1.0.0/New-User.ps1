Param(
	[Parameter(Mandatory=$true)][String]$firstName,
	[Parameter(Mandatory=$true)][String]$lastName,
	[String]$userName=$firstName + $lastName,
	[String]$pass,
    [
    [String]$OU,
    [Array]$groups
)

Import-Module activedirectory

# If you need to enforce a specific username format,
# set the format below and uncomment the line.
#$userName = "$($firstName[0])$lastName"
$securePass = (ConvertTo-SecureString -AsPlainText $pass -Force)

$name = "$firstName $lastName"
$homeDir = "\\nas1.ad.collegedaleacademy.com\nasshare\home\$userName"
$homeDrive = "P:"

# The parameters to be used with New-ADUser
$parameters = @{
    "Name"=$name
    "GivenName"=$firstName
    "SamAccountName"=$userName
    "SurName"=$lastName
    "AccountPassword"=$securePass
    "Enabled"=$true
    "HomeDirectory"=$homeDir
    "HomeDrive"=$homeDrive
}


Try{
    If ($OU){
        $parameters.Add("Path",$OU)
        New-ADUser @parameters
    }
    If ($groups){
        #
    }
    
    New-ADUser @parameters
    
    $message = "The new user has been created."
    $exitCode = "0"
}
Catch{
    # Message to be returned to HTA is whatever error occured above
    $message = $_.Exception.Message
    $exitCode = "1"
}
Finally{
    # Send result to HTA via clipboard
    $message | clip
    # Exit script, returning exit code
    Exit($exitCode)
}
