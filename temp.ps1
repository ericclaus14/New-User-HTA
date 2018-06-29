$userInput = Read-Host "Please enter text"
Write-Host $userInput


$myArr = @("Hello","My","Person")

$myArr | Out-GridView