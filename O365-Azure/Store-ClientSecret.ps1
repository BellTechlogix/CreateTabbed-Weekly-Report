# Set credentials
$Credentials = Get-Credential  # Set your id and secret
$Credentials | Export-Clixml -Path $PSScriptRoot\Access.xml -Confirm:$false


# Read the credentials
$Credentials = Import-Clixml -Path $PSScriptRoot\Access.xml