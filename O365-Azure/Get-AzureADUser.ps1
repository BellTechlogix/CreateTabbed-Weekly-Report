<#
	Get-AzureAdUser.ps1
	Created By - Kristopher Roy
	Created On - 01 Mar 2023
	Modified On - 03 May 2023

	This Script Requires an Azure account with permissions to connect and pull users. Pulls a report of all Azure User accounts
#>

# Function to check if a module is installed
function Check-Module {
  param (
    [string]$Name,
    [switch]$Import = $true
  )

  # Check if the module is installed
  if (Get-Module -ListAvailable -Name $Name -ErrorAction SilentlyContinue) {
    # The module is installed, so import it if the switch is enabled
    if ($Import) {
      Import-Module -Name $Name -Force
    }
  } else {
        # Register the WindowsPowerShell Gallery as a trusted repository
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    
    # The module is not installed, so install it
    Install-Module -Name $Name
  }
}

function Connect-toAzure{
    [CmdletBinding()]

    param (
        [Parameter(Mandatory = $false)]
        [string]
        $TenantId,

        [Parameter(Mandatory = $false)]
        [string]
        $SubscriptionId,

        [Parameter(Mandatory = $true)]
        [string]
        $ClientId,

        [Parameter(Mandatory = $true)]
        [string]
        $ClientSecret
    )

    $InformationPreference = "Continue"

    Disable-AzContextAutosave -Scope Process | Out-Null

    $creds = [System.Management.Automation.PSCredential]::new($ClientId, (ConvertTo-SecureString $ClientSecret -AsPlainText -Force))
    Connect-AzAccount -Tenant $TenantId -Credential $creds -ServicePrincipal | Out-Null
    Write-Information "Connected to Azure..."

}

#Install-Module -Name Microsoft.Graph -AllowPrerelease

#$tenants = Get-CurrentUserTenants

# Increase the function capacity
$MaximumFunctionCount = 32768

# Check if the Microsoft.Graph module is installed
#Check-Module -Name Microsoft.Graph -Import

# Check if the AZ module is installed
#Check-Module -Name Az

# Force authentication
#Connect-MgGraph

# Read the credentials
$path = 'C:\Users\tankcr\source\repos\BellTechlogix\CreateTabbed-Weekly-Report\O365-Azure'
$Credentials = Import-Clixml -Path $Path\Access.xml

#Disconnect-MgGraph
connect-MgGraph -TenantId '1f05524b-c860-46e3-a2e8-4072506b3e4a' -clientsecretcredential $credentials
#Connect-toAzure -ClientId $ClientId -ClientSecret $ClientSecret -TenantId '1f05524b-c860-46e3-a2e8-4072506b3e4a'

$tenants = Get-CurrentUserTenants

$selectedtenet = MultipleSelectionBox -inputarray $tenants.Name -listboxtype One -prompt "Select Your Tenant"
$tenetID = ($tenants|where{$_.name -eq $selectedtenet}).ID
$subscriptions = Get-AzSubscription
$selectedSubscription = MultipleSelectionBox -inputarray ($subscriptions).Name -listboxtype One -prompt "Select Your Subscription"
$subscriptionID = ($subscriptions|where{$_.name -eq $selectedsubscription}).ID
set-azContext -Subscription $SubscriptionID
# Get the list of Azure users
$users = get-mguser -filter "UserType eq 'Member'" -Property 'UserType,SignInActivity,AccountEnabled,City,CompanyName,Country,Department,DisplayName,GivenName,SurName,ID,JobTitle,Mail,MemberOf,State,UserPrincipalName,LastLogonTimestamp,SecurityIdentifier'|select *, @{N='LastLogonTimestamp';E={ [datetime]$_.LastLogonTimestamp}}
FOREACH($user in $users)
{
    #get-MgUser -UserId $user.ID -Property SignInActivity|select LastSignInDateTime
    If($User.AdditionalProperties.Values.lastSignInDateTime -ne $null)
    {
        $user.LastLogonTimestamp = get-date $User.AdditionalProperties.Values.lastSignInDateTime
    }
}

#Get-AzAdUser -select 'Department,AccountEnabled,Department,UserType,UserPrincipalName,GivenName,SurName,ApproximateLastSignInDateTime,Manager,JobTitle,Identity,EmployeeType,Department,Country,City,StreetAddress,State,PostalCode,OfficeLocation' -AppendSelected|select *|where($_.UserType -ne 'Guest')

$users|select DisplayName,@{N='SID';E={($_.SecurityIdentifier)}},givenName,surName,UserPrincipalName,@{N='Domain';E={($_.UserPrincipalName.split('@')[1])}},Department,LastLogonTimestamp,@{N='dayssincelogon';E={(new-timespan -start (get-date $_.LastLogonTimestamp -Hour "00" -Minute "00") -End (get-date -Hour "00" -Minute "00")).Days}},JobTitle,ID,Groups

#$azure_users = Get-AzADUser
#FOREACH($AZUser in $azure_users)
#{Get-MGUser}

# Create a variable to store the list of Azure users
$azure_users_table = @()

# Loop through the list of Azure users and add them to the variable
foreach ($azure_user in $azure_users) {
  $azure_users_table += [ordered]@{
    DisplayName = $azure_user.DisplayName
    FirstName = $azure_user.FirstName
    LastName = $azure_user.LastName
    Mail = $azure_user.Mail
    Directory = $azure_user.Directory
    Department = $azure_user.Department
    LastSignIn = $azure_user.LastSignIn
    DaysSinceLastSignIn = $azure_user.DaysSinceLastSignIn
    JobTitle = $azure_user.JobTitle
    AccountStatus = $azure_user.AccountStatus
    Groups = $azure_user.Groups
    OfficePhone = $azure_user.OfficePhone
    MobilePhone = $azure_user.MobilePhone
    Notes = $azure_user.Notes
    Created = $azure_user.Created
    Changed = $azure_user.Changed
    Location = $azure_user.Location
    City = $azure_user.City
    AccountLocked = $azure_user.AccountLocked
    AccountExpires = $azure_user.AccountExpires
  }
}

# Export the list of Azure users to a CSV file
Export-Csv -Path "azure_users.csv" -InputObject $azure_users_table -Header UserName,FirstName,LastName,Mail,Directory,Department,LastSignIn,DaysSinceLastSignIn,JobTitle,AccountStatus,Groups,OfficePhone,MobilePhone,Notes,Created,Changed,Location,City,AccountLocked,AccountExpires