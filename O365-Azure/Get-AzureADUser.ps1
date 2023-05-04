<#
	O365User Report.ps1
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

#This function lets you build an array of specific list items you wish
Function MultipleSelectionBox ($inputarray,$prompt,$listboxtype) {
 
# Taken from Technet - http://technet.microsoft.com/en-us/library/ff730950.aspx
# This version has been updated to work with Powershell v3.0.
# Had to replace $x with $Script:x throughout the function to make it work. 
# This specifies the scope of the X variable.  Not sure why this is needed for v3.
# http://social.technet.microsoft.com/Forums/en-SG/winserverpowershell/thread/bc95fb6c-c583-47c3-94c1-f0d3abe1fafc
#
# Function has 3 inputs:
#     $inputarray = Array of values to be shown in the list box.
#     $prompt = The title of the list box
#     $listboxtype = system.windows.forms.selectionmode (None, One, MutiSimple, or MultiExtended)
 
$Script:x = @()
 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
 
$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = $prompt
$objForm.Size = New-Object System.Drawing.Size(300,600) 
$objForm.StartPosition = "CenterScreen"
 
$objForm.KeyPreview = $True
 
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    {
        foreach ($objItem in $objListbox.SelectedItems)
            {$Script:x += $objItem}
        $objForm.Close()
    }
    })
 
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()}})
 
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(75,520)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
 
$OKButton.Add_Click(
   {
        foreach ($objItem in $objListbox.SelectedItems)
            {$Script:x += $objItem}
        $objForm.Close()
   })
 
$objForm.Controls.Add($OKButton)
 
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(150,520)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.Add_Click({$objForm.Close()})
$objForm.Controls.Add($CancelButton)
 
$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,20) 
$objLabel.Size = New-Object System.Drawing.Size(280,20) 
$objLabel.Text = "Please make a selection from the list below:"
$objForm.Controls.Add($objLabel) 
 
$objListbox = New-Object System.Windows.Forms.Listbox 
$objListbox.Location = New-Object System.Drawing.Size(10,40) 
$objListbox.Size = New-Object System.Drawing.Size(260,20) 
 
$objListbox.SelectionMode = $listboxtype
 
$inputarray | ForEach-Object {[void] $objListbox.Items.Add($_)}
 
$objListbox.Height = 470
$objForm.Controls.Add($objListbox) 
$objForm.Topmost = $True
 
$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()
 
Return $Script:x
}

function Get-CurrentUserTenants {
  param (
    [switch]$Export = $false
  )

  # Get the list of tenants that the user has access to
  $tenants = Get-AzTenant | Select-Object Name, ID, @{N = 'Domains';E = {[System.String]::Join(", ", $_.Domains)}}

  # If the Export switch is enabled, export the list of tenants to a CSV file
  if ($Export) {
    Export-Csv -Path "tenants.csv" -InputObject $tenants
  }

  # Return the list of tenants
  return $tenants
}

Install-Module -Name Microsoft.Graph -RequiredVersion 2.0.0-preview8 -AllowPrerelease

#$tenants = Get-CurrentUserTenants

# Increase the function capacity
$MaximumFunctionCount = 32768

# Check if the Microsoft.Graph module is installed
#Check-Module -Name Microsoft.Graph -Import

# Check if the AZ module is installed
Check-Module -Name Az

# Force authentication
#Connect-MgGraph
Connect-AzAccount

$tenants = Get-CurrentUserTenants

$selectedtenet = MultipleSelectionBox -inputarray $tenants.Name -listboxtype One -prompt "Select Your Tenant"
$tenetID = ($tenants|where{$_.name -eq $selectedtenet}).ID
$subscriptions = Get-AzSubscription
$selectedSubscription = MultipleSelectionBox -inputarray ($subscriptions).Name -listboxtype One -prompt "Select Your Subscription"
$subscriptionID = ($subscriptions|where{$_.name -eq $selectedsubscription}).ID
set-azContext -Subscription $SubscriptionID
# Get the list of Azure users
$users = Get-AzAdUser -select 'Department,AccountEnabled,Department,UserType,UserPrincipalName,GivenName,SurName,ApproximateLastSignInDateTime,Manager,JobTitle,Identity,EmployeeType,Department,Country,City,StreetAddress,State,PostalCode,OfficeLocation' -AppendSelected|select *
Get-AzAdUser -select 'Department,AccountEnabled,Department,UserType,UserPrincipalName,GivenName,SurName,ApproximateLastSignInDateTime,Manager,JobTitle,Identity,EmployeeType,Department,Country,City,StreetAddress,State,PostalCode,OfficeLocation' -AppendSelected|select *|where($_.UserType -ne 'Guest')

$users|select UserType

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