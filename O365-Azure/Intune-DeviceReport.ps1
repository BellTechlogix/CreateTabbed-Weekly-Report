$ver = '2'

<#
	Intune-DeviceReport.ps1
	Created By - Kristopher Roy
	Created On - Mar 2022
	Modified On - 13 Nov 2023
#>

#Verify most recent version being used
$curver = $ver
$data = Invoke-RestMethod -Method Get -Uri https://raw.githubusercontent.com/BellTechlogix/CreateTabbed-Weekly-Report/master/O365-Azure/Intune-DeviceReport.ps1
Invoke-Expression ($data.substring(0,13))
if($curver -ge $ver){powershell -Command "& {[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); [System.Windows.Forms.MessageBox]::Show('You are running the most current script version $ver')}"}
ELSEIF($curver -lt $ver){powershell -Command "& {[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); [System.Windows.Forms.MessageBox]::Show('You are running $curver the most current script version is $ver. Ending')}" 
EXIT}

#config file
$scriptpath = "C:\Projects\GTIL\Reports\CreateTabbed-Weekly-Report"
[xml]$cfg = Get-Content $scriptpath"\RptCFGFile.xml"
#Organization that the report is for
$org = $cfg.Settings.DefaultSettings.OrgName
$tenant = $cfg.Settings.DefaultSettings.TenantID
# Read the credentials
$Credentials = Import-Clixml -Path $scriptpath\O365-Azure\Access.xml

$devtype = @{"Windows"="Workstation/Laptop"; "iOS"="Mobile Device"; "Android"="Mobile Device"}

#Disconnect-MgGraph
disconnect-mggraph
#connect-MgGraph
connect-MgGraph -TenantId $tenant -clientsecretcredential $credentials

$devices = Get-MgDeviceManagementManagedDevice|select *
#Get-MgDevice -Property *|select *
Foreach($dev in $devices)
{
     
    get-mgdevice -DeviceId $dev.AzureADDeviceId
}
	
$Devices = $Devices|select @{N='Name';E={$_.deviceName}},@{N='OwnerType';E={$_.ManagedDeviceOwnerType}},@{N='PrimaryUser';E={$_.UserPrincipalName}},EnrolledDateTime,LastSyncDateTime,@{N='dayssinceSync';E={(new-timespan -start (get-date $_.LastSyncDateTime -Hour "00" -Minute "00") -End (get-date -Hour "00" -Minute "00")).Days}},@{N='deviceType';E={$devtype[$_.OperatingSystem]}},Manufacturer,Model,IMEI,SerialNumber,OperatingSystem,OSVersion
