$ver = '2.1'

<#
	Intune-DeviceReport.ps1
	Created By - Kristopher Roy
	Created On - Mar 2022
	Modified On - 30 Jan 2024
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
$rptFolder = "C:\Projects\GTIL\Reports\"
$runtime = Get-Date -Format "yyyyMMMdd"
#uncoment if using a config file for variables
#[xml]$cfg = Get-Content $scriptpath"\RptCFGFile.xml"
#Organization that the report is for uncomment the cfg file version or static
#$org = $cfg.Settings.DefaultSettings.OrgName
$org = "GTIL"
#Tenant you are connecting to uncomment the cfg file version or static
#$tenant = $cfg.Settings.DefaultSettings.TenantID
$tenant = "1f05524b-c860-46e3-a2e8-4072506b3e4a"
# Read the credentials if you have generated for automation otherwise comment out
#$Credentials = Import-Clixml -Path $scriptpath\O365-Azure\Access.xml

$devtype = @{"Windows"="Workstation/Laptop"; "iOS"="Mobile Device"; "Android"="Mobile Device"}

#Disconnect-MgGraph
#disconnect-mggraph
#connect-MgGraph
#Uncomment the connection string you need
#connect-MgGraph -TenantId $tenant -clientsecretcredential $credentials
connect-MgGraph -TenantId $tenant

$devices = Get-MgDeviceManagementManagedDevice|select *
	
$Devices|select @{N='Name';E={$_.deviceName}},@{N='OwnerType';E={$_.ManagedDeviceOwnerType}},@{N='PrimaryUser';E={$_.UserPrincipalName}},EnrolledDateTime,LastSyncDateTime,@{N='dayssinceSync';E={(new-timespan -start (get-date $_.LastSyncDateTime -Hour "00" -Minute "00") -End (get-date -Hour "00" -Minute "00")).Days}},@{N='deviceType';E={$devtype[$_.OperatingSystem]}},Manufacturer,Model,IMEI,SerialNumber,OperatingSystem,OSVersion|export-csv $rptFolder$runtime-IntuneDeviceReport.csv -NoTypeInformation

#cleanup old coppies
$Daysback = '-14'
$CurrentDate = Get-Date
$DateToDelete = $CurrentDate.AddDays($Daysback)
Get-ChildItem $rptFolder | Where-Object { $_.LastWriteTime -lt $DatetoDelete -and $_.Name -like "*IntuneDeviceReport*"} | Remove-Item