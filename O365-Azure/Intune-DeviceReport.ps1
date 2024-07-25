$ver = '2.1'

<#
	Intune-DeviceReport.ps1
	Created By - Kristopher Roy
	Created On - Mar 2022
	Modified On - 25 Jul 2024
#>

#Verify most recent version being used
<#$curver = $ver
$data = Invoke-RestMethod -Method Get -Uri https://raw.githubusercontent.com/BellTechlogix/CreateTabbed-Weekly-Report/master/O365-Azure/Intune-DeviceReport.ps1
Invoke-Expression ($data.substring(0,13))
if($curver -ge $ver){powershell -Command "& {[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); [System.Windows.Forms.MessageBox]::Show('You are running the most current script version $ver')}"}
ELSEIF($curver -lt $ver){powershell -Command "& {[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); [System.Windows.Forms.MessageBox]::Show('You are running $curver the most current script version is $ver. Ending')}" 
EXIT}
#>

#config file
$scriptpath = "C:\Scripts\CreateTabbed-Weekly-Report"
#uncoment if using a config file for variables
[xml]$cfg = Get-Content $scriptpath"\RptCFGFile.xml"
#Organization that the report is for uncomment the cfg file version or static
$org = $cfg.Settings.DefaultSettings.OrgName
#Tenant you are connecting to uncomment the cfg file version or static
$tenant = $cfg.Settings.DefaultSettings.TenantID
$rptFolder = "C:\Projects\GTIL\Reports\"
$runtime = Get-Date -Format "yyyyMMMdd"

# Read the credentials if you have generated for automation otherwise comment out
$accesscreds = Import-Clixml -Path (Join-Path $path "Access.xml")
$ClientID = $accesscreds.ClientID
$clientsecretcreds = ($accesscreds|select Credential).credential


$devtype = @{"Windows"="Workstation/Laptop"; "iOS"="Mobile Device"; "Android"="Mobile Device"}

Disconnect-MgGraph
connect-MgGraph -TenantId $tenant -clientsecretcredential $clientsecretcreds

$devices = Get-MgDeviceManagementManagedDevice|select *
	
$Devices|select @{N='Name';E={$_.deviceName}},@{N='OwnerType';E={$_.ManagedDeviceOwnerType}},@{N='PrimaryUser';E={$_.UserPrincipalName}},EnrolledDateTime,LastSyncDateTime,@{N='dayssinceSync';E={(new-timespan -start (get-date $_.LastSyncDateTime -Hour "00" -Minute "00") -End (get-date -Hour "00" -Minute "00")).Days}},@{N='deviceType';E={$devtype[$_.OperatingSystem]}},Manufacturer,Model,IMEI,SerialNumber,OperatingSystem,OSVersion|export-csv $rptFolder$runtime-IntuneDeviceReport.csv -NoTypeInformation
$Report = $rptFolder+$runtime+'-IntuneDeviceReport.csv'

#Section to generate and send mail report
$powershellVersion = $PSVersionTable.PSVersion
if ($powershellVersion.Major -ge 7){
    $EncodedAttachment = [convert]::ToBase64String((Get-Content $report -AsByteStream))
    Write-Host "PowerShell version is 7 or higher."}else{
    Write-Host "PowerShell version is less than 7."
    $EncodedAttachment = [convert]::ToBase64String((Get-Content $report -Encoding byte)) 
}
    $message = @{
        subject = $org+' - '+$runtime+'-IntuneDeviceReport.csv'
        ToRecipients = @(
                @{
                    EmailAddress = @{
                        Address = "kroy@belltechlogix.com"}
                }
                @{
                    EmailAddress = @{
                        Address = "sgray@belltechlogix.com"}
                }
                @{
                    EmailAddress = @{
                        Address = "twheeler@belltechlogix.com"}
                }
            )
        body = @{
            contentType = 'html'
            content = 'Monthly '+$org+' User Report'
        }
        Attachments = @(
			@{
				"@odata.type" = "#microsoft.graph.fileAttachment"
				name = ($report -split '\\')[-1]
				ContentType = "application/csv"
				ContentBytes = $EncodedAttachment
			}
	    )
    }
#send csv
Send-MgUserMail -Message $message -UserId "Reporting@gti.gt.com"


#cleanup old coppies
$Daysback = '-14'
$CurrentDate = Get-Date
$DateToDelete = $CurrentDate.AddDays($Daysback)
Get-ChildItem $rptFolder | Where-Object { $_.LastWriteTime -lt $DatetoDelete -and $_.Name -like "*IntuneDeviceReport*"} | Remove-Item