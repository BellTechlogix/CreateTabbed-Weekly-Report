$ver = '2'

<#
	Azure Member User Report.ps1
	Created By - Kristopher Roy
	Created On - 27 Oct 2021
	Modified On - 25 Jul 2023

	This Script Requires and O365 account with permissions to connect and pull users. Pulls a report of all O365/Azure User accounts
#>

#Verify most recent version being used
<#
$curver = $ver
$data = Invoke-RestMethod -Method Get -Uri https://raw.githubusercontent.com/BellTechlogix/CreateTabbed-Weekly-Report/master/O365-Azure/Azure-MemberUserReport.ps1
Invoke-Expression ($data.substring(0,13))
if($curver -ge $ver){powershell -Command "& {[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); [System.Windows.Forms.MessageBox]::Show('You are running the most current script version $ver')}"}
ELSEIF($curver -lt $ver){powershell -Command "& {[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); [System.Windows.Forms.MessageBox]::Show('You are running $curver the most current script version is $ver. Ending')}" 
EXIT}
#>

#config file
$scriptpath = "C:\scripts\CreateTabbed-Weekly-Report"
[xml]$cfg = Get-Content $scriptpath"\RptCFGFile.xml"
#Organization that the report is for
$org = $cfg.Settings.DefaultSettings.OrgName
$tenant = $cfg.Settings.DefaultSettings.TenantID
# Read the credentials - Using App Registration 
$Credentials = Import-Clixml -Path $scriptpath\O365-Azure\Access.xml

#Timestamp
$runtime = Get-Date -Format "yyyyMMMdd"

#folder to store completed reports
$rptfolder = $cfg.Settings.DefaultSettings.ReportFolder

#Disconnect-MgGraph
disconnect-mggraph
#connect-MgGraph
connect-MgGraph -TenantId $tenant -clientsecretcredential $credentials

$Users = Get-MgUser -all -Select UserType,SignInActivity,AccountEnabled,City,CompanyName,Country,Department,DisplayName,GivenName,SurName,ID,JobTitle,Mail,MemberOf,State,UserPrincipalName,LastLogonTimestamp,SecurityIdentifier | select-object *,@{N='LastLogonTimestamp';E={ [datetime]$_.SignInActivity.LastSignInDateTime}},@{N='LastNonInteractiveTimestamp';E={ [datetime]$_.SignInActivity.LastNonInteractiveSignInDateTime}}

$users|select DisplayName,@{N='SID';E={($_.SecurityIdentifier)}},givenName,surName,UserPrincipalName,@{N='Domain';E={($_.UserPrincipalName.split('@')[1])}},Department,LastLogonTimestamp,@{N='dayssincelogon';E={(new-timespan -start (get-date $_.LastLogonTimestamp -Hour "00" -Minute "00") -End (get-date -Hour "00" -Minute "00")).Days}},LastNonInteractiveTimestamp,JobTitle,ID,Groups|export-csv $rptFolder$runtime-AzureUserReport.csv -NoTypeInformation

$report = $rptfolder$runtime'-AzureUserReport.csv'
$EncodedAttachment = [convert]::ToBase64String((Get-Content $report -Encoding byte)) 


$message = @{
        subject = $org' Azure User Report'
        ToRecipients = @(
                @{
                    EmailAddress = @{
                        Address = "kroy@belltechlogix.com"}
                }
                @{
                    EmailAddress = @{
                        Address = "nick.kowalski@gti.gt.com"}
                }
                @{
                    EmailAddress = @{
                        Address = "cmills@belltechlogix.com"}
                }
            )
        body = @{
            contentType = 'html'
            content = 'Testing'
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



Send-MgUserMail -Message $message -UserId "Reporting@gti.gt.com"


#cleanup old coppies
$Daysback = '-14'
$CurrentDate = Get-Date
$DateToDelete = $CurrentDate.AddDays($Daysback)
Get-ChildItem $rptFolder | Where-Object { $_.LastWriteTime -lt $DatetoDelete -and $_.Name -like "*AzureUserReport*"} | Remove-Item
