<#
	Quest-AllComputers-Report.ps1
	Created By - Kristopher Roy
	Created On - May 2017
	Modified On - 28 Apr 2021

	This Script Requires that the Quest_ActiveRolesManagementShellforActiveDirectory be installed https://www.powershelladmin.com/wiki/Quest_ActiveRoles_Management_Shell_Download
	Pulls a report of all Systems in an Active Directory Structure as defined by Domain Root
#>


add-pssnapin quest.activeroles.admanagement
Import-Module activedirectory

#config file
$scriptpath = "F:\Scripts"
[xml]$cfg = Get-Content $scriptpath"\RptCFGFile.xml"

#Organization that the report is for
$org = $cfg.Settings.DefaultSettings.OrgName

#modify this for your searchroot can be as broad or as narrow as you need down to OU
$domainRoot = $cfg.Settings.DefaultSettings.DomainRoot
$DC1 = $cfg.Settings.DefaultSettings.DC

#If multiple domains uncoment and use
#$domainRoot2 = $cfg.Settings.DefaultSettings.DomainRoot2
#$DC2 = $cfg.Settings.DefaultSettings.DC2

#folder to store completed reports
$rptfolder = $cfg.Settings.DefaultSettings.ReportFolder

#mail recipients for sending report
$recipients = @("BTL SCCM <sccm@belltechlogix.com>","BTL ITAMS <ITAM@belltechlogix.com>")

#from address
$from = $cfg.Settings.EmailSettings.FromAddress

#smtpserver
$smtp = $cfg.Settings.EmailSettings.SMTPServer

#Timestamp
$runtime = Get-Date -Format "yyyyMMMdd"

#deffinition for UAC codes
$lookup = @{4096="Workstation/Server"; 4098="Disabled Workstation/Server"; 4128="Workstation/Server No PWD"; 
4130="Disabled Workstation/Server No PWD"; 528384="Workstation/Server Trusted for Delegation"; 83955712="Workstation/Server Partial Secrests Account/Trusted For Delegation/PWD not Expire";
528416="Workstation/Server Trusted for Delegation"; 532480="Domain Controller"; 66176="Workstation/Server PWD not Expire"; 
66178="Disabled Workstation/Server PWD not Expire";512="User Account";514="Disabled User Account";66048="User Account PWD Not Expire";66050="Disabled User Account PWD Not Expire"}

#creates the report folder if it doesn't exist
if(!(Test-Path -Path $rptfolder)){
    New-Item -ItemType directory -Path $rptfolder
}

$qadcomputers = Get-QADComputer -searchroot $domainRoot  -Service $DC1 -searchscope subtree -sizelimit 0 -includedproperties name,userAccountControl,whenCreated,whenChanged,lastlogondate,dayssincelogon,lastlogontimestamp,description,operatingSystem,operatingsystemservicepack|Select-Object -Property name,@{N='Domain';E={("$_.domain".split('\')[0])}},@{N='OU';E={($_.canonicalName -split('/')|select-object -skip 1 -Last 10) -Join "/"}},lastlogontimestamp,@{N='dayssincelogon';E={(new-timespan -start (get-date $_.LastLogonTimestamp -Hour "00" -Minute "00") -End (get-date -Hour "00" -Minute "00")).Days}},@{N='userAccountControl';E={$lookup[$_.userAccountControl]}},whenCreated,whenChanged,description,operatingSystem,operatingSystemVersion,operatingsystemservicepack|sort name
#If multiple Domains uncomment and use
#$qadcomputers += Get-QADComputer -searchroot $domainRoot2 -Service $DC2 -searchscope subtree -sizelimit 0 -includedproperties name,userAccountControl,whenCreated,whenChanged,lastlogondate,dayssincelogon,lastlogontimestamp,description,operatingSystem,operatingsystemservicepack|Select-Object -Property name,@{N='Domain';E={("$_.domain".split('\')[0])}},@{N='OU';E={($_.canonicalName -split('/')|select-object -skip 1 -Last 10) -Join "/"}},lastlogontimestamp,@{N='dayssincelogon';E={(new-timespan -start (get-date $_.LastLogonTimestamp -Hour "00" -Minute "00") -End (get-date -Hour "00" -Minute "00")).Days}},@{N='userAccountControl';E={$lookup[$_.userAccountControl]}},whenCreated,whenChanged,description,operatingSystem,operatingSystemVersion,operatingsystemservicepack|sort name

$qadcomputers|export-csv $rptfolder$runtime-qAD-AllComputerReport.csv -NoTypeInformation

#If Running standalone un-comment to send mail
<#
	$emailBody = "<h1>$org Weekly All Workstations Report</h1>"
	$emailBody = $emailBody + "<p><em>"+(Get-Date -Format 'MMM dd yyyy HH:mm')+"</em></p>"

	$htmlforEmail = $emailBody + @'
	<h2>Included Fields:</h2>
	<table style="height: 535px;" border="1" width="625">
	<tbody>
	<tr style="height: 47px;">
	<td style="width: 304px; height: 25px;"><strong>name</strong></td>
	<td style="width: 305px; height: 25px;"><em>&nbsp;Computer Name</em></td>
	</tr>
	<tr style="height: 47px;">
	<td style="width: 304px; height: 25px;"><strong>lastLogonTimestamp</strong></td>
	<td style="width: 305px; height: 25px;"><em>Last Recorded Timestamp for a logon</em></td>
	</tr>
	<tr style="height: 47px;">
	<td style="width: 304px; height: 25px;"><strong>dayssincelogon</strong></td>
	<td style="width: 305px; height: 25px;"><em>calculated from lastlogontimestamp</em></td>
	</tr>
	<tr style="height: 47px;">
	<td style="width: 304px; height: 25px;"><strong>userAccountControl</strong></td>
	<td style="width: 305px; height: 25px;"><em>User/Computer settings for AD</em></td>
	</tr>
	<tr style="height: 47px;">
	<td style="width: 304px; height: 25px;"><strong>whenCreated</strong></td>
	<td style="width: 305px; height: 25px;"><em>When account was created</em></td>
	</tr>
	<tr style="height: 29px;">
	<td style="width: 304px; height: 25px;"><strong>whenChanged</strong></td>
	<td style="width: 305px; height: 25px;"><em>Date AD changes were made to account</em></td>
	</tr>
	<tr style="height: 10px;">
	<td style="width: 304px; height: 25px;"><strong>description</strong></td>
	<td style="width: 305px; height: 25px;"><em>Description field from AD if populated</em></td>
	</tr>
	<tr style="height: 10px;">
	<td style="width: 304px; height: 25px;"><strong>operatingSystem</strong></td>
	<td style="width: 305px; height: 25px;"><em>&nbsp;OS Name</em></td>
	</tr>
	<tr style="height: 1px;">
	<td style="width: 304px; height: 25px;"><strong>operatingSystemVersion</strong></td>
	<td style="width: 305px; height: 25px;"><em>Version number of OS</em></td>
	</tr>
	<tr style="height: 24.3594px;">
	<td style="width: 304px; height: 25px;"><strong>operatingSystemServicePack</strong></td>
	<td style="width: 305px; height: 25px;"><em>OS Service Pack installed, if any</em></td>
	</tr>
	</tbody>
	</table>
	'@

	Send-MailMessage -from $from -to $recipients -subject "$org All Workstations Report" -smtpserver $smtp -BodyAsHtml $htmlforEmail -Attachments $rptfolder$runtime-qAD-AllComputerReport.csv
#>

#Cleanup Old Files
$Daysback = '-14'
$CurrentDate = Get-Date
$DateToDelete = $CurrentDate.AddDays($Daysback)
Get-ChildItem $rptFolder | Where-Object { $_.LastWriteTime -lt $DatetoDelete -and $_.Name -like "*qAD-AllComputerReport*"} | Remove-Item