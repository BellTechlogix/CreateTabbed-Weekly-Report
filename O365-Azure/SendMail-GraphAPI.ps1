<#
	SendMail-GraphAPI.ps1
	Created By - Kristopher Roy
	Created On - 23 May 2023
	Modified On - 23 May 2023

	This Script Requires an Azure account or an App Registration with permissions to connect and sendmail
#>

# Read the credentials
$path = 'C:\Users\tankcr\source\repos\BellTechlogix\CreateTabbed-Weekly-Report\O365-Azure'
$report = 'C:\Projects\GTIL\Reports\GTILAADDevices.csv'
$EncodedAttachment = [convert]::ToBase64String((Get-Content $report -Encoding byte)) 
$Credentials = Import-Clixml -Path $Path\Access.xml

#Disconnect-MgGraph
connect-MgGraph -TenantId '1f05524b-c860-46e3-a2e8-4072506b3e4a' -clientsecretcredential $credentials


    $message = @{
        subject = 'test automation'
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