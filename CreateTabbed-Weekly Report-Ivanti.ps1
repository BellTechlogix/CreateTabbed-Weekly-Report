<#
	CreateTabbed-Weekly Report-Ivanti.ps1
	****Ivanti Version****
	Created By - Kristopher Roy
	Created On - Feb 21 2020
	Modified On - Apr 12 2021

	This Script combines multiple reports into a single tabbed report
#>


#Organization that the report is for
$org = "MyCompany"

#folder to store completed reports
$rptfolder = "c:\reports\"

#mail recipients for sending report
$recipients = @("BTL SCCM <sccm@belltechlogix.com>","BTL ITAMS <ITAM@belltechlogix.com>")

#from address
$from = "Reports@wherever.com"

#smtpserver
$smtp = "mail.wherever.com"

#Timestamp
$runtime = Get-Date -Format "yyyyMMMdd"

#XMLFile for output
$XMLFile = $rptFolder+"ConsolidatedReport.xml"

#report1 45Day Computer Report to import
$tab1 = $rptfolder+$runtime+"-qADComputerReport-45.csv"

#csv 1 45Day Computer Report to import
$qad45report = import-csv $tab1
$qad45reportcount = $qad45report.count

#report2 All Computer Report to import
$tab2 = $rptfolder+$runtime+"-qAD-AllComputerReport.csv"

#csv 2 All Computer Report to import
$qadallsys = import-csv $tab2
$qadallsyscount = $qadallsys.count

#report3 Ivanti detailed Report to import
$tab3 = $rptfolder+"IvantiAll-Weekly.csv"

#csv 3 Ivanti detailed Report to import
$Ivantisys = import-csv $tab3
$Ivantisyscount = $Ivantisys.count

#report4 ADUser Report to import
$tab4 = $rptfolder+$runtime+"-qADUserReport.csv"

#csv 3 Ivanti detailed Report to import
$adusers = import-csv $tab4
$aduserscount = $adusers.count

#Zip File Function
function ZipFiles( $zipfilename, $sourcedir )
{
   Add-Type -Assembly System.IO.Compression.FileSystem
   $compressionLevel = [System.IO.Compression.CompressionLevel]::Optimal
   [System.IO.Compression.ZipFile]::CreateFromDirectory($sourcedir,
        $zipfilename, $compressionLevel, $false)
}

#Lets create our XML File, this is the initial formatting that it will need to understand what it is, and what styles we are using.
(
 '<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
xmlns:html="http://www.w3.org/TR/REC-html40">
<DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
<Author>Kristopher Roy</Author>
<LastAuthor>'+$env:USERNAME+'</LastAuthor>
<Created>'+(get-date)+'</Created>
<Version>16.00</Version>
</DocumentProperties>
<OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
<AllowPNG/>
</OfficeDocumentSettings>
 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
  <WindowHeight>7920</WindowHeight>
  <WindowWidth>25530</WindowWidth>
  <WindowTopX>32767</WindowTopX>
  <WindowTopY>32767</WindowTopY>
  <ProtectStructure>False</ProtectStructure>
  <ProtectWindows>False</ProtectWindows>
 </ExcelWorkbook>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
  <Style ss:ID="s62">
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#FFFFFF"
    ss:Bold="1"/>
   <Interior ss:Color="#4472C4" ss:Pattern="Solid"/>
  </Style>
    <Style ss:ID="s63">
    <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
  </Style>
    <Style ss:ID="s64">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
   <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
  </Style>
    <Style ss:ID="s65">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s66">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
   <Interior ss:Color="#00B050" ss:Pattern="Solid"/>
  </Style>
 </Styles>')>$XMLFile

#Tab1 Report
add-content $XMLFile (
 '<Worksheet ss:Name="'+($runtime)+'-ADComputerReport-45">
  <Table ss:ExpandedColumnCount="10" ss:ExpandedRowCount="'+($qad45reportcount+1)+'" x:FullColumns="1"
   x:FullRows="1" ss:DefaultRowHeight="15">
   <Column ss:AutoFitWidth="0" ss:Width="119.25"/>
   <Column ss:Width="111.75"/>
   <Column ss:Width="77.25"/>
   <Column ss:Width="99"/>
   <Column ss:AutoFitWidth="0" ss:Width="111.75" ss:Span="1"/>
   <Column ss:Index="7" ss:Width="58.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="109.5"/>
   <Column ss:Width="122.25"/>
   <Column ss:Width="141.75"/>
   <Row ss:AutoFitHeight="0">
    <Cell ss:StyleID="s62"><Data ss:Type="String">Name</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">lastLogonTimestamp</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">dayssincelogon</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">userAccountControl</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">whenCreated</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">whenChanged</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">Description</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">operatingSystem</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">operatingSystemVersion</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">operatingSystemServicePack</Data></Cell>
   </Row>')
   FOREACH($system in $qad45report)
   {
   add-content $XMLFile ('
      <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="String">'+($system.name)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system.lastLogonTimestamp)+'</Data></Cell>
    <Cell ss:StyleID="s63"><Data ss:Type="Number">'+($system.dayssincelogon)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system.userAccountControl)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system.whenCreated)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system.whenChanged)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system.Description)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system.operatingSystem)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system.operatingSystemVersion)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system.operatingSystemServicePack)+'</Data></Cell>
   </Row>
   ')
   }
   $system = $null

#Tab2 Report
  add-content $XMLFile ('</Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Header x:Margin="0.3"/>
    <Footer x:Margin="0.3"/>
    <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>
   </PageSetup>
   <Unsynced/>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
 <Worksheet ss:Name="'+($runtime)+'-AD-AllSystems">
  <Table ss:ExpandedColumnCount="10" ss:ExpandedRowCount="'+($qadallsyscount+1)+'" x:FullColumns="1"
   x:FullRows="1" ss:DefaultRowHeight="15">
   <Column ss:AutoFitWidth="0" ss:Width="119.25"/>
   <Column ss:Width="111.75"/>
   <Column ss:Width="77.25"/>
   <Column ss:Width="99"/>
   <Column ss:AutoFitWidth="0" ss:Width="111.75" ss:Span="1"/>
   <Column ss:Index="7" ss:Width="58.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="109.5"/>
   <Column ss:Width="122.25"/>
   <Column ss:Width="141.75"/>
   <Row ss:AutoFitHeight="0">
    <Cell ss:StyleID="s62"><Data ss:Type="String">Name</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">lastLogonTimestamp</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">dayssincelogon</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">userAccountControl</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">whenCreated</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">whenChanged</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">Description</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">operatingSystem</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">operatingSystemVersion</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">operatingSystemServicePack</Data></Cell>
   </Row>')
      FOREACH($system in $qadallsys)
   {
   add-content $XMLFile ('
      <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="String">'+($system.name)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system.lastLogonTimestamp)+'</Data></Cell>')
    If([int]$system.dayssincelogon -gt 90){add-content $XMLFile ('<Cell ss:StyleID="s64"><Data ss:Type="Number">'+($system.dayssincelogon)+'</Data></Cell>')}
    ElseIF([int]$system.dayssincelogon -gt 45 -and [int]$system.dayssincelogon -lt 90){add-content $xmlfile ('<Cell ss:StyleID="s65"><Data ss:Type="Number">'+($system.dayssincelogon)+'</Data></Cell>')}
    ElseIF([int]$system.dayssincelogon -lt 45){add-content $xmlfile ('<Cell ss:StyleID="s66"><Data ss:Type="Number">'+($system.dayssincelogon)+'</Data></Cell>')}
    add-content $xmlfile('
    <Cell><Data ss:Type="String">'+($system.userAccountControl)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system.whenCreated)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system.whenChanged)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system.Description)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system.operatingSystem)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system.operatingSystemVersion)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system.operatingSystemServicePack)+'</Data></Cell>
   </Row>')
   }
   $system = $null

#Tab3 Report
     add-content $XMLFile ('</Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Header x:Margin="0.3"/>
    <Footer x:Margin="0.3"/>
    <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>
   </PageSetup>
   <Unsynced/>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
 <Worksheet ss:Name="'+($runtime)+'-Ivanti-Detailed">
  <Table ss:ExpandedColumnCount="11" ss:ExpandedRowCount="'+($Ivantisyscount+1)+'" x:FullColumns="1"
   x:FullRows="1" ss:DefaultRowHeight="15">
   <Column ss:AutoFitWidth="0" ss:Width="119.25"/>
   <Column ss:Width="111.75"/>
   <Column ss:AutoFitWidth="0" ss:Width="100.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="133.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="79.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="75"/>
   <Column ss:AutoFitWidth="0" ss:Width="69"/>
   <Column ss:AutoFitWidth="0" ss:Width="71.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="55.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="229.5"/>
   <Row ss:AutoFitHeight="0">
    <Cell ss:StyleID="s62"><Data ss:Type="String">Computer</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">SerialNumber</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">Scan Date</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">Last User</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">Operating System</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">Agent</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">Type</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">IP</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">DeviceID</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">Has Client</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">Client Version</Data></Cell>
   </Row>')

   FOREACH($system in $Ivantisys)
   {
   add-content $XMLFile ('
      <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="String">'+($system."Computer Name")+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system."SerialNumber")+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system."Scan Date")+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system."Last User")+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system."Operating System")+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system."Agent Configuration Name")+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system.Type)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system.IP)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system.DeviceID)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system.IsClient)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($system."Client Version")+'</Data></Cell>
   </Row>')
   }
  $system = $null

#Tab4 Report
     add-content $XMLFile ('</Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Header x:Margin="0.3"/>
    <Footer x:Margin="0.3"/>
    <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>
   </PageSetup>
   <Unsynced/>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
 <Worksheet ss:Name="'+($runtime)+'-ADUsers">
  <Table ss:ExpandedColumnCount="10" ss:ExpandedRowCount="'+($ADUserscount+1)+'" x:FullColumns="1"
   x:FullRows="1" ss:DefaultRowHeight="15">
   <Column ss:AutoFitWidth="0" ss:Width="119.25"/>
   <Column ss:Width="111.75"/>
   <Column ss:Width="77.25"/>
   <Column ss:Width="99"/>
   <Column ss:AutoFitWidth="0" ss:Width="111.75" ss:Span="1"/>
   <Column ss:Index="7" ss:Width="58.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="109.5"/>
   <Column ss:Width="122.25"/>
   <Column ss:Width="141.75"/>
   <Row ss:AutoFitHeight="0">
    <Cell ss:StyleID="s62"><Data ss:Type="String">DisplayName</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">SamAccountName</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">givenName</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">surName</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">UserPrincipalName</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">LastLogonTimestamp</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">dayssincelogon</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">employeeType</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">userAccountControl</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">Groups</Data></Cell>
   </Row>')

   FOREACH($user in $adusers)
   {
   add-content $XMLFile ('
      <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="String">'+($user.DisplayName)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($user.SamAccountName)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($user.givenName)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($user.sn)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($user.UserPrincipalName)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($user.LastLogonTimestamp)+'</Data></Cell>    
    ')
    If([int]$user.dayssincelogon -gt 90){add-content $XMLFile ('<Cell ss:StyleID="s64"><Data ss:Type="Number">'+($user.dayssincelogon)+'</Data></Cell>')}
    ElseIF([int]$user.dayssincelogon -gt 45 -and [int]$user.dayssincelogon -lt 90){add-content $xmlfile ('<Cell ss:StyleID="s65"><Data ss:Type="Number">'+($user.dayssincelogon)+'</Data></Cell>')}
    ElseIF([int]$user.dayssincelogon -lt 45){add-content $xmlfile ('<Cell ss:StyleID="s66"><Data ss:Type="Number">'+($user.dayssincelogon)+'</Data></Cell>')}
    add-content $xmlfile('
    <Cell><Data ss:Type="String">'+($user.employeeType)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($user.userAccountControl)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($user.Groups)+'</Data></Cell>
   </Row>')
   }

   add-content $XMLFile ('</Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Header x:Margin="0.3"/>
    <Footer x:Margin="0.3"/>
    <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>
   </PageSetup>
   <Selected/>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
</Workbook>')


$emailBody = "<h1>$org Weekly Consolidated Report</h1>"
$emailBody = $emailBody + "<h2>$org 45 Day Machine Count - '$qad45reportcount'</h2>"
$emailBody = $emailBody + "<h2>$org All Machine Count - '$qadallsyscount'</h2>"
$emailBody = $emailBody + "<h2>$org All Users Count - '$aduserscount'</h2>"
$emailBody = $emailBody + "<p><em>"+(Get-Date -Format 'MMM dd yyyy HH:mm')+"</em></p>"

if(test-path $rptFolder"ConsolidatedReport.zip"){del $rptFolder"ConsolidatedReport.zip"}
Compress-Archive $rptFolder"ConsolidatedReport.xml" -DestinationPath $rptFolder"ConsolidatedReport.zip"
Send-MailMessage -from $from -to $recipients -subject "$org - Consolidated Weekly Report" -smtpserver $smtp -BodyAsHtml $emailbody -Attachments $rptFolder"ConsolidatedReport.zip"

#cleanup old coppies
$Daysback = '-30'
$CurrentDate = Get-Date
$DateToDelete = $CurrentDate.AddDays($Daysback)
Get-ChildItem $rptFolder | Where-Object { $_.LastWriteTime -lt $DatetoDelete -and $_.Name -like "*Consolidated*"} | Remove-Item