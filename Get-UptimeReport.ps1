<#    
.SYNOPSIS    
    
  Get Windows Installed Drive Space information and Server Uptime from a list of servers, then compiles the data in HTML format and sends that Report Over Mail. 
    
.DESCRIPTION    
     
  This scripts will export the Windows Installed Drive Space information i.e. total drive space and free space as well as Server Uptime from the list of given servers. Later, compiles the data into HTML format and sends that Report Over the Mail.    
    
.COMPATIBILITY     
     
  Requires PS v4. Tested against Windows 2012, 2012 R2 and 2016 servers 
      
.EXAMPLE  
  PS C:\> Get-UptimeReport.ps1
  All options are set as variables in the GLOBALS section so you simply run the script.  
  
.NOTES    
      
 1) Remote Execution of script should be enabled.
 2) The account running the script or scheduled task obviously must have the appropriate permissions on each server. 
 3) Create below folder hierarchy in any drive, where you will run this script                                                         
          UptimeReport                                                                    
                01_HTML_Report                                                               
                02_BackupLocation
                03_ZIP_Report
  
 4) SMTP Server information
 5) Author Information
    Script NAME        :   Get-UptimeReport.ps1    
    AUTHOR             :   Avijit Dutta
    CREATED            :   31/Jan/2020
    Website            :   www.virtualgyanis.com
    Disclaimer         :   You can use this Script as a reference and can modify, if required as per you need.
     
#>    
  
###################################################################################################################################### 
########################################################### GLOBALS Values ###########################################################  
######################################################################################################################################
  
######################################################################################################################################
########################################################### Server Listing ###########################################################
######################################################################################################################################
# Provide all the Server hostname into this variable. There are multiple ways to put the server information. I preferred manually. I will mentioned other ways as well.
# 1) Manual Server List

$ServerList = 'server1','server2','server3'

# Fetch all the enabled Server account from Active Directory and put it into $Serverlist Variable
# $ServerList = Get-ADComputer -Filter {(OperatingSystem -like "*windows*server*") -and (Enabled -eq "True")} -Properties OperatingSystem 

# Put the server hostname from the .TXT File.
# $ServerList = get-content -Path "Path of the .TXT File"
######################################################################################################################################
##################################################### Set the Output Directories######################################################
######################################################################################################################################
# Output Directory Path and Filename of the Exported data

$HTMLReportDIR = "E:\Script\UptimeReport\01_HTML_Report\"           # Directory Path for HTML Report
$BackupLocationDIR = "E:\Script\UptimeReport\02_BackupLocation\"    # Directory Path for the Backup of the File
$ZIPReportDIR = "E:\Script\UptimeReport\03_ZIP_Report\"             # Directory path to ZIP the file

$ScriptDIR = $HTMLReportDIR, $ZIPReportDIR, $BackupLocationDIR

# Check if the above path exist of not, if not it will automatically create the same.

Foreach ($DIR in $ScriptDIR)
{
if(!(Test-Path -Path $DIR )){
    New-Item -ItemType directory -Path $DIR
    Write-Host "Folder created"
}
else
{
  Write-Host "Folder already exists"
}
}


# Output file name with current date and time.
$OutputFile = "E:\Script\UptimeReport\01_HTML_Report\UptimeReport_$(get-date -f yyyy-MM-dd-HHmm).htm"  

######################################################################################################################################
##################################################### Get Current Date and Time ######################################################
######################################################################################################################################

# Get the current date and Time
$DateTime =  Get-Date -f F 

###################################################################################################################################### 
##################################################### For Logging and Debugging ###################################################### 
######################################################################################################################################

# Optional - To create transcript - change to $true to turn on. This will log all the activity
$CreateTranscript = $false 
 
# Start Transcript if $CreateTranscript variable above is set to $true. 
if($CreateTranscript) 
{ 
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path 
if( -not (Test-Path ($scriptDir + "\Transcripts"))){New-Item -ItemType Directory -Path ($scriptDir + "\Transcripts")} 
Start-Transcript -Path ($scriptDir + "\Transcripts\{0:yyyyMMdd}_Log.txt"  -f $(get-date)) -Append 
} 
  
 
######################################################################################################################################  
#################################################### Main Script with HTML Codding ###################################################
######################################################################################################################################  
 
$HTML = '<style type="text/css">  
#TSHead body {font: normal small sans-serif;} 
#TSHead table {border-collapse: collapse;background-color:#3f3d3d;} 
#TSHead th {font-family: Verdana;font-size:80%;padding: 10px 10px;text-align: Justify-all;border: 1px solid #3f3d3d;background-color:#7FB1B3;} 
#TSHead td {font-family: Verdana;font-size:70%;padding: 5px;text-align: justify-all;border: 1px solid #3f3d3d;} 
#TSHead tbody tr:nth-child(odd) {background: #D3D3D3;} 
    </Style>'  
 
# Report Header 
$Header = "<H5 Align=left><font face=Arial>Date & Time :: $datetime </font></H5>" 
$Header1 = "<H2 Align=center><font face=Arial>Server Uptime and Windows Drive Space Report</font></H2>" 
#$Header2 = "<H4 align=center><font face=Arial><span style=background-color:#FFF284>WARNING</span> at <span style=background-color:#FFF284>80%</span> in use. <span style=background-color:#FF9797>CRITICAL</span> at <span style=background-color:#FF9797>95%</span> in use.</font></H4>"  
 
$HTML += "<HTML><BODY><Table border=2 cellpadding=0 cellspacing=0 width=100% id=TSHead class=sortable> 
        <TR>  
            <TH><B>Server Hostname</B></TH> 
            <TH><B>Audit Date</B></TH> 
            <TH><B>Windows Drive</B></TH> 
            <TH><B>Total Size (GB)</B></TH> 
            <TH><B>Free Space (GB)</B></TH> 
            <TH><B>Last Reboot Time</B></TH>        
        </TR> 
        "  
 
Foreach($Server in $ServerList) 
{
$os = Get-WmiObject win32_OperatingSystem -computername $Server
$reboot = Get-CimInstance Win32_OperatingSystem -ComputerName $Server
$DriveInfo = Get-WMIObject Win32_Logicaldisk -filter "deviceid='$($os.systemdrive)'" -ComputerName $server | Select @{Name="Drive Letter";Expression={$_.DeviceID}},@{Name="Total Size (GB)";Expression={$_.Size/1GB -as [int]}},@{Name="Free Space (GB)";Expression={[math]::Round($_.Freespace/1GB,2)}}
$RebootTime = New-TimeSpan -Start $Reboot.LastBootUpTime -End $DateTime | Select-Object @{Label = "Last Reboot Time"; Expression = {$Reboot.LastBootUpTime}},Days,Hours,Minutes,Seconds
     
 
# HTML Table values     
                $HTML += "<TR> 
                    <TD>$($Server)</TD> 
                    <TD>$($DateTime)</TD> 
                    <TD>$($DriveInfo.'Drive Letter')</TD> 
                    <TD>$($DriveInfo.'Total Size (GB)')</TD> 
                    <TD>$($DriveInfo.'Free Space (GB)')</TD> 
                    <TD>$($RebootTime.'Last Reboot Time')</TD> 
                    </TR>" 
} 

 
$HTML += "<H2></Table></BODY></HTML>"  
$Header + $Header1 + $HTML | Out-File $OutputFile 
 
######################################################################################################################################  
########################################## Zip the export file and send the same over email. ######################################### 
###################################################################################################################################### 
 
## Command will now copy the file from "01_HTML_Report" folder to Backup Location. 
 
Copy-Item -Path $OutputFile -Destination "E:\Script\UptimeReport\02_BackupLocation" 
 
## Now we will ZIP the report and copy the same into 03_ZIP_Report folder

$ZIPSource = "E:\Script\UptimeReport\02_BackupLocation" 
$ZIPDestination = "E:\Script\UptimeReport\03_ZIP_Report\UptimeReport_$(get-date -f yyyy-MM-dd-HHmm).zip" 
If(Test-path $ZIPDestination) {Remove-item $ZIPDestination} 
Add-Type -assembly "system.io.compression.filesystem" 
[io.compression.zipfile]::CreateFromDirectory($ZIPSource, $ZIPDestination)  
 
###################################################################################################################################### 
################################################### SMTP Settings for Email Report ################################################### 
###################################################################################################################################### 

## Specify the "From" email address 
 
$EmailFrom = "Report@abc.com" 
 
## Specify the "To" email address/addressess

$EmailTo1 = "TOrecipient1@abc.com" 
$EmailTo2 = "TOrecipient2@abc.com" 
$EmailTo3 = "TOrecipient3@abc.com" 
$EmailTo4 = "TOrecipient4@abc.com" 
$EmailTo5 = "TOrecipient5@abc.com" 
 
## Specify the "CC" email address/addressess

$EmailCC = "CCrecipient1@abc.com" 
 
## Specify the Subject of the email 

$EmailSubject = "Server Uptime Report :: Dated :: $(get-date -f yyyy-MM-dd)"  
 
## Specify the Body of the email  

$EmailBody = @" 
<html> 
<body style="font-family: Arial; font-size: 11pt;"> 
Dear All,<br></br> 
    Please find attached Server Uptime and Windows Installed Drive Space Information - Dated :: $(get-date -f yyyy-MM-dd).<br></br> 
<br></br> 
Regards,<br> 
Server Team</br> 
</body> 
</html> 
<br></br> 
 
"@ 
$attachment = $ZIPDestination 
$SMTPServer = "smtp.exchangeserver.com" 
$SMTPPort = "25" 
 
$Message = New-Object System.Net.Mail.MailMessage 
$Message.From = $EmailFrom 
$Message.To.Add($EmailTo1) 
$Message.To.Add($EmailTo2) 
$Message.To.Add($EmailTo3) 
$Message.To.Add($EmailTo4) 
$Message.To.Add($EmailTo5) 
$Message.CC.Add($EmailCC) 
$Message.Body = $EmailBody 
$message.Subject = $EmailSubject 
$attach = new-object Net.Mail.Attachment($attachment) 
$message.Attachments.Add($attach) 
$message.IsBodyHTML = $true 
$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, $SMTPPort) 
$SMTPClient.Send($Message) 
$Message.dispose() 
 
## Below command will delete the extra file from Backup Location. 
 
Get-ChildItem “E:\Script\UptimeReport\02_BackupLocation” -recurse -include *.htm -force | remove-item

######################################################################################################################################
###################################### END OF SCRIPT (© www.VirtualGyanis.com) #######################################################
######################################################################################################################################