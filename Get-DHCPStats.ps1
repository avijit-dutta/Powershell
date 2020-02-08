<#    
.SYNOPSIS    
    
  Gets DHCP scopes from a list of servers, compiles the data in HTML and sends that HTML Report Over Mail. 
    
.DESCRIPTION    
     
  This script gets DHCP Scope info from a list of servers, compiles the data in HTML, changes the cell color based on percentage in use and sends an email the HTML report.   
    
.COMPATIBILITY     
     
  Requires PS v4. Tested against 2012 and 2012 R2 DHCP servers 
      
.EXAMPLE  
  PS C:\> Get-DHCPStats.ps1  
  All options are set as variables in the GLOBALS section so you simply run the script.  
  
.NOTES    
      
 1) This script requires the DhcpServer module. 
 2) The account running the script or scheduled task obviously must have the appropriate permissions on each server. 
 3) Create below folder hierarchy                                                          
          DHCP_STAT                                                                    
                01_HTML_Report                                                                  
                02_BackupLocation                                                           
                03_ZIP_Report 
  
 4) SMTP Server information 
 5) This script also requires the sortable.js script if you want to make the table columns sortable. Download the script and place it in the same directory as $OutputFile. Get it here: http://www.kryogenix.org/code/browser/sorttable  
    
  NAME               :   Get-DHCPStats.ps1    
  Original AUTHOR    :   Brian D. Arnold 
  Modified           :   Avijit Dutta - I have modified the script as per my need and environment. 
  CREATED            :   07/02/2014   
  LASTEDIT           :   09/02/2017   
#>    
  
###################  
##### GLOBALS #####  
###################  
 
# Get all Authorized DCs from AD configuration with below command. But, I have mentioned them manually. 
# $DHCPs = Get-DhcpServerInDC | Select DnsName 
 
# Manually, mentioned the DHCP Server list  
$DHCPs = 'Server1','Server2','Server3','Server4' 
 
# Path and Filename of the Exported data 
$OutputFile = "E:\Script\DHCP_STAT\01_HTML_Report\DHCP_Stat_$(get-date -f yyyy-MM-dd-HHmm).htm"  
  
# Get domain name, date and time for report title  
$DomainName = (Get-ADDomain).NetBIOSName   
$DateTime =  Get-Date -f F 
 
# THreshold variables 
$Alert = '95' 
$Warn = '80' 
 
# Option to create transcript - change to $true to turn on. 
$CreateTranscript = $false 
 
############### 
##### PRE ##### 
############### 
 
# Start Transcript if $CreateTranscript variable above is set to $true. 
if($CreateTranscript) 
{ 
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path 
if( -not (Test-Path ($scriptDir + "\Transcripts"))){New-Item -ItemType Directory -Path ($scriptDir + "\Transcripts")} 
Start-Transcript -Path ($scriptDir + "\Transcripts\{0:yyyyMMdd}_Log.txt"  -f $(get-date)) -Append 
} 
  
# Import modules 
Import-Module DhcpServer 
 
################  
##### MAIN #####  
################  
 
$HTML = '<style type="text/css">  
#TSHead body {font: normal small sans-serif;} 
#TSHead table {border-collapse: collapse;background-color:#3f3d3d;} 
#TSHead th {font-family: Verdana;font-size:80%;padding: 10px 10px;text-align: Justify-all;border: 1px solid #3f3d3d;background-color:#7FB1B3;} 
#TSHead td {font-family: Verdana;font-size:70%;padding: 5px;text-align: justify-all;border: 1px solid #3f3d3d;} 
#TSHead tbody tr:nth-child(odd) {background: #D3D3D3;} 
    </Style>'  
 
# Report Header 
$Header = "<H5 Align=left><font face=Arial>Date & Time :: $datetime </font></H5>" 
$Header1 = "<H2 Align=center><font face=Arial>$DomainName DHCP Statistics</font></H2>" 
$Header2 = "<H4 align=center><font face=Arial><span style=background-color:#FFF284>WARNING</span> at <span style=background-color:#FFF284>80%</span> in use. <span style=background-color:#FF9797>CRITICAL</span> at <span style=background-color:#FF9797>95%</span> in use.</font></H4>"  
 
$HTML += "<HTML><BODY><script src=E:\Script\DHCP_STAT\01_HTML_Report\sorttable.js></script><Table border=2 cellpadding=0 cellspacing=0 width=100% id=TSHead class=sortable> 
        <TR>  
            <TH><B>DHCP Server</B></TH> 
            <TH><B>Scope Name</B></TH> 
            <TH><B>Scope State</B></TH> 
            <TH><B>In Use</B></TH> 
            <TH><B>Free</B></TH> 
            <TH><B>% In Use</B></TH> 
            <TH><B>Reserved</B></TH> 
            <TH><B>Scope ID</B></TH> 
            <TH><B>Start of Range</B></TH> 
            <TH><B>End of Range</B></TH> 
            <TH><B>Subnet Mask</B></TH> 
            <TH><B>Gateway</B></TH> 
            <TH><B>Lease Duration</B></TH> 
            <TH><B>SuperScope Name</B></TH> 
         
        </TR> 
        "  
 
Foreach($Server in $ServerList) 
{ 
$ScopeList = Get-DhcpServerv4Scope -ComputerName $Server 
ForEach($Scope in $ScopeList.ScopeID)  
{ 
    Try{ 
    $ScopeInfo = Get-DhcpServerv4Scope -ComputerName $Server -ScopeId $Scope 
    $ScopeStats = Get-DhcpServerv4ScopeStatistics -ComputerName $Server -ScopeId $Scope | Select ScopeID,AddressesFree,AddressesInUse,PercentageInUse,ReservedAddress,SuperscopeName 
    $ScopeReserved = (Get-DhcpServerv4Reservation -ComputerName $server -ScopeId $scope).count 
    $ScopeRouterList = (Get-DhcpServerv4OptionValue -OptionId 3 -ScopeID $Scope -ComputerName $Server -ErrorAction:SilentlyContinue) 
    } 
    Catch{ 
    } 
 
# HTML Table values     
                $HTML += "<TR> 
                    <TD>$($Server)</TD> 
                    <TD>$($ScopeInfo.Name)</TD> 
                    <TD bgcolor=`"$(if($ScopeInfo.State -eq "Inactive"){"AAAAB2"})`">$($ScopeInfo.State)</TD> 
                    <TD>$($ScopeStats.AddressesInUse)</TD> 
                    <TD>$($ScopeStats.AddressesFree)</TD> 
                    <TD bgcolor=`"$(if($ScopeStats.PercentageInUse -gt $Alert){"FF9797"}elseif($ScopeStats.PercentageInUse -gt $Warn){"FFF284"}else{"A6CAA9"})`">$([System.Math]::Round($ScopeStats.PercentageInUse))</TD> 
                    <TD>$($ScopeReserved)</TD> 
                    <TD>$($ScopeInfo.ScopeID.IPAddressToString)</TD> 
                    <TD>$($ScopeInfo.StartRange)</TD> 
                    <TD>$($ScopeInfo.EndRange)</TD> 
                    <TD>$($ScopeInfo.SubnetMask)</TD> 
                    <TD>$($ScopeRouterList.Value)</TD> 
                    <TD>$($ScopeInfo.LeaseDuration)</TD> 
                    <TD>$($ScopeInfo.SuperscopeName)</TD> 
                    </TR>" 
}  
} 
 
$HTML += "<H2></Table></BODY></HTML>"  
$Header + $Header1 + $Header2 + $HTML | Out-File $OutputFile 
 
##################################################################  
# Now, We will zip the export file and send the same over email. # 
################################################################## 
 
#Below command will copy the file from 01_Report folder to Backup Location. 
 
Copy-Item -Path $OutputFile -Destination "E:\Script\DHCP_STAT\02_BackupLocation" 
 
## Below Code is to ZIP the folder ## 
$ZIPSource = "E:\Script\DHCP_STAT\02_BackupLocation" 
$ZIPDestination = "E:\Script\DHCP_STAT\03_ZIP_Report\DHCP_Stat_$(get-date -f yyyy-MM-dd-HHmm).zip" 
If(Test-path $ZIPDestination) {Remove-item $ZIPDestination} 
Add-Type -assembly "system.io.compression.filesystem" 
[io.compression.zipfile]::CreateFromDirectory($ZIPSource, $ZIPDestination)  
 
################################## 
# SMTP Settings for Email Report # 
################################## 
# "From" email address 
 
$EmailFrom = "Report@abc.com" 
 
# "To" email address 
$EmailTo1 = "TOrecipient1@abc.com" 
$EmailTo2 = "TOrecipient2@abc.com" 
$EmailTo3 = "TOrecipient3@abc.com" 
$EmailTo4 = "TOrecipient4@abc.com" 
$EmailTo5 = "TOrecipient5@abc.com" 
 
# "CC" email address 
$EmailCC = "CCrecipient1@abc.com" 
 
# Subject of the email 
$EmailSubject = "DHCP Scope Information :: Dated :: $(get-date -f yyyy-MM-dd)"  
 
# Body of the email  
$EmailBody = @" 
<html> 
<body style="font-family: Arial; font-size: 11pt;"> 
Dear All,<br></br> 
    Please find attached DHCP Scope information - Dated :: $(get-date -f yyyy-MM-dd).<br></br> 
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
 
# Below command with delete the extra file from Backup Location. 
 
Get-ChildItem “E:\Script\DHCP_STAT\02_BackupLocation” -recurse -include *.htm -force | remove-item