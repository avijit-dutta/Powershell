##########################################################################
# Name: Avijit Dutta                                                      
# ScriptName :   Get-DHCPScopeOption-DNS
# Purpose : This script will fetch the DNS configured in the Scope Option or Server Option.# 
# Creation Date : 01-01-2016                    # 
# Note: Taken a reference from Microsoft Script Center and modified as per our enviroment  # 
############################################################################################ 
 
############################################################################################ 
# Prerequsites                                                                             # 
# 1) Create below folder hierarchy                                                         # 
#          DHCP_DNS_Info                                                                   # 
#                01_Report                                                                 # 
#                02_BackupLocation                                                         #         
#                03_03_ZIP_Report                                                          # 
# 2) SMTP Server                                                                           # 
# 3) Recipient email address.                                                              # 
############################################################################################ 
 
Import-Module DHCPServer 
# Get all Authorized DCs from AD configuration with below command. But, I have mentioned them manually. 
# $DHCPs = Get-DhcpServerInDC 
 
# Manually, mentioned the DHCP Server list  
$DHCPs = 'Server1','Server2','Server3','Server4' 
 
# Path and Filename of the Exported data 
$Filename = "E:\Script\DHCP_DNS_Info\01_Report\DHCPScopes_DNS_$(get-date -Uformat "%Y%m%d-%H%M%S").csv" 
 
$Report = @() 
$k = $null 
Write-Host -foregroundcolor Green "`n`n`n`n`n`n`n`n`n" 
foreach ($dhcp in $DHCPs) { 
    $k++ 
    Write-Progress -activity "Getting DHCP scopes:" -status "Percent Done: " ` 
    -PercentComplete (($k / $DHCPs.Count)  * 100) -CurrentOperation "Now processing $($dhcp)" 
    $scopes = $null 
    $scopes = (Get-DhcpServerv4Scope -ComputerName $dhcp -ErrorAction:SilentlyContinue) 
    If ($scopes -ne $null) { 
        #getting global DNS settings, in case scopes are configured to inherit these settings 
        $GlobalDNSList = $null 
        $GlobalDNSList = (Get-DhcpServerv4OptionValue -OptionId 6 -ComputerName $dhcp -ErrorAction:SilentlyContinue).Value 
        $scopes | % { 
            $row = "" | select Hostname,ScopeID,SubnetMask,Name,State,StartRange,EndRange,LeaseDuration,Description,DNS1,DNS2,DNS3,GDNS1,GDNS2,GDNS3 
            $row.Hostname = $dhcp 
            $row.ScopeID = $_.ScopeID 
            $row.SubnetMask = $_.SubnetMask 
            $row.Name = $_.Name 
            $row.State = $_.State 
            $row.StartRange = $_.StartRange 
            $row.EndRange = $_.EndRange 
            $row.LeaseDuration = $_.LeaseDuration 
            $row.Description = $_.Description 
            $ScopeDNSList = $null 
            $ScopeDNSList = (Get-DhcpServerv4OptionValue -OptionId 6 -ScopeID $_.ScopeId -ComputerName $dhcp -ErrorAction:SilentlyContinue).Value 
            #write-host "Q: Use global scopes?: A: $(($ScopeDNSList -eq $null) -and ($GlobalDNSList -ne $null))" 
            If (($ScopeDNSList -eq $null) -and ($GlobalDNSList -ne $null)) { 
                $row.GDNS1 = $GlobalDNSList[0] 
                $row.GDNS2 = $GlobalDNSList[1] 
                $row.GDNS3 = $GlobalDNSList[2] 
                $row.DNS1 = $GlobalDNSList[0] 
                $row.DNS2 = $GlobalDNSList[1] 
                $row.DNS3 = $GlobalDNSList[2] 
                } 
            Else { 
                $row.DNS1 = $ScopeDNSList[0] 
                $row.DNS2 = $ScopeDNSList[1] 
                $row.DNS3 = $ScopeDNSList[2] 
                } 
            $Report += $row 
            } 
        } 
    Else { 
        write-host -foregroundcolor Yellow """$($dhcp)"" is either running Windows 2003, or is somehow not responding to querries. Adding to report as blank" 
        $row = "" | select Hostname,ScopeID,SubnetMask,Name,State,StartRange,EndRange,LeaseDuration,Description,DNS1,DNS2,DNS3,GDNS1,GDNS2,GDNS3 
        $row.Hostname = $dhcp 
        $Report += $row 
        } 
    write-host -foregroundcolor Green "Done Processing ""$($dhcp)""" 
    } 
  
$Report  | Export-csv -NoTypeInformation -UseCulture $filename 
 
##################################################################  
# Now, We will zip the export file and send the same over email. # 
################################################################## 
 
#Below command will copy the file from 01_Report folder to Backup Location. 
 
Copy-Item -Path $OutputFile -Destination "E:\Script\DHCP_DNS_Info\02_BackupLocation" 
 
## Below Code is to ZIP the folder ## 
$ZIPSource = "E:\Script\DHCP_DNS_Info\02_BackupLocation" 
$ZIPDestination = "E:\Script\DHCP_DNS_Info\03_ZIP_Report\DHCP_Stat_$(get-date -f yyyy-MM-dd-HHmm).zip" 
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
#$EmailTo3 = "TOrecipient3@abc.com" 
#$EmailTo4 = "TOrecipient4@abc.com" 
#$EmailTo5 = "TOrecipient5@abc.com" 
 
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
Server Admin</br> 
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
#$Message.To.Add($EmailTo3) 
#$Message.To.Add($EmailTo4) 
#$Message.To.Add($EmailTo5) 
 
#$Message.CC.Add($EmailCC) 
$Message.Body = $EmailBody 
$message.Subject = $EmailSubject 
$attach = new-object Net.Mail.Attachment($attachment) 
$message.Attachments.Add($attach) 
$message.IsBodyHTML = $true 
$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, $SMTPPort) 
$SMTPClient.Send($Message) 
$Message.dispose() 
 
# Below command with delete the extra file from Backup Location. 
 
Get-ChildItem “E:\Script\DHCP_DNS_Info\02_BackupLocation” -recurse -include *.htm -force | remove-item