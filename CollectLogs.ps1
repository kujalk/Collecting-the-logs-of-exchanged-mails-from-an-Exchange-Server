####################################################################################
#This script is used to collect the logs in an Exchange Server                     #
#Developer - Kugathasan Janarthanan                                                #
#Date - 04/04/2018                                                                 #
###################################################################################

#Clearing the screen
cls


Write-Host "Please select one of the any option given below `n`n" -foreground yellow

Write-host "[1] Mails sent from _______ to _______ " -foreground green
Write-host "[2] Mails sent by _______ " -foreground green
Write-host "[3] Mails received by _______ " -foreground green
Write-host "[4] All Logs" -foreground green
Write-Host "[5] Mail Box Statistics" -foreground green
Write-host "[6] Exit" -foreground green

Write-Host "`nPlease input relavant number " -foreground DarkBlue 
$reply = Read-Host  


#check whether all details are correct
If ( $reply -eq "1")
{

Write-Host "`nPlease provide the sender name " -foreground Magenta 
$sender = Read-Host 

Write-Host "`nPlease provide the recepient name " -foreground Magenta
$receiver = Read-Host  

$sender = "$sender@example5.com"
$receiver = "$receiver@example5.com"

$answer=Get-MessageTrackingLog -EventId "Send" -Source "SMTP" -Sender $sender -Recipients $receiver | select Timestamp,MessageSubject
}

elseif ( $reply -eq "2")
{
Write-Host "`nPlease provide the sender name " -foreground Magenta 
$sender = Read-Host 

$sender = "$sender@example5.com"

$answer=Get-MessageTrackingLog -EventId "Send" -Source "SMTP" -Sender $sender | select Timestamp,MessageSubject
}

elseif ( $reply -eq "3")
{
Write-Host "`nPlease provide the recepient name " -foreground Magenta
$receiver = Read-Host  

$receiver = "$receiver@example5.com"

$answer=Get-MessageTrackingLog -EventId "Send" -Source "SMTP" -Recipients $receiver | select Timestamp,MessageSubject
}

elseif ( $reply -eq "4")
{
$answer=Get-MessageTrackingLog -EventId "Send" -Source "SMTP"  | select Timestamp,Sender,Recipients,MessageSubject
}

elseif ( $reply -eq "5")
{
Write-Host "`nPlease provide the recepient name " -foreground Magenta
$who = Read-Host

$answer=Get-MailboxStatistics $who | select-object DisplayName, ItemCount,TotalItemSize, TotalDeletedItemSize, LastLogonTime, LastLogoffTime
}

elseif ( $reply -eq "6")
{
Write-Host "`nExiting........`n`n" -foreground DarkRed   
exit
}

else
{
Write-Host "`nPlease select correct number`n`n" -foreground DarkRed   
exit
}

#There are some output
Write-Host "`nHow to display the results,`n		[1]Text file`n		[2]On Screen " -foreground DarkGray
$display = Read-Host  


while (($display -ne "1") -and ($display -ne "2"))
{
cls
Write-Host "`nPlease provide correct value" -foreground DarkRed 
Write-Host "`nHow to display the results,`n		[1]Text file`n		[2]On Screen " -foreground DarkGray
$display = Read-Host   
}

#Output 1
if ( $display -eq "1")
{
$a=Get-Date -Format G
$a=$a.replace("/","_")
$a=$a.replace(" ","-")
$a=$a.replace(":","-")


$log = Read-Host "`nPlease provide the location to store log file"
$location = "$log\Log_$a.txt"

echo $answer >> $location

Write-Host "`nLog saved" -foreground Magenta  
}

#Output 2
elseif ( $display -eq "2")
{
Write-Host "`n`n"
echo $answer
}


