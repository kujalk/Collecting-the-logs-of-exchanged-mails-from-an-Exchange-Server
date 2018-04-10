# Collecting-the-logs-of-exchanged-mails-from-an-Exchange-Server
Simple PowerShell script to collect the logs of exchanged mails from an Exchange Server


In order to collect the logs, I have used the cmdlet "Get-MessageTrackingLog" and to get the mail box statistics I have used the cmdlet "Get-MailboxStatistics". There are 5 options provided for the user/admin (Exchange Admin) to analyze the logs according to the preference. They are,

[1] Mails sent from _______ to _______ 

[2] Mails sent by _______ 

[3] Mails received by _______ 

[4] All Logs

[5] Mail Box Statistics

[6] Exit

If the user input wrong value for the above selection, the script will terminate by throwing error message "Exiting........"

After selecting the correct preference, the script will ask for the user input related to the "sender" and "recipient" mail address depending on the selected preference.

The user only need to input the first part of the mail address only (no need to mention @...., as it is concatenated with the sender/recipient name).

$sender = Read-Host 

$sender = "$sender@example5.com"

Another feature added to this script is that the user/admin can capture the output on the screen itself or capture it on the text file, so that he can used to send/store the logs whenever needed.Text logs are named depending on the current time when it was created.

$a=Get-Date -Format G

$a=$a.replace("/","_")

$a=$a.replace(" ","-")

$a=$a.replace(":","-")

$log = Read-Host "`nPlease provide the location to store log file"

$location = "$log\Log_$a.txt"

To execute the script follow the below mentioned steps,
1. Store the script (CollectLog.ps1) in a local directory (C:\Exchange_Log) of the exchange server
2. Change the directory to the location where the script was stored (Can use Exchange Management Shell or Power shell)
-"cd C:\Exchange_Log"
3. Execute the script by running, "./CollectLog.ps1"
4. Follow the on screen instructions
