#add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010 -erroraction silentlyContinue
#Set-ADServerSettings -ViewEntireForest $True
###################################################################################################################################################################################### 
# Report used to check for specific Event ID's on servers and exporting the data to a HTML Report  
#Version 0.1
# Author: Tiens van Zyl
# Date 10 December 2015
# Updated by Tiens van Zyl 7 December 2015
# Updates: 
# This script outputs Exchange Quota Defaults True for our weekly report 
# 1. The script exports a CSV file with the current date.
# 2. The CSV file export displays the DisplayName, IssueWarningQuota, ProhibitSendQuota and ProhibitSendReceiveQuota 
# 3. The CSV file contains raw data that needs to be formatted in Excel
# 4. Enter your mailbox server/s name in place of "ServerName". Use a wilcard if you have more than one server you need to query i.e. mailbox0* if your servers names are mailbox01, mailbox02 etc.
# 5. Set the path to where you'd like to export the txt file. 
# 6. The FileName will be appended with the date that the script is run. i.e. WeeklyExchangeDefaultScript 2015-05-01.csv
######################################################################################################################################################################################
#Styling the HTML Report - Credit to https://technet.microsoft.com/en-us/library/ff730936.aspx
$a = "<style>"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:thistle}"
$a = $a + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:palegoldenrod}"
$a = $a + "</style>"
#Declaring the variables and then adding them to the Array name $MachineNames
#Angola
$AngolaHC1 = "ServerName"
$AngolaHC2 = "ServerName"
$AngolaMB = "ServerName"
$AngolaMB = "ServerName"
#Botswana
$BotswanaHC1 = "ServerName"
$BotswanaHC2 = "ServerName"
$BotswanaMB = "ServerName"
$BotswanaMB = "ServerName"
#Ghana
$GhanaHC1 = "ServerName"
$GhanaHC2 = "ServerName"
$GhanaMB1 = "ServerName"
$GhanaMB2 = "ServerName"
#Mozambique
$MozambiqueHC1 = "ServerName"
$MozambiqueHC2 = "ServerName"
$MozambiqueMB1 = "ServerName"
$MozambiqueMB2 = "ServerName"
#Nigeria
$NigeriaHC1 = "ServerName"
$NigeriaHC2 = "ServerName"
$NigeriaMB1 = "ServerName" 
$NigeriaMB2 = "ServerName"
#SouthAfrica
$SAHT1 = "ServerName"
$SAHT2 = "ServerName"
$SAEWS1 = "ServerName"
$SAEWS2 = "ServerName"
$SAEWS1 = "ServerName"
$SAEWS2 = "ServerName"
#Namibia
$NamibiaMB03 = "ServerName"

#ArrayName
$MachineNames = $AddVariablesHereSeperatedByComma
Get-Eventlog -Logname Application -After (Get-Date).Adddays(-4)-computer $MachineNames | Where-Object {$_.EventId -eq 4002 -or $_.EventId -eq 4002} | Format-Table MachineName, TimeWritten, Source, EventID, Message -auto
| ConvertTo-HTML -head $a | Out-File C:\Temp\Eventlogs.htm