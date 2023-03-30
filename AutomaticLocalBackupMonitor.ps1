#Title:             Automatic Local Backup Monitor
#Author:            Andrew (AJ) Opfer
#Creation Date:     2022-12-04
#Phase:             Production
#Version:           1.0.1
#Description:       Small executable designed to create automated email alerting job on a machine.  User can define
#                   file(s) to be monitored, time to check for them to be expected to be updated, and provide emails
#                   for sending and receiving the alert(s).  This script is designed only to function with SMTP
#                   emails with SSL enabled.  Only supports a single sending email address.
#
#Updates:           3/22/2023 - 1.0.1 - Updated Scheduled Task creation to disable "Synchronize across time zones" setting

# ------------------ Variable Reference ------------------
# ========================================================
#  List of variables used in this script with description
# ========================================================
#      -- This does not include function variables --
#
# Intro:          Contains header text
# sysPath:        Installation folder
# ReadMe:         Contains text for readme file
# sysPrior:       Boolean for whether or not monitors exist
# sysScriptPath:  Script folder
# sysUser:		  User account running script
# sysCont:		  Main loop variable
# sysYesNo:       Basic Yes / No variable
# sysOrigUser:    User that originally saved email information
# sysEmail:       Sending email address
# sysServer:      Sending email server
# sysPort:        Sending email port
# emlPath:        Path for email-related data
# emlInfo:		  Saved email information file
# emlInfoContent: Contents of email info file
# usrOptions:     Multiple versions, for user options
# usrChoice:      Multiple versions, for user selection
# usrConfirm:     Multiple versions, for user confirmation
# cliName:        Reference name for client
# cliBackupName:  Referential name of backup
# cliFilePath:    Path to backup file being monitored
# cliFileName:    Full name of file
# cliFullFile:    Full path of monitored backup
# cliRecipient:   Recipient of email alerts

#Set variable for introduction
$Intro = @"
=============================================
  Automatic Local Backup Verification v1.0
---------------------------------------------
       Scripted By: Andrew (AJ) Opfer        
=============================================

"@

$sysPath = "C:\AutomaticLocalBackupMonitor\"


$ReadMe = $Intro + @"
The purpose of this utility is to monitor
locally stored backup files.  If you would
like to have everything created by this uti-
lity removed, you simply need to delete the
following file path:
	$sysPath

Additionally, open Task Scheduler and delete
all tasks within the subfolder titled:
	AutomaticLocalBackupMonitor

You can then safely delete that folder too.
	
If you are a technician troubleshooting an
issue surrounding this folder, please refer
to internal documentation regarding the Auto-
matic Local Backup Monitor.
"@

#Set email variables
$emlPath = $sysPath + "Email\"
$emlInfo = $emlPath + "info.txt"


#Set script path variable
$sysScriptPath = $sysPath + "Scripts\"

#Used for error checking Y/N prompts
$sysYesNo = "Y","N"

#Default main loop variable
$sysCont = "Y"

#Define variables required for pre-install check
$sysUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name


#Function to check if user input is valid
function chkInput($fctInput, $fctOptions){
	#Require user input value and list of valid options
	
	if( $fctInput -notin $fctOptions ){
		#If user input is invalid, prompts for correct input until valid
		do{
			Write-Host "The value you input was invalid.  Valid options include: " $fctOptions
			$fctInput = Read-Host
			
			#Empty line for formatting
			""
		}until( $fctInput -in $fctOptions )
		#Loop ends when user input is valid
		
	}
#Returns valid user input
return $fctInput
	
}

#Function to create file structure
function crtPaths{
	
	#Test to see if file structure already exists
	if( (test-path $sysPath) -eq $false ){
		#Create installation folder
		New-Item -Path $sysPath -ItemType Directory | Out-Null
		New-Item -Path $sysScriptPath -ItemType Directory | Out-Null
		New-Item -Path $emlPath -ItemType Directory | Out-Null
		
		#Create readme file for reference
		$ReadMe | Out-File ($sysPath + "README.txt")
		
		#Create scheduled task folder
		$scheduleObject = New-Object -ComObject schedule.service
		$scheduleObject.connect()
		$rootFolder = $scheduleObject.GetFolder("\")
		$rootFolder.CreateFolder("AutomaticLocalBackupMonitor") | Out-Null
	}
	
}

#Function to add saved email information
function setEmail{
	
	#Check if email information is already saved
	if( test-path ($emlInfo) ){
		#Grabs information 
		$emlInfoContent = Get-Content $emlInfo
		#Existing email already configured - Outputs stored email info to user
		Write-Host "The following email information is already recorded:"
		Write-Host "Saved By:            " $emlInfoContent.split(",")[0]
		Write-Host "Sending Email:       " $emlInfoContent.split(",")[1]
		Write-Host "SMTP Server:         " $emlInfoContent.split(",")[2]
		Write-Host "Port: 				 " $emlInfoContent.split(",")[3] "`n"
		Write-Host "Would you like to update this information? [Y/N]" 
		$emlChoice = Read-Host
		$emlChoice = (chkInput $emlChoice $sysYesNo)
		
		#Empty line for formatting
		""
		
	}else{
		#No existing email information stored
		$emlChoice = "Y"
	}
	if( $emlChoice -eq "Y" ){
		#Begin prompting user for email values
		Write-Host "SSL must be enabled on this email account."
		Write-Host "Please provide the email credentials."
		
		#Securely prompt for email creds
		$emlCredentials = Get-Credential
		$emlUsername = $emlCredentials.GetNetworkCredential().Username
		$emlPassword = $emlCredentials.GetNetworkCredential().Password
		
		Write-Host "Please input SMTP server address"
		$emlServer = Read-Host
		Write-Host "Please input sending port"
		$emlPort = Read-Host
		
		#Store user email information in this order, csv format:
		$sysUser + "," + `			#System user account that saved email info
		$emlUsername + "," + `			#Sending email
		$emlServer + "," + `		#Sending SMTP server
		$emlPort | `				#Sending email port
		Out-File ($emlInfo)
		#Encrypt and store password separately
		$emlPassword | Convertto-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Out-File ($emlPath + "login.txt")
		
		Write-Host "The information you provided has been saved."
		
	}else{"Email information has not been changed."}
}

#Function to create a monitor
function crtMonitor($MonitorPath,$MonitorFile,$BackupName,$MonitorTime,$Recipient,$Sender,$Port,$Server,$ClientName,$EmailPath,$ScriptPath){
	
	#First, create a Powershell Script
	Write-Host "Creating script..."
	
	@"
#This script checks to make sure the following file is created every day:
#	$MonitorPath$MonitorFile

`$Sender = "$Sender"
`$Password = Get-Content "${Emailpath}login.txt" | Convertto-SecureString
`$creds = New-Object System.Management.Automation.Pscredential -Argumentlist `$Sender,`$Password
`$TestFile = Get-ChildItem "$MonitorPath" -File | Sort-Object LastWriteTime -Descending | Where-Object {`$_.Name -like "$MonitorFile"} | Select-Object -First 1

if( (`$TestFile.LastWriteTime).DayOfYear -eq (Get-Date).DayOfYear ){
Send-MailMessage -To $Recipient -From $Sender -UseSsl -Port $Port -SmtpServer $Server -Subject "$ClientName - Local Backup Succeeded" -Body "$BackupName File was created.  No action is required." -Credential `$creds
} Else {
Send-MailMessage -To $Recipient -From $Sender -UseSsl -Port $Port -SmtpServer $Server -Subject "$ClientName - Local Backup Failed" -Body "$BackupName File was not created.  Please look into this." -Credential `$creds
}
"@ | Out-File ($ScriptPath + $BackupName + ".ps1")
	
	#Then, create a scheduled task
	Write-Host "Creating scheduled task..."
	$User = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
	$Credentials = Get-Credential -Credential $User
	$Password = $Credentials.GetNetworkCredential().Password
	
	$argument = "cd " + $ScriptPath + ";& '.\" + $BackupName + ".ps1'"
	$action = New-ScheduledTaskAction -Execute "%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe" -Argument $argument
	$trigger =  New-ScheduledTaskTrigger -Daily -At $MonitorTime
    $trigger.StartBoundary = (Get-Date -Date $MonitorTime -Format 'yyyy-MM-ddTHH:mm:ss')    #1.0.1 - "Synchronize across time zones"
	Register-ScheduledTask -Action $action -RunLevel Highest -Trigger $trigger -User $User -Password $Password -TaskName $BackupName -Description ("Checks to see if today's " + $BackupName + " file was created.") -TaskPath "AutomaticLocalBackupMonitor" | Out-Null

	$sysPrior = "TRUE"
	Write-Host "Monitor finished being created."
}

#Introduce software
$Intro
	
	#Main
	do{
		#If files created locally, assigns variables to them at start of loop
		if( test-path ($emlInfo) ){
			$emlInfoContent = Get-Content $emlInfo
			$sysOrigUser = $emlInfoContent.split(",")[0]
			$sysEmail = $emlInfoContent.split(",")[1]
			$sysServer = $emlInfoContent.split(",")[2]
			$sysPort = $emlInfoContent.split(",")[3]
		}
		
		#List options for user to proceed
		"What would you like like to do..."
		"  1. Manage email information"
		$usrOptions = 0,1
		#Checks to see if software has "installed" before
		if( test-path $sysPath ){
			#Sets variable for whether or not monitors exist
			if( (Get-ChildItem $sysScriptPath | Measure-Object).Count -eq 0){
				$sysPrior = "FALSE"
			}else{$sysPrior = "TRUE"}
			#Checks to see if email information is saved
			if ( test-path ($emlInfo) ){
				$usrOptions += 2
				"  2. Create a new monitor"
			}
			#Checks to see if any monitors already exist
			if( $sysPrior -eq "TRUE" ){
				$usrOptions += 3
				"  3. Delete an existing monitor"
			}
			$usrOptions += 4
			"  4. Remove everything created with this script"
		}
		"  0. Cancel"
		$usrChoice = Read-Host
		$usrChoice = (chkInput $usrChoice $usrOptions)
		
		#Empty line for formatting
		""
		
		#Separates actions via switch case
		switch($usrChoice) {
			1{	#Case 1: User manages email information
				"You've chosen to manage your email information."
				"---------------------------------------------"
				
				#Check if email information is already saved
				if( test-path ($emlInfo) ){
					#Existing email already configured
					$usrOptions2 = 1,2
					"There is already email information saved.  What would you like to do?"
					"  1. Check / Update settings"
					"  2. Delete existing settings"
					$usrChoice2 = Read-Host
					$usrChoice2 = (chkInput $usrChoice2 $usrOptions2)
					
					#Empty line for formatting
					""
					
					switch($usrChoice2){
						1{
							#Call function to check if file structure already exists
							crtPaths
							#Call function to update email settings
							setEmail
						}
						2{
							#Ask user to verify they'd like to delete email settings
							"You have chosen to delete your existing email information."
							"If you do not add new email information after doing so, all monitors will stop working."
							"Are you sure you'd like to proceed? [Y/N]"
							$usrConfirm2 = Read-Host
							$usrConfirm2 = (chkInput $usrConfirm2 $sysYesNo)
							
							#Empty line for formatting
							""
							
							if( $usrConfirm2 -eq "Y" ){
								#Delete existing email settings
								del $emlInfo
								del ($emlPath + "login.txt")
								"Email login information deleted."
							}else{
								#User has changed mind
								"You've opted not to delete the saved email information. No changes have been made."
							}
						}
					}
				#Email information not yet configured - send user to set email info
				}else{
					crtPaths
					setEmail
					}
					
				"============================================="
			}#End of case 1
			2{	#Case 2: User creates new monitor
				
				#Checks for file structure; creates if needed
				crtPaths
				
				if( $sysOrigUser -ne $sysUser ){
					"The user account you are signed into is not the account that was used to encrypt the sending email password."
					"This was the account that was used:"
					$sysOrigUser
					"`nIf you'd like to continue using that account, please sign into it and rerun this script."
					"If you'd prefer to use the account you are currently signed into, please select Option 1 at the main menu and reinput the email information."
				}else{
					
					"You've chosen to create a new monitor."
					"---------------------------------------------"
					"Please provide the following."
					""
					#Prompt user for necessary variables
					"Input client name, followed by Enter"
					$cliName = Read-Host
					"Input file referential name, followed by Enter"
					$cliBackupName = Read-Host
					"Input backup file path (in the format of 'C:\example\'), followed by Enter"
					$cliFilePath = Read-Host
					"Input the full name of the backup file (use asterisk for any non-static text, such as 'example_*.bak'), followed by Enter"
					$cliFileName = Read-Host
					
					#Merge file path and file name
					$cliFullFile = $cliFilePath + $cliFileName
					
					#Empty line for formatting
					""

					#Check if user input was valid
					if( test-path $cliFullFile ){
						#File exists
						"You've provided the following file to be monitored:"
						$cliFullFile
						"Are you sure this is the file you'd like to monitor? [Y/N]"
						$usrConfirm1 = Read-Host
						$usrConfirm1 = (chkInput $usrConfirm1 $sysYesNo)
						
						#Empty line for formatting
						""
						
					}else{
						#File does not exist
						"The following file does not exist:"
						$cliFullFile
						"Would you like to proceed with creating a monitor for this file regardless? [Y/N]"
						$usrConfirm1 = Read-Host
						$usrConfirm1 = (chkInput $usrConfirm1 $sysYesNo)
						
						#Empty line for formatting
						""
					
					}
					
					if( $usrConfirm1 -eq "Y" ){ #User opted to create monitor
						"Input the local system time this file should be monitored (in the format of '12pm'), followed by Enter"
						$cliTime = Read-Host
						
						"Please input the email address you'd like to send alerts to"
						$cliRecipient = Read-Host
						
						#Empty line for formatting
						""
						
						#Call function to create monitor
						crtMonitor $cliFilePath $cliFileName $cliBackupName $cliTime $cliRecipient $sysEmail $sysPort $sysServer $cliName $emlPath $sysScriptPath
					}else{ #User opted to cancel
						"No monitor has been created."
					}
				}
				"============================================="
			}#End of case 2
			3{	#Case 3: User deletes existing monitor
				"You've chosen to delete an existing monitor."
				"---------------------------------------------"
				
				#Prepares array of existing monitors
				[array]$MonitorList = "Cancel"
				[array]$MonitorList = $MonitorList + (Get-ChildItem $sysScriptPath | Foreach-Object {$_.BaseName})
				
				#List all existing monitors
				"Which monitor would you like to delete?"
				for($a = 0;$a -lt $MonitorList.Count;$a++){
					Write-Host "  " $a ". " $MonitorList[$a]
					[array]$usrOptions3 += $a
				}
				$usrChoice3 = Read-Host
				$usrChoice3 = (chkInput $usrChoice3 $usrOptions3)
				
				#Empty line for formatting
				""
				
				if( $usrChoice3 -eq 0 ){
					#User has changed mind
					"You've opted not to delete any monitors. No changes have been made."
				}else{
					"`nDeleting monitor " + $MonitorList[$usrChoice3]
					Unregister-ScheduledTask -TaskName $MonitorList[$usrChoice3] -Confirm:$false
					$delScript = $sysScriptPath + $MonitorList[$usrChoice3] + ".ps1"
					del $delScript
					
					#If no more monitors exist, change Installation.txt value
					if( (Get-ChildItem $sysScriptPath | Measure-Object).Count -eq 0){
						"No more monitors created by this script exist."
					}else{"Monitor finished being deleted."}
				}
				
				"============================================="
			}#End of case 3
			4{	#Case 4: User 'uninstalls' monitoring software
				"You've chosen to remove all existing monitors and files created by this script."
				"---------------------------------------------"
				"Proceeding with this option will delete all monitors and saved email information."
				""
				"Are you sure you'd like to fully remove all monitors and information? [Y/N]"
				$usrConfirm4 = Read-Host
				$usrConfirm4 = (chkInput $usrConfirm4 $sysYesNo)
				
				#Empty line for formatting
				""
				
				if( $usrConfirm4 -eq "Y"){
					#Prepares array of existing monitors
					[array]$MonitorList = Get-ChildItem $sysScriptPath | Foreach-Object {$_.BaseName}
					#Deletes each monitor individually - unable to delete container until each monitor is gone
					for($a = 0;$a -lt $MonitorList.Count;$a++){
						"Deleting monitor " + $MonitorList[$a] + "..."
						Unregister-ScheduledTask -TaskName $MonitorList[$a] -Confirm:$false
					}
					
					#Deletes Scheduled Task container
					"Deleting Task Scheduler container..."
					$scheduleObject = New-Object -ComObject schedule.service
					$scheduleObject.connect()
					$rootFolder = $scheduleObject.GetFolder("\")
					$rootFolder.DeleteFolder("AutomaticLocalBackupMonitor",$null)

					#Deletes file paths
					"Deleting files and folders..."
					Remove-Item -path $sysPath -recurse

                    "Everything that has ever been created with this script has been fully deleted."
				}else{
					#User has changed mind
					"You've opted not to remove the software. No changes have been made."
				}

		        "============================================="

			}#End of case 4
			0{  #Case 0: User is canceling operations
				"You've chosen to cancel. No changes have been made."
				"---------------------------------------------"
			}#End of case 0
		}
		
		"Would you like to do anything else? [Y/N]"
		$sysCont = Read-Host
		$sysCont = (chkInput $sysCont $sysYesNo)
		
        #Empty line for formatting
		""

	}until($sysCont -eq "N")

"============================================="
"End of application.  Press Enter to close..."
Read-Host
