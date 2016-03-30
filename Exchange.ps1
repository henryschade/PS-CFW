###########################################
# Updated Date:	29 March 2016
# Purpose:		Exchange routines.
# Requirements:	.\EWS-Files.txt  ($strEWSFiles)
#				CreateMailBox() needs Jobs.ps1 if you want to run it in a background process, 
#					and the following need to be defined/setup: $global:objJobs, and $global:objExchPool.
#				CreateMailBox() uses a routine UpdateResults() that is in the calling project (AScII and Exchange-GUI).
##########################################
	#https://msdn.microsoft.com/en-us/library/office/jj900166(v=exchg.150).aspx
	#https://msdn.microsoft.com/en-us/library/dn567668(v=exchg.150).aspx
	#https://msdn.microsoft.com/en-us/library/dd633696(v=exchg.80).aspx

	#EWS Managed API v1.2 -->  http://www.microsoft.com/en-us/download/details.aspx?id=28952
	#EWS v2.0  -->  http://www.microsoft.com/en-us/download/details.aspx?id=35371
	#EWS v2.2  -->  http://go.microsoft.com/fwlink/?LinkId=255472

	##To Include this Script/File.
	#$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;
	#. ($ScriptDir + "\Exchange.ps1")

	$strEWSFiles = "EWS-Files.txt";
	$strThisFileDir = "";
	if (($MyInvocation.MyCommand.Path -ne "") -and ($MyInvocation.MyCommand.Path -ne $null)){
		$strThisFileDir = Split-Path $MyInvocation.MyCommand.Path;
	}

	#URL's from initial Data searching:
	<#
		#Some good info that might be helpful:
		#http://www.msexchange.org/articles-tutorials/exchange-server-2010/monitoring-operations/monitoring-exchange-2007-2010-powershell-part1.html

		#EWS (Streaming) notifications / subscriptions
		#http://gsexdev.blogspot.com/2011/09/using-ews-streaming-notifications-in.html
			#... default is 20 Subs per account ...
		#https://msdn.microsoft.com/en-us/library/office/dn458791(v=exchg.150).aspx
		#http://stackoverflow.com/questions/21636201/how-to-unsubscribe-from-ews-push-notification-using-managed-api

		#http://blogs.technet.com/b/heyscriptingguy/archive/2011/12/02/learn-to-use-the-exchange-web-services-with-powershell.aspx
		#http://stackoverflow.com/questions/22213607/script-to-download-emails-from-exchange-as-html-cant-use-outlook
		#http://gsexdev.blogspot.com/2014/08/getting-folder-sizes-and-other-stats.html
		#C# and PS sample  -->  http://stackoverflow.com/questions/26539285/ews-wont-work-in-powershell
		#https://github.com/krispharper/Powershell-Scripts/blob/master/Mark-ExchangeItemsAsRead.ps1
		#http://stackoverflow.com/questions/4454165/how-to-check-an-exchange-mailbox-via-powershell
		#http://stackoverflow.com/questions/19439222/download-attachments-with-multiple-subjects-from-exchange
		#http://stackoverflow.com/questions/29776157/retrieving-extended-properties-from-contacts-using-ews-and-powershell
		#http://poshcode.org/2520

		#Exchange Impersonation
		#http://www.thesoftwaregorilla.com/2010/06/exchange-web-services-example-part-3-exchange-impersonation/
	#>

	function SampleSendEmail{
		#Send email with EWS
		$strTo = "henry.schade@nmci-isf.com; andrew.k.freeman@nmci-isf.com; henry.schade@hpe.com";
		#$strTo = "henry.schade@nmci-isf.com";
		$strFrom = "henry.schade@nmci-isf.com";
		$strSub = "Test eMail from PS EWS";
		$strBody = "Test eMail from PS EWS";
		$strFile = "C:\Users\henry.schade\Desktop\ps command issue.txt";

		#$objRet = SendEmailEWS $strTo $strFrom $strSub $strBody;
		$objRet = SendEmailEWS $strTo $strFrom $strSub $strBody -EmailAttach $strFile;
		if ($objRet.Results -ne $True){
			Write-Host $objRet.Message;
		}
	}

	function TestMe{

		#permissions to a mailbox:
		#SetupConn
		#& "C:\Projects\PS-CFW\Exchange.ps1" SetUp
		#$MailboxName = "henry.schade@nmci-isf.com";
		#$strUserToAdd = "henry.schade.adm";
		#Check permissions:
		#Get-MailboxPermission -Identity $MailboxName;
		#Check for Auto-Mapping:
			#If there is anything in the msExchDelegateListLink or msExchDelegateListBL attributes then the account has Auto-Mapping.
		#Get-ADUser -Identity first.last -Properties msExchDelegateListLink, msDelegateListBL -Server domain;
		#Get-ADUser -Identity first.last -Properties * -Server domain;

		#Grant Mailbox permissions (Member)
		#Add-MailboxPermission -Identity $MailboxName -User $strUserToAdd -AccessRights FullAccess,DeleteItem -InheritanceType All -Automapping $False;
		#Grant Mailbox permissions (Admin)
		#Add-MailboxPermission -Identity $MailboxName -User $strUserToAdd -AccessRights FullAccess,DeleteItem,ReadPermission,ChangePermission,ChangeOwner -InheritanceType All -Automapping $False;

		#Remove Mailbox permissions (Any and All).  (Works even if have less perms)
		#Remove-MailboxPermission -Identity $MailboxName -User $strUserToAdd -AccessRights FullAccess,SendAs,ExternalAccount,DeleteItem,ReadPermission,ChangePermission,ChangeOwner -InheritanceType All -Confirm:$False;



		. C:\Projects\PS-Scripts\Exchange.ps1;
		$MailboxName = "hd_voicemail@nmci-isf.com";
		$MailboxName = "servicedesk_navy@nmci-isf.com";
		$MailboxName = "SRM_UA_APP_ERRORS@nmci-isf.com";
		$MailboxName = "SRM_CustomerSat@nmci-isf.com";
		$MailboxName = "henry.schade@nmci-isf.com";
		$intNumGet = 10;
		$intNumGet = 0;
		$strFolder = "3 Month Complete Inbox Reference Copy";
		$strFolder = "\Inbox\Camping";
		$strFolder = "Inbox\Camping";
		$strFolder = "Inbox\CLaPR";
		$strFolder = "\Inbox\Programming\CLIN23";
		$strFolder = "Inbox";
		$objReturn = $null;

		#$objReturn = EWSVerifyInstall "1.2";
		#EWSGetFolder() calls EWSVerifyInstall().
		#$objReturn = EWSGetFolder $MailboxName $strFolder;
		#$objFolder = $objReturn.Returns;
			#$objSubFolders = $objFolder.FindFolders($objFolder.ChildFolderCount)
			#foreach ($objSub in $objSubFolders){
			#	Write-Host $objSub.DisplayName $objSub.TotalCount;
			#}
		#$objReturn = EWSGetEmails -objEWSFolder $objFolder -NumToRet $intNumGet;
		##$objReturn = EWSGetEmails -objEWSFolder $objFolder -NumToRet $intNumGet -bolDoProperties $False;
			#$objItems = $objReturn.Returns;

		#or:

		#EWSGetEmails() calls EWSGetFolder(), IF -objEWSFolder switch is NOT used.
		$objReturn = EWSGetEmails -MailboxName $MailboxName -WhatFolder $strFolder -NumToRet $intNumGet;
		#$objReturn = EWSGetEmails -MailboxName $MailboxName -WhatFolder $strFolder -NumToRet $intNumGet -bolDoProperties $False;

		#$objItems = $objReturn.Returns;
		#Write-Host @($objItems).Count;

		#Loop through each email
		Write-Host $objReturn.Returns.TotalCount;
		for ($intX = 0; $intX -lt $objReturn.Returns.TotalCount; $intX++){
			#replace any whitespace with a single space then get the 1st 90 chars
			$strBody = $objReturn.Returns.Items[$intX].Body.Text -replace '\s+', ' ';
			if (($strBody -ne "") -and ($strBody -ne $null)){
				$strBody = $strBody.Trim().SubString(0, 90);
			}else{
				$strBody = "[" + $objReturn.Returns.Items[$intX].ItemClass + "] ";
				if ($objReturn.Returns.Items[$intX].ItemClass -eq "IPM.Note"){
					$strBody = "[blank]";
				}
			}
			$strBody = "$strBody...";

			# output the results - first of all the From, Subject, References and Message ID
			Write-Host ("+" + ("-" * 113) + "+");
			Write-Host ("|From             : " + $objReturn.Returns.Items[$intX].From.Name + (" " * (94 - $objReturn.Returns.Items[$intX].From.Name.Length)) + "|");
			#Write-Host ("|Subject          : " + [String]$objReturn.Returns.Items[$intX].Subject + (" " * (94 - $objReturn.Returns.Items[$intX].Subject.Length)) + "|");
			$strSubject = $objReturn.Returns.Items[$intX].Subject.Trim();
			if ($strSubject.Length -gt 94){
				$strSubject = $strSubject.SubString(0, 90);
				$strSubject = $strSubject + "...";
			}
			Write-Host ("|Subject          : " + $strSubject + (" " * (94 - $strSubject.Length)) + "|");
			Write-Host ("|Body             : " + $strBody + (" " * (94 - $strBody.Length)) + "|");
			#Write-Host ("|HasAttachments   : " + $objReturn.Returns.Items[$intX].HasAttachments + (" " * (89 - $objReturn.Returns.Items[$intX].HasAttachments.Length))) -NoNewLine;
			Write-Host ("|HasAttachments   : " + $objReturn.Returns.Items[$intX].HasAttachments + (" " * (94 - $objReturn.Returns.Items[$intX].HasAttachments.ToString().Length)) + "|");
			if ($objReturn.Returns.Items[$intX].HasAttachments){
				foreach ($objAttach in $objReturn.Returns.Items[$intX].Attachments){
					Write-Host ("|  Attachment Name: " + $objAttach.Name + (" " * (94 - $objAttach.Name.Length)) + "|");
					Write-Host ("|  Attachment Size: " + $objAttach.Size + (" " * (94 - $objAttach.Size.ToString().Length)) + "|");
				}
			}
			Write-Host ("+" + ("-" * 113) + "+");
		}



		#Get all ITEMS from Folders and Sub-folders (of PublicFolders)
		#http://stackoverflow.com/questions/13877629/how-to-get-all-items-from-folders-and-sub-folders-of-publicfolders-using-ews-man



		#Sample delete blocks:
		<#
			#Get the emails in a folder, then loop through the records/objects and delete each one.
			Get-Date;
			$MailboxName = "servicedesk_navy@nmci-isf.com";
			$strFolder = "\Inbox\ITSS\Quote";
			$intNumGet = 0;		#Even if specify no limit (0 = no limit) EWS only returns 1000 items.
			#do{
				$objReturn = $null;
				$objReturn = EWSGetEmails -MailboxName $MailboxName -WhatFolder $strFolder -NumToRet $intNumGet -bolDoProperties $False;
				Get-Date;
				#$objReturn | FL;
				Get-Date;
				for ($intX = 0; $intX -lt $objReturn.Results; $intX++){
					if ($objReturn.Returns.Items[$intX].SomeDateField -ge "todaysDate"){
						$objReturn.Returns.Items[$intX].Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete);
					}else{
						break;
					}
				}
			#} while($objReturn.Returns.TotalCount -ge 1000)
				#} while($objReturn.Returns.MoreAvailable -eq $True)
			Get-Date;



			#Empty a directory/folder.  About 1 min to do 10,000 emails.  Then delete the folder.
			Get-Date;
			$MailboxName = "servicedesk_navy@nmci-isf.com";
			$strFolder = "\Inbox\ITSS\Quote";
			$objFolder = $null;
			$objFolder = EWSGetFolder $MailboxName $strFolder;
			Write-Host "Emptying " $objFolder.Returns.TotalCount " emails from " $strFolder " .";
			$objFolder.Returns.Empty([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete, $False);
			#Now delete the empty folder:
			#$objFolder.Returns.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete);		#Perminate delete
			Get-Date;

		#>



		#Manipulate mail:
		#http://gsexdev.blogspot.com/2012/02/ews-managed-api-and-powershell-how-to_22.html
			#$objReturn.Returns.Items[0].Copy($TargetFolderObject.Returns.ID)
			#$objReturn.Returns.Items[0].Move($TargetFolderObject.Returns.ID)

			#$objReturn.Returns.Items[1].IsRead = $False;
			#$objReturn.Returns.Items[1].Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite);

			#$objReturn.Returns.Items[1].Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete);		#Sends to Dumpster
			#$objReturn.Returns.Items[1].Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete);		#Perminate delete




		#To open a specific email:
		#. C:\SRM_Apps_N_Tools\PS-Scripts\Exchange.ps1;
		#$strID = "AAMkADE4MDAyMmZjLTM5OTAtNGRmOC1hYTAxLTEzMzQ3MmEwYjRjOQBGAAAAAAAvrFUL8mFTTY8vHRCWCfoYBwAhDPj33Z4iQZV9S7BKG8SJAAAAYmR6AACuTYmN9cQLRbE2n73/sFJrAAAy2iVdAAA=";
		#$MailboxName = "henry.schade@nmci-isf.com";
		#$objReturn = $null;
		#$objService = (EWSGetFolder $MailboxName "").Returns;
		#$objItemId = New-Object Microsoft.Exchange.WebServices.Data.ItemId($strID);		#The Item MUST be in the Folder thats been connected to.
		##$objMessage = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($objService, $objItemId, $psPropertySet);
		#$objMessage = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($objService, $objItemId);
		#$objMessage;



		#Subscription
		<#
			#http://stackoverflow.com/questions/5911904/where-is-my-streaming-subscription-going

			#7/16/2015 15:47:54 Subscription Error
			#event:  System.Management.Automation.PSEventArgs
			#.MessageData:  Microsoft.Exchange.WebServices.Data.StreamingSubscriptionConnection
			#.Exception.Message:
			#7/16/2015 15:47:54 Disconnecting...


			#https://msdn.microsoft.com/en-us/library/office/dn458791(v=exchg.150).aspx
			#http://stackoverflow.com/questions/21636201/how-to-unsubscribe-from-ews-push-notification-using-managed-api

			. C:\Projects\PS-CFW\Exchange.ps1;
			$MailboxName = "henry.schade@nmci-isf.com";


			$objActionScript = [scriptblock]::create($function:EWSOnEventDisplay);
			$objReturn = EWSCreateSubscriptionStream $MailboxName $objActionScript;
			$objReturn

			#Dealing w/ Events:
			#http://blogs.technet.com/b/heyscriptingguy/archive/2011/06/17/manage-event-subscriptions-with-powershell.aspx
			#Get-EventSubscriber;
			##Unsubscribe-Event -SubscriptionId 1;   #PS3+
			#Unregister-Event -SubscriptionId 1;



			#Original test code:
			$objReturn = $null;

			$arrFoldersToWatch = New-Object Microsoft.Exchange.WebServices.Data.FolderId[] 1;
			$Inboxid = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $MailboxName);
			#$Inboxid = (EWSGetFolder $MailboxName "Inbox").Returns;		#Does NOT work.  Cannot convert the "Microsoft.Exchange.WebServices.Data.Folder" value of type "Microsoft.Exchange.WebServices.Data.Folder" to type "Microsoft.Exchange.WebServices.Data.FolderId".
			$arrFoldersToWatch[0] = $Inboxid;

			#$objService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList $strExchVer;
			$objReturn = EWSGetFolder $MailboxName "";
			$objService = $objReturn.Returns;

			#http://gsexdev.blogspot.com/2012/05/ews-managed-api-and-powershell-how-to_28.html
			#Create the "NewMail" subscription object.
			$objSubscription = $objService.SubscribeToStreamingNotifications($arrFoldersToWatch, [Microsoft.Exchange.WebServices.Data.EventType]::NewMail);
			$objConnection = New-Object Microsoft.Exchange.WebServices.Data.StreamingSubscriptionConnection($objService, 30);
				#exchange.SubscribeToPullNotifications(new FolderId[] { _calendarFolderId }, Settings.SubscriptionTimeout, null, EventType.Created, EventType.Deleted, EventType.Modified);
			$objConnection.AddSubscription($objSubscription);

			#http://gsexdev.blogspot.com/2012/05/ews-managed-api-and-powershell-how-to_28.html
				#In Powershell the Register-ObjectEvent cmdlet allows you to subscribe to the events that are generated by the Microsoft .NET Framework.
			#Register the OnNotificationEvent() event.
			Register-ObjectEvent -inputObject $objConnection -eventName "OnNotificationEvent" -Action $function:EWSOnEventDisplay -MessageData $objService;
			#Should Register the OnSubscriptionError () event too (testing).
				#http://stackoverflow.com/questions/5911904/where-is-my-streaming-subscription-going
			Register-ObjectEvent -inputObject $objConnection -eventName "OnSubscriptionError" -Action {Write-Host (Get-Date)  "Subscription Error "; $event.Exception.Message;} -MessageData $objConnection;
			#Register the OnDisconnect() event.
			Register-ObjectEvent -inputObject $objConnection -eventName "OnDisconnect" -Action {Write-Host "Disconnecting..."; $Error.Clear(); $event.MessageData.Open(); if($Error){Write-Host "Error ReConnecting.`r`n $Error `r`n`r`n"}else{Write-Host "ReConnected. `r`n`r`n"};} -MessageData $objConnection;
			#$objResults = Register-ObjectEvent -inputObject $objConnection -eventName "OnDisconnect" -Action {Write-Host "Disconnecting..."; $Error.Clear(); $event.MessageData.Open(); if($Error){Write-Host "Error ReConnecting.`r`n $Error `r`n`r`n"}else{Write-Host "ReConnected. `r`n`r`n"};} -MessageData $objConnection;

			$objConnection.Open();

			#$objConnection.Close();

		#>

	}


	function CleanUpConn{
		#From PS-ExchConn.ps1.
		$Session = Get-PSSession | Select Name;
		if ($Session -is [array]){
			for ($intX = 0; $intX -lt $Session.length; $intX++){
				Remove-PSSession -Name $Session[$intX].Name;
			}
		}else{
			$Session = (Get-PSSession).Name;
			if (($Session -ne "") -and ($Session -ne $null)){
				Remove-PSSession -Name $Session;
			}
		}
	}

	function CreateMailBox{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strUserName, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strAlias, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strEmail, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strExchServer, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strExchMBDB, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strDomain, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strDC, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][Boolean]$bDoBackGround, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strExchStore = "", 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Boolean]$bUpdateResults = $True
		)
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters passed in.
			#$objReturn.Results		= $True or $False.  Was a Mailbox Created, or the background job started.
			#$objReturn.Message		= The Running WorkLog ($strRunningWorkLog).
			#$objReturn.Returns		= The name of the background job, or blank if NOT run as a background process.
		#$strUserName = The SamAccountName of the user to create a MailBox for.
		#$strAlias = The Alias to use.
		#$strEmail = The PrimarySmtpAddress to give the account.
		#$strExchServer = The Exchange Server being used.
		#$strExchStore = The Exchange Storage Group being used.
		#$strExchMBDB = The MailBox DataStore to create the MailBox on.
		#$strDomain = The Domain to connect to if NOT running as a background job.
		#$strDC = The DomainController that was used to create the Account on, and that should be used to create the MailBox with.
		#$bDoBackGround = $True or $False.  Do the work in a background process.
		#$bUpdateResults = $True or $False.  Do the UpdateResults() routines.

		#Setup the PSObject to return.
		#http://stackoverflow.com/questions/21559724/getting-all-named-parameters-from-powershell-including-empty-and-set-ones
		$CommandName = $PSCmdlet.MyInvocation.InvocationName;
		$ParameterList = (Get-Command -Name $CommandName).Parameters;
		$strTemp = "";
		foreach ($key in $ParameterList.keys){
			$var = Get-Variable -Name $key -ErrorAction SilentlyContinue;
			if($var){$strTemp += "[$($var.name) = $($var.value)] ";}
		}
		$strTemp = $CommandName + "(" + $strTemp.Trim() + ")";
		$objReturn = New-Object PSObject -Property @{
			Name = $strTemp
			Results = $False
			Message = ""
			Returns = ""
		}

		$bDidMailBox = $False;
		$strJobName = "";

		if ($strUserName.EndsWith(".dev")){
			$strMessage = "Skipping MailBox creation; Dev accounts don't get Exchange MailBoxes.`r`n";
			$strRunningWorkLog = $strRunningWorkLog + $strMessage;
			if ($bUpdateResults -eq $True){
				if (Get-Command UpdateResults -errorAction SilentlyContinue){
					UpdateResults "$strMessage `r`n" $False;
				}
			}
		}
		else{
			$strMessage = "Starting Mailbox creation for user '$strUserName'.`r`n";
			#$strMessage = $strMessage + "   " + ([System.DateTime]::Now).ToString();
			if ($bUpdateResults -eq $True){
				if (Get-Command UpdateResults -errorAction SilentlyContinue){
					UpdateResults "$strMessage `r`n" $False;
				}
			}

			#Should check if a mailbox exists already.

			#$strExchServer = $dgvBulk.SelectedRows[0].Cells['ExchSvr'].Value;
			#$strExchStore = $dgvBulk.SelectedRows[0].Cells['ExchSG'].Value;
			#$strExchMBDB = $dgvBulk.SelectedRows[0].Cells['ExchMS'].Value;
			if (($strExchServer -eq "") -or ($strExchMBDB -eq "")){
				$strMessage = "Skipping Mailbox creation, because could not determin what Exch Server (and Mail Store) to create the account on.`r`n";
				$strRunningWorkLog = $strRunningWorkLog + $strMessage;
				if ($bUpdateResults -eq $True){
					if (Get-Command UpdateResults -errorAction SilentlyContinue){
						UpdateResults "$strMessage `r`n" $False;
					}
				}
			}
			else{
				#$strEmail = $txbEmail.Text.Trim();
				#$strAlias = $txbAlias.Text.Trim();
				$strOtherName = "";

				$strMessage = "Creating MailBox for '" + $strUserName + "' on '" + $strExchServer + "\";
				if (($strExchStore -eq "") -or ($strExchStore -eq $null)){
					$strMessage = $strMessage + $strExchMBDB + "'.";
				}else{
					$strMessage = $strMessage + $strExchStore + "\" + $strExchMBDB + "'.";
				}
				$strMessage = $strMessage + "`r`n";
				$strRunningWorkLog = $strRunningWorkLog + $strMessage;
				if ($bUpdateResults -eq $True){
					if (Get-Command UpdateResults -errorAction SilentlyContinue){
						UpdateResults "$strMessage `r`n" $False;
					}
				}

				$objUser = FindUser $strUserName;
				if ($objUser.DistinguishedName.Contains("OU=NNPI")){
					#if NNPI need to run the following
					if ($strUserName.EndsWith(".nnpi")){
						#for the .nnpi account ($strUserName = "first.last.nnpi")
						$strOtherName = $strUserName.SubString(0, ($strUserName.Length - 5))
					}else{
						#for the regular account ($strUserName = "first.last")
						$strOtherName = ($strUserName + ".nnpi")
					}

					$objUser = $null;
					$objUser = FindUser $strOtherName;
					if ($objUser -eq $null){
						$strOtherName = "";
					}
				}

				#Remove-MailboxPermission -Identity $strUserName -User $strOtherName -AccessRights FullAccess -Confirm:$False;
				#Add-MailboxPermission -Identity $strUserName -User $strOtherName -AccessRights FullAccess -InheritanceType All -AutoMapping $False;

				$Error.Clear();
				#$bDoBackGround = $chkBackGrndMB.Checked;
				if ($bDoBackGround -ne $True){
					#Need Exchange commands from here on...
					if ((!(Get-PSSession)) -or (!(Get-Command "Enable-Mailbox" -ErrorAction SilentlyContinue)) -or (!(Get-Command "Remove-MailboxPermission" -ErrorAction SilentlyContinue))){
						$strMessage = "  Importing Exchange commands.  " + ([System.DateTime]::Now).ToString() + "`r`n";
						if ($bUpdateResults -eq $True){
							UpdateResults $strMessage $False;
						}

						Switch ($strDomain){
							"nadsusea"{
								$objResult = SetupConn "e" "Default";
							}
							"nadsuswe"{
								$objResult = SetupConn "w" "Default";
							}
							"pads"{
								$objResult = SetupConn "p" "Default";
							}
							"nmci-isf"{
								$objResult = SetupConn "e" "Default";
							}
							default{
								$objResult = SetupConn "w" "Default";
							}
						}
						$strMessage = "  Done importing Exchange commands.  " + ([System.DateTime]::Now).ToString() + "`r`n";
						if ($bUpdateResults -eq $True){
							UpdateResults $strMessage $False;
						}
					}

					$Error.Clear();
					#From Ecxhange GUI
					#$objResult = CreateMailBox ($strDomain + "\" + $strUserName) $strExchMBDB $strAlias $True;
					if ($strAlias.EndsWith(".ctr", 1)){				#the 1 makes it case-insensitive.
						$objResult = Enable-Mailbox $strUserName -Alias $strAlias -Database $strExchMBDB -DomainController $strDC -PrimarySmtpAddress $strEmail;

						$objResult = Set-Mailbox -Identity $strUserName -DomainController $strDC -EmailAddressPolicyEnabled $False;
						$objResult = Set-Mailbox -Identity $strUserName -DomainController $strDC -EmailAddresses @{Add="$strAlias@navy.mil"}
						$objResult = Set-Mailbox -Identity $strUserName -DomainController $strDC -EmailAddresses @{Add="$strAlias@nmci-isf.com"};
						$objResult = Set-Mailbox -Identity $strUserName -DomainController $strDC -EmailAddresses @{Add="$strAlias.ctr@navy.mil"}
						$objResult = Set-Mailbox -Identity $strUserName -DomainController $strDC -EmailAddresses @{Add="$strAlias.ctr@nmci-isf.com"}
						$objResult = Set-Mailbox -Identity $strUserName -DomainController $strDC -EmailAddressPolicyEnabled $True;
					}else{
						$objResult = Enable-Mailbox $strUserName -Alias $strAlias -Database $strExchMBDB -DomainController $strDC;
					}

					if ($Error){
						$strMessage = "Error creating user Mailbox.";
						#$strMessage = $strMessage + "   " + ([System.DateTime]::Now).ToString();
						$strMessage = $strMessage + "`r`n" + $Error;
						$strMessage = "`r`n" + ("-" * 100) + "`r`n" + $strMessage + "`r`n";
					}else{
						$bDidMailBox = $True;
						if (($strOtherName -ne "") -and ($strOtherName -ne $null)){
							#Remove-MailboxPermission -Identity $strUserName -DomainController $strDC -User $strOtherName -AccessRights "FullAccess";
							$objResult = Remove-MailboxPermission -Identity $strUserName -DomainController $strDC -User $strOtherName -AccessRights "FullAccess" -Confirm:$False | Out-Null;
							#Add-MailboxPermission -Identity $strUserName -DomainController $strDC -User $strOtherName -AccessRights "FullAccess" -InheritanceType All -AutoMapping $False;
							$objResult = Add-MailboxPermission -Identity $strUserName -DomainController $strDC -User $strOtherName -AccessRights "FullAccess" -InheritanceType All -AutoMapping $False | Out-Null;
						}

						$Error.Clear();
						#Hide the email in the GAL.
						if ($strUserName.EndsWith(".nnpi", 1)){		#the 1 makes it case-insensitive.
							$objResult = Set-Mailbox -Identity $strUserName -DomainController $strDC -HiddenFromAddressListsEnabled $True;
						}

						$strMessage = "Successfully Created MailBox for '" + $strUserName + "' (Alias: " + $strAlias + ") on '" + $strExchServer + "\";
						if (($strExchStore -eq "") -or ($strExchStore -eq $null)){
							$strMessage = $strMessage + $strExchMBDB + "'.";
						}else{
							$strMessage = $strMessage + $strExchStore + "\" + $strExchMBDB + "'.";
						}
						#$strMessage = $strMessage + "   " + ([System.DateTime]::Now).ToString();
						$strMessage = $strMessage + "`r`n";
					}
				}
				else{
					$Error.Clear();
					if ($strAlias.EndsWith(".ctr", 1)){				#the 1 makes it case-insensitive.
						#if ($strUserName.EndsWith(".nnpi")){
						#	#$objJobCode = [scriptblock]::create({param($strUserName, $strAlias, $strEmail, $strMBDB, $strOpsMaster); Enable-Mailbox $strUserName -Alias $strAlias -PrimarySmtpAddress $strEmail -Database $strMBDB -DomainController $strOpsMaster; Set-Mailbox -Identity $strUserName -DomainController $strOpsMaster -EmailAddressPolicyEnabled $False; Set-Mailbox -Identity $strUserName -DomainController $strOpsMaster -EmailAddresses @{Add="$strAlias@nmci-isf.com"}; Set-Mailbox -Identity $strUserName -DomainController $strOpsMaster -EmailAddressPolicyEnabled $True; Remove-MailboxPermission -Identity $strUserName -DomainController $strOpsMaster -User ($strUserName.SubString(0, ($strUserName.Length - 5))) -AccessRights "FullAccess"; Add-MailboxPermission -Identity $strUserName -DomainController $strOpsMaster -User ($strUserName.SubString(0, ($strUserName.Length - 5))) -AccessRights "FullAccess" -InheritanceType All -AutoMapping $False;});
						#	$objJobCode = [scriptblock]::create({param($strUserName, $strAlias, $strEmail, $strMBDB, $strOpsMaster); Enable-Mailbox $strUserName -Alias $strAlias -PrimarySmtpAddress $strEmail -Database $strMBDB -DomainController $strOpsMaster; Set-Mailbox -Identity $strUserName -DomainController $strOpsMaster -EmailAddressPolicyEnabled $False; Set-Mailbox -Identity $strUserName -DomainController $strOpsMaster -EmailAddresses @{Add="$strAlias@nmci-isf.com"}; Set-Mailbox -Identity $strUserName -DomainController $strOpsMaster -EmailAddressPolicyEnabled $True; Remove-MailboxPermission -Identity $strUserName -DomainController $strOpsMaster -User $strOtherName -AccessRights "FullAccess" | Out-Null; Add-MailboxPermission -Identity $strUserName -DomainController $strOpsMaster -User $strOtherName -AccessRights "FullAccess" -InheritanceType All -AutoMapping $False | Out-Null;});
						#}else{
							$objJobCode = [scriptblock]::create({param($strUserName, $strAlias, $strEmail, $strMBDB, $strOpsMaster); Enable-Mailbox $strUserName -Alias $strAlias -PrimarySmtpAddress $strEmail -Database $strMBDB -DomainController $strOpsMaster; Set-Mailbox -Identity $strUserName -DomainController $strOpsMaster -EmailAddressPolicyEnabled $False; Set-Mailbox -Identity $strUserName -DomainController $strOpsMaster -EmailAddresses @{Add="$strAlias@nmci-isf.com"}; Set-Mailbox -Identity $strUserName -DomainController $strOpsMaster -EmailAddressPolicyEnabled $True;});
						#}
					}else{
						#if ($strUserName.EndsWith(".nnpi")){
						#	#$objJobCode = [scriptblock]::create({param($strUserName, $strAlias, $strEmail, $strMBDB, $strOpsMaster); Enable-Mailbox $strUserName -Alias $strAlias -Database $strMBDB -DomainController $strOpsMaster; Remove-MailboxPermission -Identity $strUserName -DomainController $strOpsMaster -User ($strUserName.SubString(0, ($strUserName.Length - 5))) -AccessRights "FullAccess"; Add-MailboxPermission -Identity $strUserName -DomainController $strOpsMaster -User ($strUserName.SubString(0, ($strUserName.Length - 5))) -AccessRights "FullAccess" -InheritanceType All -AutoMapping $False;});
						#	$objJobCode = [scriptblock]::create({param($strUserName, $strAlias, $strEmail, $strMBDB, $strOpsMaster); Enable-Mailbox $strUserName -Alias $strAlias -Database $strMBDB -DomainController $strOpsMaster; Remove-MailboxPermission -Identity $strUserName -DomainController $strOpsMaster -User $strOtherName -AccessRights "FullAccess" | Out-Null; Add-MailboxPermission -Identity $strUserName -DomainController $strOpsMaster -User $strOtherName -AccessRights "FullAccess" -InheritanceType All -AutoMapping $False | Out-Null;});
						#}else{
							$objJobCode = [scriptblock]::create({param($strUserName, $strAlias, $strEmail, $strMBDB, $strOpsMaster); Enable-Mailbox $strUserName -Alias $strAlias -Database $strMBDB -DomainController $strOpsMaster;});
						#}
					}
					$strJobName = "CreateMailBox_" + $strUserName;
					if (($strEmail -eq "") -or ($strEmail -eq $null)){
						$strEmail = $strAlias + "@navy.mil";
					}
					#Run the background job.
					$global:objJobs += CreateRunSpaceJob -RSPool $global:objExchPool -JobName $strJobName -JobScript $objJobCode -Arguments @($strUserName, $strAlias, $strEmail, $strExchMBDB, $strDC);
					if ($Error){
						$strMessage = "Error creating background process to create the user Mailbox.";
						#$strMessage = $strMessage + "   " + ([System.DateTime]::Now).ToString();
						$strMessage = $strMessage + "`r`n" + $Error;
						$strMessage = "`r`n" + ("-" * 100) + "`r`n" + $strMessage + "`r`n";
					}else{
						$strMessage = "Creating the MailBox in a background process... `r`n";
						$bDidMailBox = $True;
					}
				}

				$strRunningWorkLog = $strRunningWorkLog + $strMessage;
				if ($bUpdateResults -eq $True){
					if (Get-Command UpdateResults -errorAction SilentlyContinue){
						UpdateResults "$strMessage `r`n" $False;
					}
				}
			}
		}

		$objReturn.Message = $strRunningWorkLog;
		$objReturn.Results = $bDidMailBox;
		$objReturn.Returns = $strJobName;

		return $objReturn;
	}

	function EWSCreateSubscriptionPull{
		#Maybe some day.
		<#
			. C:\Projects\PS-CFW\Exchange.ps1;
			$MailboxName = "henry.schade@nmci-isf.com";


			$objActionScript = [scriptblock]::create($function:EWSOnEventDisplay);
			$objReturn = EWSCreateSubscriptionStream $MailboxName $objActionScript;
			$objReturn

		#>

	}

	function EWSCreateSubscriptionStream{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$MailboxName, 
			[ValidateNotNull()][Parameter(Mandatory=$True)]$strActionScript, 
			[ValidateNotNull()][Parameter(Mandatory=$False)]$objService, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$WhatFolder = "Inbox", 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Int]$intTimeOut = 30, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Boolean]$bolCleanEvents = $True, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strEventName = "OnNotificationEvent", 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Boolean]$bolDoErr = $True, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strDoRecon = "Notify"
		)
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters passed in.
			#$objReturn.Results		= $True or $False.  Were the Subscription to the MailBox created.
			#$objReturn.Message		= "Success" and/or the error message(s).
			#$objReturn.Returns		= An array of the SubscriptionId's created.
		#MailboxName = Email/Mailbox to Subscribe to (i.e. "henry.schade@nmci-isf.com").
		#$strActionScript = The Script/Code the Event should run.
		#$objService = The MailBox Root object.
		#$WhatFolder = What "folder", to Subscribe to. (Inbox is the only option right now.)    (i.e. "root", "inbox", "calendar", "\Inbox\folder1\folder2\folderX" etc)
		#$intTimeOut = How long to wait until a TimeOut event  (30 min default).
		#$bolCleanEvents = $True or $False.  Clean ALL existing Event Subscriptions.
		#$strEventName = The EventName to register ("OnNotificationEvent" is the default).
		#$bolDoErr = $True or $False.  Register the "OnSubscriptionError" and "OnError" events too, to write to the scrren if it happens.
		#$strDoRecon = "False", "Notify", or "Silent"  ("Notify" is default).  "Notify" writes to the screen when disconnected, and will try reconnecting, and will write if succeded or not.

		#https://msdn.microsoft.com/en-us/library/office/dn458791(v=exchg.150).aspx
		#http://stackoverflow.com/questions/21636201/how-to-unsubscribe-from-ews-push-notification-using-managed-api

		#Setup the PSObject to return.
		#http://stackoverflow.com/questions/21559724/getting-all-named-parameters-from-powershell-including-empty-and-set-ones
		$CommandName = $PSCmdlet.MyInvocation.InvocationName;
		$ParameterList = (Get-Command -Name $CommandName).Parameters;
		$strTemp = "";
		foreach ($key in $ParameterList.keys){
			$var = Get-Variable -Name $key -ErrorAction SilentlyContinue;
			if($var){$strTemp += "[$($var.name) = $($var.value)] ";}
		}
		$strTemp = $CommandName + "(" + $strTemp.Trim() + ")";
		$objReturn = New-Object PSObject -Property @{
			Name = $strTemp
			Results = $False
			Message = ""
			Returns = "";
		}

		[System.Collections.ArrayList]$arrEventIDs = @();

		if (($objService -eq "") -or ($objService -eq $null) -or ($objService.GetType().FullName -ne "Microsoft.Exchange.WebServices.Data.ExchangeService")){
			#objService was NOT passed in.  Need to create the Root of the Mailbox object.
			#$objService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList $strExchVer;
			$objResults = EWSGetFolder $MailboxName "";
			if ($objResults.Results -eq $True){
				$objService = $objResults.Returns;
			}else{
				$objReturn.Message = "Error connecting to Mailbox $MailboxName.`r`n" + $objResults.Message;
			}
		}

		if (($objService.GetType().FullName -eq "Microsoft.Exchange.WebServices.Data.ExchangeService")){
			if ($intTimeOut -lt 1){
				$intTimeOut = 30
			}

			#(Inbox is the only option right now.) (Currently ignoring $WhatFolder)
			$arrFoldersToWatch = New-Object Microsoft.Exchange.WebServices.Data.FolderId[] 1;
			$Inboxid = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $MailboxName);
				#$Inboxid = (EWSGetFolder $MailboxName "Inbox").Returns;		#Does NOT work.  Cannot convert the "Microsoft.Exchange.WebServices.Data.Folder" value of type "Microsoft.Exchange.WebServices.Data.Folder" to type "Microsoft.Exchange.WebServices.Data.FolderId".
			$arrFoldersToWatch[0] = $Inboxid;

			$Error.Clear();
			#http://gsexdev.blogspot.com/2012/05/ews-managed-api-and-powershell-how-to_28.html
			#Create the "NewMail" subscription object.
			$objSubscription = $objService.SubscribeToStreamingNotifications($arrFoldersToWatch, [Microsoft.Exchange.WebServices.Data.EventType]::NewMail);
			if ($Error){
				$objReturn.Message = "Error creating a subscription to Mailbox $MailboxName.`r`n" + $Error;
			}else{
				$objConnection = New-Object Microsoft.Exchange.WebServices.Data.StreamingSubscriptionConnection($objService, $intTimeOut);
					#exchange.SubscribeToPullNotifications(new FolderId[] { _calendarFolderId }, Settings.SubscriptionTimeout, null, EventType.Created, EventType.Deleted, EventType.Modified);
				$objConnection.AddSubscription($objSubscription);
				if ($Error){
					$objReturn.Message = "Error creating a subscription connecting to Mailbox $MailboxName.`r`n" + $Error;
				}else{
					if ($bolCleanEvents -eq $True){
						#Dealing w/ Events:
						#http://blogs.technet.com/b/heyscriptingguy/archive/2011/06/17/manage-event-subscriptions-with-powershell.aspx
						$objResults = Get-EventSubscriber;
						foreach ($objEvent in $objResults){
							#Write-Host $objEvent.SubscriptionId " - " $objEvent.EventName
							if (($objEvent -ne $null) -and ($objEvent -ne "")){
								##Unsubscribe-Event -SubscriptionId 1;   #PS3+
								#Unregister-Event -SubscriptionId 1;
								Unregister-Event -SubscriptionId $objEvent.SubscriptionId;
							}
						}
					}

					$Error.Clear();
					#Register the OnNotificationEvent() event.
					#http://gsexdev.blogspot.com/2012/05/ews-managed-api-and-powershell-how-to_28.html
						#In Powershell the Register-ObjectEvent cmdlet allows you to subscribe to the events that are generated by the Microsoft .NET Framework.
					#Register-ObjectEvent -inputObject $objConnection -eventName "OnNotificationEvent" -Action $function:EWSOnEventDisplay -MessageData $objService;
					$objResults = Register-ObjectEvent -inputObject $objConnection -eventName $strEventName -Action $strActionScript -MessageData $objService;
					if ($Error){
						$objReturn.Message = "Error registering Event '$strEventName'. `r`n" + $Error + "`r`n";
						$Error.Clear();
					}else{
						$arrEventIDs += $objResults.ID;
					}

					#Register the OnSubscriptionError() and OnError() events.
					#https://msdn.microsoft.com/en-us/library/office/dn458788(v=exchg.150).aspx#bk_recover
					if ($bolDoErr -eq $True){
						#7/16/2015 15:47:54 Subscription Error
						#$event =  System.Management.Automation.PSEventArgs
						#$event.MessageData = Microsoft.Exchange.WebServices.Data.StreamingSubscriptionConnection
						#$event.Exception.Message = ''
						#$event.SourceEventArgs = 
						#7/16/2015 15:47:54 Disconnecting...

						##$event.MessageData  -->  $objConnection
						#$arrFoldersToWatch = New-Object Microsoft.Exchange.WebServices.Data.FolderId[] 1;
						#$Inboxid = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $MailboxName);
						#$arrFoldersToWatch[0] = $Inboxid;
						#$objSubscription = $objService.SubscribeToStreamingNotifications($arrFoldersToWatch, [Microsoft.Exchange.WebServices.Data.EventType]::NewMail);
						#$objConnection.AddSubscription($objSubscription);


						#Register the OnSubscriptionError() event.
						$Error.Clear();
						#http://stackoverflow.com/questions/5911904/where-is-my-streaming-subscription-going
							#object sender  -->  $event.MessageData  -->  $objConnection  ???
						#$objResults = Register-ObjectEvent -inputObject $objConnection -eventName "OnSubscriptionError" -Action {Write-Host "Subscription Error "; $event.Exception.Message;} -MessageData $objConnection;
						$objResults = Register-ObjectEvent -inputObject $objConnection -eventName "OnSubscriptionError" -Action {Write-Host (Get-Date) " Subscription Error"; Write-Host $event.SourceEventArgs; Write-Host "`r`n"; $Error.Clear(); $event.MessageData.Open(); if($Error){Write-Host "Error ReConnecting. `r`n"; Write-Host $Error " `r`n";}else{Write-Host "ReConnected. `r`n`r`n";};} -MessageData $objConnection;
						if ($Error){
							$objReturn.Message = $objReturn.Message + "Error registering Event 'OnSubscriptionError'. `r`n" + $Error + "`r`n";
							$Error.Clear();
						}else{
							$arrEventIDs += $objResults.ID;
						}

						##Register the OnError() event.
						#$Error.Clear();
						##$objResults = Register-ObjectEvent -inputObject $objConnection -eventName "OnError" -Action {Write-Host "Error "; $event.Exception.Message;} -MessageData $objConnection;
						#$objResults = Register-ObjectEvent -inputObject $objConnection -eventName "OnError" -Action {Write-Host (Get-Date) " Error "; Write-Host "event: " $event; Write-Host ".MessageData: " $event.MessageData; Write-Host ".Exception.Message: " $event.Exception.Message;} -MessageData $objConnection;
						#if ($Error){
						#	$objReturn.Message = $objReturn.Message + "Error registering Event 'OnError'. `r`n" + $Error + "`r`n";
						#	$Error.Clear();
						#}else{
						#	$arrEventIDs += $objResults.ID;
						#}
					}

					#$strDoRecon = "False", "Notify", or "Silent"
					if ($strDoRecon -ne "False"){
						#Register the OnDisconnect() event.
						$Error.Clear();
						if ($strDoRecon -eq "Silent"){
							$objResults = Register-ObjectEvent -inputObject $objConnection -eventName "OnDisconnect" -Action {$event.MessageData.Open();} -MessageData $objConnection;
						}
						if ($strDoRecon -eq "Notify"){
							$objResults = Register-ObjectEvent -inputObject $objConnection -eventName "OnDisconnect" -Action {Write-Host (Get-Date) " Disconnecting..."; $Error.Clear(); $event.MessageData.Open(); if($Error){Write-Host "Error ReConnecting. `r`n"; Write-Host $Error " `r`n";}else{Write-Host "ReConnected. `r`n`r`n";};} -MessageData $objConnection;
						}

						if ($Error){
							$objReturn.Message = $objReturn.Message + "Error registering Event 'OnDisconnect' in '$strDoRecon' mode. `r`n" + $Error + "`r`n";
							$Error.Clear();
						}else{
							$arrEventIDs += $objResults.ID;
						}
					}

					if ($objConnection.IsOpen){
						$objConnection.Close();
					}

					$Error.Clear();
					$objConnection.Open();
					if ($Error){
						$objReturn.Message = $objReturn.Message + "Error opening the Subscription connection. `r`n" + $Error + "`r`n";
						$Error.Clear();
					}else{
						#Made a connection, so lets return the results.
						$objReturn.Results = $True;
						$objReturn.Returns = $arrEventIDs;
						if ($objReturn.Message -ne ""){
							$objReturn.Message = "Success, with errors. `r`n" + $objReturn.Message;
						}else{
							$objReturn.Message = "Success";
						}
					}
				}
			}
		}

		return $objReturn;

	}

	function EWSGetEmails{
		Param(
			#[ValidateNotNull()][Parameter(Mandatory=$True)][String]$MailboxName, 
			#[ValidateNotNull()][Parameter(Mandatory=$False)][String]$WhatFolder = "", 
			[ValidateNotNull()][Parameter(Mandatory=$True,ParameterSetName="FromScratch")][String]$MailboxName, 
			[ValidateNotNull()][Parameter(Mandatory=$False,ParameterSetName="FromScratch")][String]$WhatFolder = "Inbox", 
			[ValidateNotNull()][parameter(Mandatory=$True,ParameterSetName="Existing")]$objEWSFolder, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Int]$NumToRet = 0, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Boolean]$bolDoProperties = $True, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Boolean]$bolTryAuto = $True, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strExchVer = "Exchange2010_SP1", 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Boolean]$bolAltCreds = $False
		)
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters passed in.
			#$objReturn.Results		= The # of MailItems found/returned.
			#$objReturn.Message		= "Success" or the error message.
			#$objReturn.Returns		= The MailItems objects/collection.
		#--=== Provide ("$MailboxName" and "$WhatFolder"), or ("$objEWSFolder"). ===---
		#$MailboxName = Email/Mailbox to connect to.  (i.e. "henry.schade@nmci-isf.com")
		#$WhatFolder = What "folder", to return.  (i.e. "root", "inbox", "calendar", "\Inbox\folder1\folder2\folderX", etc)
		#$objEWSFolder = The EWS Mail Folder object.
		#$NumToRet = How many emails to return.  [0 is default (All)].
		#$bolDoProperties = Do the ::FirstClassProperties before returning.
		#$bolTryAuto = $True or $False.  Use the .AutodiscoverUrl property to try and find the Exchange URL for the ExchangeService object.
		#$strExchVer = The version of Exchange that is being connected to. ("Exchange2007_SP1", or "Exchange2010", or "Exchange2010_SP1", or "Exchange2010_SP2", or "Exchange2013")
		#$bolAltCreds = $True or $False.  Prompt for alternate credentials.

		#Setup the PSObject to return.
		#http://stackoverflow.com/questions/21559724/getting-all-named-parameters-from-powershell-including-empty-and-set-ones
		$CommandName = $PSCmdlet.MyInvocation.InvocationName;
		$ParameterList = (Get-Command -Name $CommandName).Parameters;
		$strTemp = "";
		foreach ($key in $ParameterList.keys){
			$var = Get-Variable -Name $key -ErrorAction SilentlyContinue;
			if($var){$strTemp += "[$($var.name) = $($var.value)] ";}
		}
		$strTemp = $CommandName + "(" + $strTemp.Trim() + ")";
		$objReturn = New-Object PSObject -Property @{
			Name = $strTemp
			Results = $False
			Message = "Error"
			Returns = "";
		}
		#Write-Host $strTemp;

		if ((($objEWSFolder -eq "") -or ($objEWSFolder -eq $null))){
			if ((($MailboxName -ne "") -and ($MailboxName -ne $null))){
				$objResults = EWSGetFolder $MailboxName $WhatFolder $bolTryAuto $strExchVer $bolAltCreds;
				#$objResults | FL;
				#Write-Host $objResults.Message;
				$objFolder = $objResults.Returns;
			}else{
				$objFolder = $null;
				$objReturn.Message = "Error, no mailbox specified.";
			}
		}else{
			$objFolder = $objEWSFolder;
		}

		if ($objFolder -ne $null){
			#$objReturn = EWSGetEmails -MailboxName $MailboxName -WhatFolder $strFolder -NumToRet $intNumGet -bolDoProperties $False;

			#There is a 1000 Item limit.  Need to accomodate for it.
			#http://blogs.msdn.com/b/akashb/archive/2011/07/29/another-example-of-using-ews-managed-api-1-1-from-powershell-impersonation-searchfilter-finditems-paging.aspx
			#if ($objFolder.TotalCount -gt 1000){
			#	#http://gsexdev.blogspot.com/2012/02/ews-managed-api-and-powershell-how-to.html
			#	#Define ItemView to retrive just 1000 Items  
			#	$objItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000);
			#	$objItems = $objFolder.FindItems($objFolder.Id, $objItemView);
			#	do{
			#		if ($objFolder.MoreAvailable -eq $True){
			#			$objItems = $objItems + $objFolder.FindItems($objFolder.Id, $objItemView);
			#			$objItemView.Offset += $objItems.Items.Count;
			#		}
			#	} while (($objFolder.MoreAvailable -eq $True) -and ($objItems.Items.Count -lt $NumToRet));

			#}else{
				if (($objFolder.TotalCount -gt $NumToRet) -and ($NumToRet -gt 0)){
					$objItems = $objFolder.FindItems($NumToRet);
				}elseif($objFolder.TotalCount -eq 0){
					$objItems = $null;
				}else{
					$objItems = $objFolder.FindItems($objFolder.TotalCount);
				}
			#}
			#Write-Host @($objItems).Count;

			if (($bolDoProperties -eq $True) -and ($objReturn.Results -ne 0)){
				#create a property set (to let us access the body & other details not available from the FindItems call)
				$objPropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties);
				$objPropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text;

				foreach ($item in $objItems.Items){
					# load the ::FirstClassProperties.
					$item.load($objPropertySet);
				}
			}

			$objReturn.Message = "Success";
			#$objReturn.Results = $True;
			if ($objItems -ne $null) {
				$objReturn.Results = @($objItems).Count;
			}else{
				$objReturn.Results = 0;
			}
			$objReturn.Returns = $objItems;
		}else{
			$objReturn.Results = $False;
			$objReturn.Returns = $null;
		}


		#http://www.garrettpatterson.com/2014/04/18/checkread-messages-exchangeoffice365-inbox-with-powershell/
		#http://stackoverflow.com/questions/4454165/how-to-check-an-exchange-mailbox-via-powershell
		#http://mellositmusings.com/2013/10/29/powershell-script-to-download-attachments-from-an-email/
		#http://gsexdev.blogspot.com/

		return $objReturn;

	}

	function EWSGetFolder{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$MailboxName, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$WhatFolder = "Inbox", 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Boolean]$bolTryAuto = $True, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strExchVer = "Exchange2010_SP1", 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Boolean]$bolAltCreds = $False, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Boolean]$bolTrace = $False
		)
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters passed in.
			#$objReturn.Results		= $True or $False.  Was a connection to the MailBox created.
			#$objReturn.Message		= "Success, found" or "Success, installed" or the error message.
			#$objReturn.Returns		= The requested "folder".  (.ExchangeService by default)
		#MailboxName = Email/Mailbox to connect to.  (i.e. "henry.schade@nmci-isf.com")
		#$WhatFolder = What "folder", to return. (Still need to code all the .WellKnownFolderName values)  (i.e. "root", "inbox", "calendar", "\Inbox\folder1\folder2\folderX" etc)
		#$bolTryAuto = $True or $False.  Use the .AutodiscoverUrl property to try and find the Exchange URL for the ExchangeService object.
		#$strExchVer = The version of Exchange that is being connected to. ("Exchange2007_SP1", or "Exchange2010", or "Exchange2010_SP1", or "Exchange2010_SP2", or "Exchange2013")
		#$bolAltCreds = $True or $False.  Prompt for alternate credentials.
		#$bolTrace = $True or $False.  Turn $objExchServ.TraceEnabled on/off, for troubleshooting.

		#Folder properties:
		#https://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.folder_properties(v=exchg.80).aspx
		#https://msdn.microsoft.com/en-us/library/dn567668(v=exchg.150).aspx

		#Setup the PSObject to return.
		#http://stackoverflow.com/questions/21559724/getting-all-named-parameters-from-powershell-including-empty-and-set-ones
		$CommandName = $PSCmdlet.MyInvocation.InvocationName;
		$ParameterList = (Get-Command -Name $CommandName).Parameters;
		$strTemp = "";
		foreach ($key in $ParameterList.keys){
			$var = Get-Variable -Name $key -ErrorAction SilentlyContinue;
			if($var){$strTemp += "[$($var.name) = $($var.value)] ";}
		}
		$strTemp = $CommandName + "(" + $strTemp.Trim() + ")";
		$objReturn = New-Object PSObject -Property @{
			Name = $strTemp
			Results = $False
			Message = "Error"
			Returns = "";
		}

		if ($WhatFolder.Contains("\")){
			if (!($WhatFolder.StartsWith("\"))){
				$WhatFolder = "\" + $WhatFolder;
			}
		}

		$objResults = EWSVerifyInstall;
		if ($objResults.Results -eq $True){
			#$dllPath = "C:\Program Files (x86)\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll";
			#$dllPath = "C:\Program Files (x86)\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"
			#$dllPath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
			$dllFile = "Microsoft.Exchange.WebServices.dll";

			$dllPath = $objResults.Returns + $dllFile;
			[void][Reflection.Assembly]::LoadFile($dllPath);
				#Import-Module -Name $dllPath;
				#Add-Type -Path $dllPath;

			#$objExchServ = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1);
			#$objExchServ = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013);
			$objExchServ = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList $strExchVer;
			#$objExchServ.TraceEnabled = $False;
			$objExchServ.TraceEnabled = $bolTrace;
			#if ($bolTrace){
			#	#$objExchServ.TraceFlags = TraceFlags.All;
			#}

			#$objExchServ.UseDefaultCredentials = $True;
			if ($MailboxName.Contains(($env:username + "@"))){
				#Targeting an on-premises Exchange server and your client is domain joined.
				$objExchServ.UseDefaultCredentials = $True;
			}else{
				#Connect using impersonation / alternate credentials.
				#$Credential = Get-Credential -Credential "SelectAnAccount" | Out-Null;
				$Credential = $host.ui.PromptForCredential("Need Outlook Credentials", "Please select/enter your Outlook information below.", "", "NetBiosUserName") | Out-Null;
				#Set the credentials for Exchange
				$objExchServ.UseDefaultCredentials = $False;
				$objExchServ.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $Credential.UserName, $Credential.GetNetworkCredential().Password;
			}

			$Error.Clear();
			#Determine/Set the EWS endpoint
			if ($bolTryAuto){
				$objExchServ.AutodiscoverUrl($MailboxName, {$True}) | Out-Null;			#This is taking FOREVER.  Looks like it is setting the URL of the ExchangeService object, which is required for subscriptions.
			}
			if (($Error) -or ($bolTryAuto -ne $True)){
				$Error.Clear();
				$strURL = [system.URI] "https://cas-sdni.nadsuswe.nads.navy.mil/ews/exchange.asmx";
				#$strURL = [system.URI] "https://cas-sdni.nadsuswe.nads.navy.mil/PowerShell/";

					##Get all the Exchange Servers in the Domain (filter the results).
					#$objExchServers = GetExchangeServers | where {(($_.FQDN -match "nmci-isf") -and (($_.Roles -match 4) -or ($_.Roles -match 36)))};
					##select a random server from the list
					#if($objExchServers.Count -gt 1) {
					#	$intRandom = Get-Random -Minimum 0 -Maximum $objExchServers.Count;
					#	$strServer = $objExchServers[$intRandom].FQDN;
					#}else{
					#	$strServer = $objExchServers[0].FQDN;
					#}
					##"http://$strServer/PowerShell/" -Authentication Kerberos
					#$strURL = [system.URI] "http://$strServer/PowerShell/";
					#$strURL = [system.URI] "http://NMCINRFKXF01V.nmci-isf.com/PowerShell/";		#Unauthorized
					#$strURL = [system.URI] "http://NMCINRFKXF02V.nmci-isf.com/PowerShell/";		#Unauthorized

				$objExchServ.Url = $strURL;
			}

			#if (($WhatFolder -eq "") -or ($WhatFolder -eq $null)){
			#	$objReturn.Returns = $objExchServ;
			#	$objReturn.Results = $True;
			#	$objReturn.Message = "Success";
			#}else{
				$Error.Clear();
				$bolFoundOne = $False;
				$objFolder = $null;
				if ($MailboxName.Contains(($env:username + "@"))){
					if ($WhatFolder -eq "Inbox"){
						$objFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($objExchServ, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox);
						$bolFoundOne = $True;
						$objReturn.Message = "Success, " + $WhatFolder;
					}
					if ($WhatFolder -eq "Calendar"){
						$objFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($objExchServ, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar);
						$bolFoundOne = $True;
						$objReturn.Message = "Success, " + $WhatFolder;
					}
					if ($WhatFolder -eq "SentItems"){
						$objFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($objExchServ, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems);
						$bolFoundOne = $True;
						$objReturn.Message = "Success, " + $WhatFolder;
					}
					if ($WhatFolder -eq ""){
						$objFolder = $objExchServ;
						$bolFoundOne = $True;
						$objReturn.Message = "Success, .ExchangeService";
					}
				}
				if ($Error){
					$objReturn.Message = "Error, connecting to " + $MailboxName + "\" + $WhatFolder + ". `r`n" + $Error;
				}
				
				if (($bolFoundOne -ne $True) -or ($WhatFolder -eq "Root") -or ($WhatFolder.Contains("\"))){
					$Error.Clear();
					#http://gsexdev.blogspot.com/2012/01/ews-managed-api-and-powershell-how-to_23.html
					#Bind to the MSGFolder Root
					if ($MailboxName.Contains(($env:username + "@"))){
						$folderid = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot);
					}else{
						$folderid = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $MailboxName);
						if (!($WhatFolder.Contains("\"))){
							$WhatFolder = "\" + $WhatFolder;
						}
					}

					$objFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($objExchServ, $folderid);
					if ($Error){
						$objReturn.Message = "Error, connecting to " + $MailboxName + "" + $WhatFolder + ". `r`n" + $Error;
					}else{
						$objReturn.Message = "Success, " + $MailboxName + "\";

						if ($WhatFolder.Contains("\")){
							#Split the Search path into an array
							$arrFolders = $WhatFolder.Split("\");
							#Loop through the Split Array and do a Search for each level of folder
							for ($intX = 1; $intX -lt $arrFolders.Length; $intX++){
								#Perform search based on the displayname of each folder level
								$objFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1);
								$objSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $arrFolders[$intX]);
								$objFolderResults = $objExchServ.FindFolders($objFolder.Id, $objSearchFilter, $objFolderView);
								if ($objFolderResults.TotalCount -gt 0){
									foreach($folder in $objFolderResults.Folders){
									   $objFolder = $folder;
									   $objReturn.Message = $objReturn.Message + $objFolder.DisplayName + "\";
									   break;
								   }
								}else{
									#Write-Host "Error Folder Not Found";
									#$objFolder = $null;
									#$objReturn.Message = "Error Folder Not Found.";
									break;
								}
							}  
						}
					}
				}

				#if $WhatFolder was provided, but not found, then MsgFolderRoot is returned.
				$objReturn.Returns = $objFolder;
				if ($objReturn.Message.Contains("Error")){
					$objReturn.Results = $False;
				}else{
					$objReturn.Results = $True;
				}
			#}
		}else{
			$objReturn.Message = "Error, EWS dll is not installed.";
		}

		return $objReturn;
	}

	function EWSOnEventDisplay{
		#http://blogs.technet.com/b/heyscriptingguy/archive/2011/06/17/manage-event-subscriptions-with-powershell.aspx
			#automatic variables --> $event, $eventSubscriber, $sender, $sourceEventArgs, and $sourceArgs

		$intDispLen = 105;
		foreach ($objNotificationEvent in $event.SourceEventArgs.Events){
			#Write-Host "EventType: " $objNotificationEvent.EventType "`r`n";

			#if ($objNotificationEvent.EventType -eq "NewMail"){
			#	Write-Host "ItemId: " $objNotificationEvent.ItemId "`r`n";
			#	#Next is same value as above.
			#	#Write-Host "ItemId 2: " $objNotificationEvent.ItemId.UniqueId "`r`n";
			#}else{
			#	Write-Host "FolderId: " $objNotificationEvent.FolderId.UniqueId "`r`n";
			#}

			#To get more information about what caused the notification you need to bind to this Item Id.
			[String]$strItemId = $objNotificationEvent.ItemId.UniqueId.ToString();

			$objMessage = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($event.MessageData, $strItemId);
			$strSender = $objMessage.From.Name;
			$strFor = $objMessage.DisplayTo;
			if ($strFor.Length -gt $intDispLen){
				$strFor = $strFor.Trim().SubString(0, $intDispLen);
				$strFor = $strFor + "...";
			}
			$strSubject = $objMessage.Subject;
			if ($strSubject.Length -gt $intDispLen){
				$strSubject = $strSubject.Trim().SubString(0, $intDispLen);
				$strSubject = $strSubject + "...";
			}

			#create a property set (to let us access the body & other details not available from the FindItems call)
			$objPropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties);
			$objPropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text;
			# load the property set to allow us to get to the body
			$objMessage.load($objPropertySet);
			$strBody = $objMessage.Body.Text;
			$strBody = $strBody -Replace '\s+', ' ';
			if (($strBody -ne "") -and ($strBody -ne $null)){
				if ($strBody.Length -gt $intDispLen){
					$strBody = $strBody.Trim().SubString(0, $intDispLen);
					$strBody = $strBody + "...";
				}
			}else{
				$strBody = "[" + $objMessage.ItemClass + "] ";
				if ($objMessage.ItemClass -eq "IPM.Note"){
					$strBody = "[blank]";
				}
			}

			$strEmailDate = [String]$objMessage.DateTimeReceived;

			$strOutput = "`r`n" + $objNotificationEvent.EventType + " --> " + $objMessage.ItemClass + "`r`n";
				#https://msdn.microsoft.com/en-us/library/Ee200767(v=EXCHG.80).aspx
					#$objMessage.ItemClass:
				#IPM.Note										#From HP (Not dig signed.
				#IPM.Note.SMIME.MultipartSigned					#From NMCI (was Dig signed)
				#IPM.Note.SMIME									#From NMCI (Encrypted, may or may NOT be signed) (Can't read Body.   Body: [IPM.Note.SMIME])
				#IPM.Schedule.Meeting.Request
				#IPM.Schedule.Meeting.Canceled
			Write-Host $strOutput;

			$strOutput = "          " + [String](Get-Date) + "`r`n" + "Received: " + $strEmailDate + "`r`n" + "From:     " + $strSender + "`r`n" + "To:       " + $strFor + "`r`n" + "Subject:  " + $strSubject + "`r`n" + "Body:     " + $strBody + "`r`n`r`n";
			Write-Host $strOutput;

		}
	}

	function EWSVerifyInstall{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strDesiredVer = "1.2"
		)
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters passed in.
			#$objReturn.Results		= $True or $False.  Was the *.dll found/installed.
			#$objReturn.Message		= "Success, found" or "Success, installed" or the error message.
			#$objReturn.Returns		= The path where the *.dll was found/installed.
		#Verifies that "Microsoft.Exchange.WebServices.dll" is installed.  Installs it if not present.
		#$strDesiredVer = The desired version to check/install.  ("1.2", "2.1")

		#(2.0) Microsoft.Exchange.WebServices dll  ->  http://www.microsoft.com/en-us/download/details.aspx?id=35371
		#(Latest [2.2]) Microsoft.Exchange.WebServices dll  ->  http://go.microsoft.com/fwlink/?LinkId=255472

		#Setup the PSObject to return.
		#http://stackoverflow.com/questions/21559724/getting-all-named-parameters-from-powershell-including-empty-and-set-ones
		$CommandName = $PSCmdlet.MyInvocation.InvocationName;
		$ParameterList = (Get-Command -Name $CommandName).Parameters;
		$strTemp = "";
		foreach ($key in $ParameterList.keys){
			$var = Get-Variable -Name $key -ErrorAction SilentlyContinue;
			if($var){$strTemp += "[$($var.name) = $($var.value)] ";}
		}
		$strTemp = $CommandName + "(" + $strTemp.Trim() + ")";
		$objReturn = New-Object PSObject -Property @{
			Name = $strTemp
			Results = $False
			Message = "Error"
			Returns = "";
		}

		#http://www.getautomationmachine.com/en/company/news/item/embedding-files-in-powershell-scripts
		#Talks about using base85 for smaller data --> http://trevorsullivan.net/2012/07/24/powershell-embed-binary-data-in-your-script/

		if (!(Get-Command "GetPathing" -ErrorAction SilentlyContinue)){
			$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;
			if ((Test-Path ($ScriptDir + "\Common.ps1"))){
				. ($ScriptDir + "\Common.ps1")
			}
		}
		$strFileFile = (GetPathing "CFW").Returns.Rows[0]['Path'];
		if ([String]::IsNullOrWhiteSpace($strFileFile)){
			$strFileFile = "\\nawesdnifs101v.nadsuswe.nads.navy.mil\NMCIISF02$\ITSS-Tools\PS-CFW\EWS-Files.txt";
		}
		else{
			$strFileFile = $strFileFile + "EWS-Files.txt";
		}

		if (($strDesiredVer -eq "1.2") -or ($strDesiredVer -eq "2.1") -or ($strDesiredVer -eq "2.2")){
			$strFileName = "Microsoft.Exchange.WebServices.dll";
			$arrFiles = @();

			$strFileDLL = "";
			$strFileDLLAuth = "";
			$strFileXML = "";
			$strFileXMLAuth = "";
			$strFileReadme = "";
			$strFileLiscense = "";
			$strFileRedist = "";

			if ($strDesiredVer -eq "1.2"){
				$strFilePath = "C:\Program Files (x86)\Microsoft\Exchange\Web Services\1.2\";
				#1.2 installed from "http://www.microsoft.com/en-us/download/details.aspx?id=28952", using the defaults.
					#GettingStarted.doc, License Terms.rtf, Microsoft.Exchange.WebServices.dll, Microsoft.Exchange.WebServices.xml, README.htm, Redist.txt
				#My nmci system existing path was "C:\Program Files (x86)\Microsoft\"
				#All the files being encoded in here makes this *.ps1 file HUGH, so moved the encoded files into a *.txt file to be read as needed.
			}
			if ($strDesiredVer -eq "2.1"){
				$strFilePath = "C:\Program Files (x86)\Microsoft\Exchange\Web Services\2.1\";
				#On my Win8 system at home these files were already installed.
					#License Terms.rtf, Microsoft.Exchange.WebServices.Auth.dll, Microsoft.Exchange.WebServices.Auth.xml, Microsoft.Exchange.WebServices.dll, 
					#Microsoft.Exchange.WebServices.xml, README.htm, Redist.txt
				#All the files being encoded in here makes this *.ps1 file HUGH, so moved the encoded files into a *.txt file to be read as needed.
			}
			if ($strDesiredVer -eq "2.2"){
				#I have not found an install of this ver to be able to set this up yet.
				$strFilePath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\";
			}

			#Check if the *.dll exists
			if (!(Test-Path -Path ($strFilePath + $strFileName))){
				#Read the Encrypted file strings out of $strEWSFiles.  Keeping the strings in here was making the *.ps1 take a while to load.
				foreach ($strLine in [System.IO.File]::ReadAllLines($strFileFile)) {
					$strFile = "";
					$strData = "";
					if ($strLine.StartsWith($strDesiredVer)){
						$strFile = $strLine.SubString(3, $strLine.IndexOf("=") - 4).Trim();
						$strData = $strLine.SubString($strLine.IndexOf("=") + 1).Trim();
						#Write-Host $strLine.SubString(0, 90);
						#Write-Host $strFile;
						#Write-Host $strData.SubString(0, 90);
					}
					if (($strFile -ne "") -and ($strData -ne "")){
						Switch ($strFile){
							"FileDLL"{
								$strFileDLL = $strData;
							}
							"FileDLLAuth"{
								$strFileDLLAuth = $strData;
							}
							"FileXML"{
								$strFileXML = $strData;
							}
							"FileXMLAuth"{
								$strFileXMLAuth = $strData;
							}
							"FileReadme"{
								$strFileReadme = $strData;
							}
							"FileLiscense"{
								$strFileLiscense = $strData;
							}
							"FileRedist"{
								$strFileRedist = $strData;
							}
						}
					}
				}

				#Now populate the array of files.
				$arrFiles = @($strFileDLL, $strFileDLLAuth, $strFileXMLAuth, $strFileXML, $strFileReadme, $strFileLiscense, $strFileRedist);
				$arrFileNames = @("Microsoft.Exchange.WebServices.dll", "Microsoft.Exchange.WebServices.Auth.dll", "Microsoft.Exchange.WebServices.Auth.xml", "Microsoft.Exchange.WebServices.xml", "README.htm", "License Terms.rtf", "Redist.txt");
				$strInstalledFiles = "";

				if ($arrFiles.Count -gt 0){
					$objReturn.Message = "";
					if (!(Test-Path -Path ($strFilePath))){
						$Error.Clear();
						$strResults = mkdir $strFilePath;
					}
					if (Test-Path -Path ($strFilePath)){
						for ($intX = 0; $intX -lt $arrFiles.Count; $intX++){
							if ($arrFiles[$intX] -ne ""){
								$Error.Clear();
								$Content = [System.Convert]::FromBase64String($arrFiles[$intX]);
								$strReturn = Set-Content -Path ($strFilePath + $arrFileNames[$intX]) -Value $Content -Encoding Byte;
								if ($Error){
									$objReturn.Message = $objReturn.Message + "Error installing file '" + $arrFileNames[$intX] + "'.`r`n  " + $Error + "`r`n";
								}else{
									$strInstalledFiles = $strInstalledFiles + " " + $arrFileNames[$intX];
								}
							}else{
								if (!(Test-Path -Path ($strFilePath + $strFileName))){
									$objReturn.Message = $objReturn.Message + "Error data stream for file '" + $arrFileNames[$intX] + "' is blank.`r`n";
								}
							}
						}
					}else{
						$objReturn.Message = "Error creating the directory.`r`n  " + $Error + "`r`n";
					}

					if (Test-Path -Path ($strFilePath + $strFileName)){
						$objReturn.Results = $True;
						$objReturn.Returns = $strFilePath;
						if ($objReturn.Message -eq ""){
							$objReturn.Message = "Success, installed.";
						}else{
							$objReturn.Message = "Success, installed with problems. `r`n" + $objReturn.Message;
						}
					}else{
						$objReturn.Message = "Error, install failed.`r`n";
						if ($strInstalledFiles.Trim() -eq ""){
							$objReturn.Message = $objReturn.Message + "Error, No files were installed. `r`n";
						}
					}
				}else{
					$objReturn.Message = "Error creating array of files to be installed. `r`n";
				}
			}else{
				$objReturn.Results = $True;
				$objReturn.Message = "Success, found.";
				$objReturn.Returns = $strFilePath;
			}
		}

		return $objReturn;
	}

	function GetExchangeServers{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$False)][Boolean]$bolRetRoleNames = $False
		)
		#$bolRetRoleNames = Return Sever Role Names, rather than the #.

		#http://www.mikepfeiffer.net/2010/04/find-exchange-servers-in-the-local-active-directory-site-using-powershell/
		#Sample usage, with filtering by site code:
			#$objExchServers = GetExchangeServers | where { $_.FQDN -match "nrfk" }
			#$objExchServers = GetExchangeServers | where { $_.FQDN -match "nmci-isf" }

		#Using [ADSI] does NOT require the PS AD module be loaded first.

		if (($strSearch4 -eq $null) -or ($strSearch4 -eq "")){
			$strSearch4 = "";
		}
		#if (($bolRetRoleNames -eq $null) -or ($bolRetRoleNames -eq "") -or (($bolRetRoleNames -ne $True) -and ($bolRetRoleNames -ne $False))){
		#	$bolRetRoleNames = $False;
		#}

		#For the msexchcurrentserverroles (Roles) values
		$arrRoles = @{
			2  = "Mailbox Role"
			4  = "Client Access Role"
			16 = "Unified Messaging Role"
			32 = "Hub Transport Role"
			64 = "Edge Transport Role"
		};

		#Domain Initalization
		#$ADSite = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite];
		#$siteDN = $ADSite::GetComputerSite().GetDirectoryEntry().distinguishedName;
		$configNC = ([ADSI]"LDAP://RootDse").configurationNamingContext;

		$objSearcher = New-Object DirectoryServices.DirectorySearcher([ADSI]"LDAP://$configNC");
		$objectClass = "objectClass=msExchExchangeServer";
		$version = "versionNumber>=1937801568";
		$objSearcher.Filter = "(&($objectClass)($version))";
		$objSearcher.PageSize = 1000;
		[void] $objSearcher.PropertiesToLoad.Add("name");
		[void] $objSearcher.PropertiesToLoad.Add("msexchcurrentserverroles");
		[void] $objSearcher.PropertiesToLoad.Add("msexchserverrole");
		[void] $objSearcher.PropertiesToLoad.Add("serverrole");
		[void] $objSearcher.PropertiesToLoad.Add("networkaddress");
		$objSearcher.FindAll() | %{
			$strName = $_.Properties.name[0];
			$strFQDN = $_.Properties.networkaddress | %{if ($_ -Match "ncacn_ip_tcp") {$_.Split(":")[1]}};
			$strRoles = $_.Properties.msexchcurrentserverroles[0];
			if ($bolRetRoleNames -eq $True){
				$strRoles = ($arrRoles.keys | ?{$_ -band $strRoles} | %{$arrRoles.Get_Item($_)}) -join ", ";
			}
			#$strRoles = $strRoles + " (" +  $_.Properties.serverrole + ")";

			#This next block creates a PowerShell object that gets returned.
			New-Object PSObject -Property @{
				Name = $strName;
				FQDN = $strFQDN;
				Roles = $strRoles;
			}
		}
	}

	function MoveEmail{
		#Moving emails between folders
		#https://msdn.microsoft.com/en-us/magazine/dn189202.aspx

		#http://gsexdev.blogspot.com/2012/02/ews-managed-api-and-powershell-how-to_22.html
			#$objReturn.Returns.Items[0].Move($TargetFolderObject.Returns.ID)

	}

	function ReadEmail_1{
		#Read email folders and emails.  (Has some issues)
		#http://blogs.technet.com/b/heyscriptingguy/archive/2011/05/26/use-powershell-to-data-mine-your-outlook-inbox.aspx
		#http://www.computerperformance.co.uk/powershell/powershell_outlook_email.htm
		#https://msdn.microsoft.com/en-us/magazine/dn189202.aspx

		#Works w/ my Regular account.  No popups.
		Add-Type -Assembly "Microsoft.Office.Interop.Outlook" | Out-Null;
		$objFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type];
		$objOutlook = New-Object -ComObject Outlook.Application;
		$objNamespace = $objOutlook.GetNameSpace("MAPI")

		#Get Folder Count.
		$objNamespace.Folders.Count;
		#Get all Folder Names.
		$objNamespace.Folders | Select Name;
		#Get details/info about the first folder.
		$objNamespace.Folders.Item(1);
		#Get details/info about the last folder.
		$objNamespace.Folders.Item($objNamespace.Folders.Count);

		#Get details/info about the first folder, in the first folder.
		$objNamespace.Folders.Item(1).Folders.Item(1);
		#Get details/info about the first folder, in the last folder.
		$objNamespace.Folders.Item($objNamespace.Folders.Count).Folders.Item(1);
		#Get the name of the first folder, in the last folder.
		$objNamespace.Folders.Item($objNamespace.Folders.Count).Folders.Item(1).Name;

		#Inbox, Folders
		#List all the Folders in the Inbox of the last Folder.
		$objNamespace.Folders.Item($objNamespace.Folders.Count).Folders.Item('Inbox').Folders;
		#List all the Folders (Name only) in the Inbox of the last Folder.
		$objNamespace.Folders.Item($objNamespace.Folders.Count).Folders.Item('Inbox').Folders | Select Name;

		#List all the Email in the Camping folder, in the Inbox folder, of the Folder "Henry.Schade@nmci-isf.com".
		$objNamespace.Folders.Item("Henry.Schade@nmci-isf.com").Folders.Item('Inbox').Folders.Item('Camping').Items;
		#List all the Email (Subjects) in the Camping folder, in the Inbox folder, of the Folder "Henry.Schade@nmci-isf.com".
		$objNamespace.Folders.Item("Henry.Schade@nmci-isf.com").Folders.Item('Inbox').Folders.Item('Camping').Items | Select Subject;

		#Inbox, Email method 1
		$objInbox = $objNamespace.getDefaultFolder($objFolders::olFolderInBox);
		#Display ALL the data about Each email.
		$objInbox.Items
		#Display the specified data about Each email.
		$objInbox.Items | Select-Object -Property Subject, Size, ReceivedTime, Sender;

		#Inbox, Email method 2
		$strFolder = "InBox"
		#Get all the Email in the Inbox of the first Folder.
		$objEmails = $objNamespace.Folders.Item(1).Folders.Item($strFolder).Items;
		#Get all the Email in the Inbox of the last Folder.
		$objEmails = $objNamespace.Folders.Item($objNamespace.Folders.Count).Folders.Item($strFolder).Items;
		#Sort the emails gotten in the line above by Unique Sender, and output it Formatted in a Table.
			#Depending on the # of emails this can be SUPER slow/time consuming.
		#$objEmails | Sort-Object Sender -Unique | FT;
		#Display the specified data about Each email.
		$objEmails | Select-Object -Property Subject, Size, ReceivedTime, Sender;

		#Inbox, Email method 3
		$objEmails = $objNamespace.Folders.Item($objNamespace.Folders.Count).Folders.Item('Inbox').Folders.Item('Camping').Items;
		$objEmails.Item(2);
		$objEmails.Item(2).Subject;
		$objEmails.Item(2).SenderName;

	}

	function ReadEmail_2{
		#Gets/returns all Email up to a # of days old.
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$EmailAccount, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Int]$NumDays = 1
		)
		#$EmailAccount = The account to get emails from.
		#$NumDays = The # of days worth of email to get.  Default is 1 day.

		#$EmailAccount = "Some.User@Company.com";

		#http://stackoverflow.com/questions/22940569/powershell-script-against-outlook-mail-box-is-selecting-old-data
		$objOutlook = New-Object -ComObject Outlook.Application
		#$objNamespace = $objOutlook.GetNamespace("MAPI") | ?{$_.SMTP -Match $EmailAccount}
		$objNamespace = $objOutlook.GetNamespace("MAPI")
		$objAccount = $objNamespace.Folders | ?{$_.Name -Match $EmailAccount}
		if ($objAccount.Count -gt 1){
			for ($intX = 0; $intX -lt $objAccount.Count; $intX++){
				if ($objAccount[$intX].Name -eq $EmailAccount){
					$objAccount = $objAccount[$intX];
					break;
				}
			}
		}
		$objInbox = $objAccount.Folders | ?{$_.Name -Match "Inbox"}
		$dteDaysOld = (Get-Date).AddDays((-1 * $NumDays))
		$objEmails = $objInbox.Items | ?{$_.ReceivedTime -gt $dteDaysOld}

		return $objEmails;
	}

	function SendEmailEWS{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$EmailTo, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$EmailFrom, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$EmailSubject, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$EmailBody, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$EmailCC, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$EmailBCC, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$EmailAttach, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Boolean]$AsHTML = $False, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Boolean]$bolTryAuto = $True, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strExchVer = "Exchange2010_SP1"
		)
		#http://blogs.technet.com/b/heyscriptingguy/archive/2011/09/24/send-email-from-exchange-online-by-using-powershell.aspx

		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters passed in.
			#$objReturn.Results		= $True or $False.  Was an email composed/sent.
			#$objReturn.Message		= "Success" or the error message.
			#$objReturn.Returns		= The composed/sent email object.
		#$EmailTo = Who to send the email to.  (i.e. "andrew.k.freeman@nmci-isf.com")
		#$EmailFrom = Email/Mailbox to send from.  (i.e. "henry.schade@nmci-isf.com")
		#$EmailSubject = What to put in the email Subject.
		#$EmailBody = What to put in the email Body.
		#$EmailCC = Who to CC on the email.  (i.e. "andrew.k.freeman@nmci-isf.com")
		#$EmailBCC = Who to BCC on the email.  (i.e. "andrew.k.freeman@nmci-isf.com")
		#$EmailAttach = Full path to file to attach.
		#$AsHTML = $True or $False.  Compose the body as HTML?
		#$bolTryAuto = $True or $False.  Use the .AutodiscoverUrl property to try and find the Exchange URL for the ExchangeService object.
		#$strExchVer = The version of Exchange that is being connected to. ("Exchange2007_SP1", or "Exchange2010", or "Exchange2010_SP1", or "Exchange2010_SP2", or "Exchange2013")

		#Setup the PSObject to return.
		#http://stackoverflow.com/questions/21559724/getting-all-named-parameters-from-powershell-including-empty-and-set-ones
		$CommandName = $PSCmdlet.MyInvocation.InvocationName;
		$ParameterList = (Get-Command -Name $CommandName).Parameters;
		$strTemp = "";
		foreach ($key in $ParameterList.keys){
			$var = Get-Variable -Name $key -ErrorAction SilentlyContinue;
			if($var){$strTemp += "[$($var.name) = $($var.value)] ";}
		}
		$strTemp = $CommandName + "(" + $strTemp.Trim() + ")";
		$objReturn = New-Object PSObject -Property @{
			Name = $strTemp
			Results = $False
			Message = "Error"
			Returns = "";
		}

		#Check that EWS is installed
		$objResults = EWSVerifyInstall;
		if ($objResults.Results -eq $True){
			#Load the EWS Managed API Assembly
			$dllFile = "Microsoft.Exchange.WebServices.dll";
			$dllPath = $objResults.Returns + $dllFile;
			[void][Reflection.Assembly]::LoadFile($dllPath);

			#Insatiate the EWS service object 
			$objExchServ = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList $strExchVer;
			#$objExchServ.TraceEnabled = $bolTrace;
			#if (($EmailFrom.Contains(($env:username + "@"))) -or ($env:username -eq "") -or ($env:username -eq $null)){
			if (($EmailFrom.Contains(($env:username + "@")))){
				$objExchServ.UseDefaultCredentials = $True;
			}else{
				#Connect using impersonation / alternate credentials.
				#$Credential = Get-Credential -Credential "SelectAnAccount";
				$Credential = $host.ui.PromptForCredential("Need Outlook Credentials", "Please select/enter your Outlook information below.", "", "NetBiosUserName");
				#Set the credentials for Exchange
				$objExchServ.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $Credential.UserName, $Credential.GetNetworkCredential().Password;
				$objExchServ.UseDefaultCredentials = $False;
			}
			$Error.Clear();
			#Determine/Set the EWS endpoint
			if ($bolTryAuto){
				if ($EmailFrom.Contains(($env:username + "@"))){
					$strRet = $objExchServ.AutodiscoverUrl($EmailFrom, {$True});			#This is taking FOREVER.  Looks like it is setting the URL of the ExchangeService object, which is required for subscriptions.
				}else{
					$strRet = $objExchServ.AutodiscoverUrl($Credential.UserName, {$True});
				}
			}
			if (($Error) -or ($bolTryAuto -ne $True)){
				$Error.Clear();
				$strURL = [system.URI] "https://cas-sdni.nadsuswe.nads.navy.mil/ews/exchange.asmx";
				$objExchServ.Url = $strURL;
			}

			#Create the email message and set the Subject and Body
			$Error.Clear();
			$objMessage = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $objExchServ;
			$objMessage.Subject = $EmailSubject;
			$objMessage.Body = $EmailBody;
			#If the $AsHTML parameter is not used, send the message as plain text
			if(!$AsHTML) {
				$objMessage.Body.BodyType = 'Text';
			}

			#Add each specified recipient.
			if (($EmailTo.IndexOf(",") -gt 0) -or ($EmailTo.IndexOf(";") -gt 0)){
				if ($EmailTo.IndexOf(",") -gt 0){
					$arrAdds = $EmailTo.Split(",");
				}else{
					$arrAdds = $EmailTo.Split(";");
				}
				for ($intX = 0; $intX -lt $arrAdds.Count; $intX++){
					$arrAdds[$intX] = $arrAdds[$intX].Trim();
					if (($arrAdds[$intX] -ne "") -and ($arrAdds[$intX] -ne $null)){
						$strRet = $objMessage.ToRecipients.Add($arrAdds[$intX]);
					}
				}
			}else{
				$strRet = $objMessage.ToRecipients.Add($EmailTo);
			}

			#Add each CC recipient.
			if (($EmailCC -ne "") -and ($EmailCC -ne $null)){
				if (($EmailCC.IndexOf(",") -gt 0) -or ($EmailCC.IndexOf(";") -gt 0)){
					if ($EmailCC.IndexOf(",") -gt 0){
						$arrAdds = $EmailCC.Split(",");
					}else{
						$arrAdds = $EmailCC.Split(";");
					}
					for ($intX = 0; $intX -lt $arrAdds.Count; $intX++){
						$arrAdds[$intX] = $arrAdds[$intX].Trim();
						if (($arrAdds[$intX] -ne "") -and ($arrAdds[$intX] -ne $null)){
							$strRet = $objMessage.CcRecipients.Add($arrAdds[$intX]);
						}
					}
				}else{
					$strRet = $objMessage.CcRecipients.Add($EmailCC);
				}
			}

			#Add each BCC recipient.
			if (($EmailBCC -ne "") -and ($EmailBCC -ne $null)){
				if (($EmailBCC.IndexOf(",") -gt 0) -or ($EmailBCC.IndexOf(";") -gt 0)){
					if ($EmailBCC.IndexOf(",") -gt 0){
						$arrAdds = $EmailBCC.Split(",");
					}else{
						$arrAdds = $EmailBCC.Split(";");
					}
					for ($intX = 0; $intX -lt $arrAdds.Count; $intX++){
						$arrAdds[$intX] = $arrAdds[$intX].Trim();
						if (($arrAdds[$intX] -ne "") -and ($arrAdds[$intX] -ne $null)){
							$strRet = $objMessage.BccRecipients.Add($arrAdds[$intX]);
						}
					}
				}else{
					$strRet = $objMessage.BccRecipients.Add($EmailBCC);
				}
			}

			#https://msdn.microsoft.com/en-us/library/office/hh532564(v=exchg.80).aspx
			if (($EmailAttach -ne "") -and ($EmailAttach -ne $null)){
				$strRet = $objMessage.Attachments.AddFileAttachment($EmailAttach.Split("\")[-1], $EmailAttach);
				#$strRet = $objMessage.Attachments.AddFileAttachment($EmailAttach, $EmailAttach);		#This works just fine too.
				#$objMessage.Attachments[0].IsInline = $True;
				#$objMessage.Attachments(0).ContentId = $EmailAttach.Split("\")[-1];
			}

			if ($Error){
			}else{
				if ($EmailFrom.Contains(($env:username + "@"))){
					#Send the message and save a copy in the Sent Items folder.
					$strRet = $objMessage.SendAndSaveCopy();
					#Send the message and DO NOT save a copy in the Sent Items folder.  Can NOT be used by an account without a mailbox.
					#$strRet = $objMessage.Send();
				}else{
					$strRet = $objMessage.SendAndSaveCopy("SentItems");
				}
			}

			if ($Error){
				$objReturn.Results = $False;
				$objReturn.Message = "Error composing and sending message. `r`n" + $Error;
			}else{
				$objReturn.Results = $True;
				$objReturn.Message = "Success";
				$objReturn.Returns = $objMessage;
			}
		}else{
			$objReturn.Results = $False;
			$objReturn.Message = "Error, EWS dll is not installed.";
		}

		return $objReturn;
	}

	function SendEmailPS{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$EmailTo, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$EmailFrom, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$EmailSubject, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$EmailBody, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$SMTPServer
		)
		#This should work, but I am having issues with it.  Use SendEmailEWS(), as it does work.

		#Send an email from PowerShell, via SMTP:
		#https://4sysops.com/archives/send-email-with-powershell/
		#http://blogs.msdn.com/b/rkramesh/archive/2012/03/16/sending-email-using-powershell-script.aspx
		#http://www.philerb.com/2011/11/sending-mail-with-powershell/

		if (($SMTPServer -eq "") -or ($SMTPServer -eq $null)){
			$SMTPServer = "NAWESDNIXM05V.nmci-isf.com";
		}

		#PS ver 2.0+
		#$PSEmailServer = "";
		#Send-MailMessage -to "henry.schade@nmci-isf.com" -from "PowerShell <power.shell@domain.com>" -Subject "Test" -body "Test for Send-MailMessage";
		#Send-MailMessage -To $strTo -From $strFrom -Subject $strSub -Body $strBody -SmtpServer $strServer -Credential "SelectAnAccount";

		Send-MailMessage -To $EmailTo -From $EmailFrom -Subject $EmailSubject -Body $EmailBody -SmtpServer $SMTPServer;

	}

	function SetupConn{
		Param(
			[Parameter(Mandatory=$False)][String]$WhatSide, 
			[Parameter(Mandatory=$False)][String]$strServer
		);
		#Returns a list of all the PowerShell commandlets imported.
		#$WhatSide = $args[0];
		#$strServer = "Default", "Random", "naeaNRFKxh01v"

		$Session = "";
		#Clear-Host;

		CleanUpConn;

		if (($WhatSide -eq $null) -or ($WhatSide -eq "")){
			$WhatSide = Read-Host 'What Domain? (nadsus[W]e or nadsus[E]a or [P]acom)';
		}
		if ($WhatSide.Length -gt 1){
			if (($WhatSide -eq "nadsusea") -or ($WhatSide -eq "nadsuswe") -or ($WhatSide -eq "pads")){
				if ($WhatSide -eq "nadsusea"){
					$WhatSide = "e";
				}
				if ($WhatSide -eq "nadsuswe"){
					$WhatSide = "w";
				}
				if ($WhatSide -eq "pads"){
					$WhatSide = "p";
				}
			}
			else{
				$WhatSide.substring(0, 1)
			}
		}
		if (($WhatSide -ne "e") -and ($WhatSide -ne "w") -and ($WhatSide -ne "p")){
			$WhatSide = Read-Host 'What Domain? (nadsus[W]e or nadsus[E]a or [P]acom)';
		}
		if ($WhatSide.Length -gt 1){
			$WhatSide.substring(0, 1)
		}

		if (($strServer -eq $null) -or ($strServer -eq "")){
			$strServer = Read-Host 'What Server? ([D]efault, [R]andom, or Exch Svr Name [i.e. naeaNRFKxh01v])';
			if (($strServer -eq "D") -or ($strServer -eq "Default") -or ($strServer -eq "") -or ($strServer -eq $null) -or ($strServer.Length -lt 10)){
				$strServer = "Default"
			}
			if (($strServer -eq "R") -or ($strServer -eq "Random")){
				$strServer = "Random"
			}
		}
		else{
			if ($strServer.Length -eq 1){
				switch ($strServer){
					"R"{
						$strServer = "Random";
					}
					"D"{
						$strServer = "Default";
					}
					default{
						$strServer = "Default";
					}
				}
			}
		}

		$InitializeDefaultDrives=$False;
		if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};

		$strRIDMaster = "";

		if ($WhatSide -eq "e"){
			#Write-Host "East it is";
			$strDomain = "nadsusea";
			if (($strServer -eq "Default") -or ($strServer -eq "D")){
				$strServer = "naeaNRFKxh01v.nadsusea.nads.navy.mil";
				#Test-Connection -CN $strComputer -buffersize 16 -Count 1 -ErrorAction 0 -quiet
				#if ((Test-Connection -CN $strFQDN -buffersize 16 -Count 1 -ErrorAction 0 -quiet) -ne $True){
					#Specify a new Server to connect to.
				#}
			}
			else{
				if (($strServer -eq "Random") -or ($strServer -eq "R")){
					#Get all the Exchange Servers in the Domain (filter the results).
					$objExchServers = GetExchangeServers | where {(($_.FQDN -match "nadsusea") -and (($_.Roles -match 4) -or ($_.Roles -match 36)))};

					#select a random server from the list
					if($objExchServers.Count -gt 1) {
						$intRandom = Get-Random -Minimum 0 -Maximum $objExchServers.Count;
						$strServer = $objExchServers[$intRandom].FQDN;
					}else{
						$strServer = $objExchServers[0].FQDN;
					}
				}
			}
			$strRIDMaster = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $strDomain))).RidRoleOwner.Name;
		}
		if ($WhatSide -eq "w"){
			#Write-Host "West it is";
			$strDomain = "nadsuswe";
			if (($strServer -eq "Default") -or ($strServer -eq "D")){
				$strServer = "naweSDNIxh01v.nadsuswe.nads.navy.mil";
				#Test-Connection -CN $strComputer -buffersize 16 -Count 1 -ErrorAction 0 -quiet
			}
			else{
				if (($strServer -eq "Random") -or ($strServer -eq "R")){
					#Get all the Exchange Servers in the Domain (filter the results).
					$objExchServers = GetExchangeServers | where {(($_.FQDN -match "nadsuswe") -and (($_.Roles -match 4) -or ($_.Roles -match 36)))};

					#select a random server from the list
					if($objExchServers.Count -gt 1) {
						$intRandom = Get-Random -Minimum 0 -Maximum $objExchServers.Count;
						$strServer = $objExchServers[$intRandom].FQDN;
					}else{
						$strServer = $objExchServers[0].FQDN;
					}
				}
			}
			$strRIDMaster = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $strDomain))).RidRoleOwner.Name;
		}
		if ($WhatSide -eq "p"){
			#Write-Host "Pacom it is";
			$strDomain = "pads";
			if (($strServer -eq "Default") -or ($strServer -eq "D")){
				$strServer = "PADSPRLHXF01V.pads.pacom.mil";
				#Test-Connection -CN $strComputer -buffersize 16 -Count 1 -ErrorAction 0 -quiet
			}
			else{
				if (($strServer -eq "Random") -or ($strServer -eq "R")){
					#Get all the Exchange Servers in the Domain (filter the results).
					$objExchServers = GetExchangeServers | where {(($_.FQDN -match "pacom") -and (($_.Roles -match 4) -or ($_.Roles -match 36)))};

					#select a random server from the list
					if($objExchServers.Count -gt 1) {
						$intRandom = Get-Random -Minimum 0 -Maximum $objExchServers.Count;
						$strServer = $objExchServers[$intRandom].FQDN;
					}else{
						$strServer = $objExchServers[0].FQDN;
					}
				}
			}
			$strRIDMaster = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $strDomain))).RidRoleOwner.Name;
		}
		$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$strServer/PowerShell/" -Authentication Kerberos;

		if (($Session -ne "") -and ($Session -ne $null)){
			$strModule = (Import-PSSession $Session -AllowClobber);
			#$strModule.Name will then be the ModuleName to look for.  i.e. "tmp_2b3babb5-acb0-405e-ac5e-f4bd3cd3c07d_z1kcca4b.hdr"
			#But $strModule.ExportedFunctions has all the functions/commandlets imported.
			if (($strModule -ne "") -and ($strModule -ne $null)){
				# http://www.get-exchangeserver.com/powershell-script-ile-exchange-server-20132010-uzerinde-health-check-report-olusturma/
				#Set-ADServerSettings -ViewEntireForest $true -WarningAction SilentlyContinue;
				if (!(Get-ADServerSettings).ViewEntireForest){
					Set-ADServerSettings -ViewEntireForest $true;
				}

				#if ((Get-ADServerSettings).UserPreferredDomainControllers -eq ""){
					$Error.Clear();
					if (($strRIDMaster -ne "") -and ($strRIDMaster -ne $null)){
						Set-ADServerSettings -PreferredServer $strRIDMaster -ErrorAction SilentlyContinue;
					}
					if ($Error){
						Write-Host "";
						Write-Host "An error occurred trying to run this command:";
						Write-Host "";
						Write-Host "Set-ADServerSettings -PreferredServer $strRIDMaster;";
						Write-Host "";
					}
				#}
				return $strModule;
			}
			else{
				Write-Host "";
				Write-Host "An error occurred trying to import the Exchange PowerShell CommandLets.";
				Write-Host "";
				Write-Host "Import-PSSession $Session -AllowClobber;";
				Write-Host "";
			}
		}
	}


	#From PS-ExchConn.ps1.
	if ($args[0] -eq "Setup"){
		#Write-Host "Args is: " $args[0];
		if ($args[1] -eq "Ret"){
			#To use this feature here the the powershell command to run:
			#$strReturn = & "C:\SRM_Apps_N_Tools\PS-Scripts\Exchange.ps1" "Setup" "Ret"
				#$strReturn will then have a string of all the Exchange PowerShell command that were imported w/ the session.

			$arrRet = SetupConn;

			$arrPSCmds = $arrRet.ExportedCommands;
			if ($arrPSCmds.GetType() -ne "ArrayList"){
				[System.Collections.ArrayList]$arrPSCmds = $arrPSCmds;
			}

			if ($arrPSCmds.GetType() -eq "ArrayList"){
				for ($intX = $arrPSCmds.Count; $intX -ge 0; $intX--){
					if ($arrPSCmds[$intX].Name -ne ""){
						$arrPSCmds.Remove($arrPSCmds[$intX]);
					}
				}
			}

			$arrPSCmds = $arrPSCmds | Sort-Object -Property Name;
			$arrPSCmds = $arrPSCmds | Select Name;

			$strPSCmds = "";
			foreach ($strCmd in $arrPSCmds){
				$strPSCmds = $strPSCmds + ", " + $strCmd;
			}
			$strPSCmds = $strPSCmds.SubString(1).Trim();
			$strPSCmds = $strPSCmds.Replace("@{Name=", "");
			$strPSCmds = $strPSCmds.Replace("}", "");

			return $strPSCmds;
		}
		else{
			if (($args[1] -ne "") -and ($args[1] -ne $null) -and ($args[2] -ne "") -and ($args[2] -ne $null)){
				SetupConn $args[1] $args[2];
			}
			else{
				SetupConn;
			}
		}
	}
	if ($args[0] -eq "CleanUp"){
		#Write-Host "Args is: " $args[0];
		CleanUpConn;
	}
