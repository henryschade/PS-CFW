###########################################
# Updated Date:	17 July 2015
# Purpose:		Routines that require a Computer, or that interact w/ a Computer.
# Requirements: None
##########################################

	function CheckIfOnline($strComp){
		if(($strComp -eq "") -OR ($strComp -eq $null)){
			#$strComp = "ALSDCP002656";		#Henry Laptop;
			$strComp = Read-Host 'What computer? (i.e. ALSDCP002656)';
		}

		$ErrorActionPreference = 'SilentlyContinue';
		$strIP = [System.Net.DNS]::GetHostAddresses($strComp);
		$ErrorActionPreference = 'Continue';
		#Write-Host $strIP;

		if ($strIP -eq $null){ 
			$Message = "Host cannot be resolved in DNS.";
		}else{
			#host is valid; now check if it is online by pinging it
			$ping = New-Object System.Net.NetworkInformation.Ping;
			$Reply = $ping.send($strComp);

			if ($Reply.status -eq "Success"){
				#Host is online
				$Message = "Host is online.";

				$Message = LoggedInUser($strComp);

				#ShutDown a computer.
				#$Message = Stop-Computer -comp $strComp -force;
			}else{
				$Message = "Host not online (ping failed).";
			}
		}

		#displays the result
		#$a = new-object -comobject wscript.shell;
		#$b = $a.popup($Message,0,"Logged In User",1);
		#Write-Host $Message;

		return $Message;
	}

	function CreateSchedTask{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$ScriptDir
		)
		#$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;

		#Get Scheduled Tasks
		##$strTest = C:\..path..\GetScheduledTask.ps1 -Hidden -SubFolders
		##$strTest = C:\..path..\GetScheduledTask.ps1 "*load*" -Hidden -SubFolders
		#$strPowerSave = . ($ScriptDir + "\Get-ScheduledTask.ps1") | where { $_.TaskName -match "PowerSave" };
		##Write-Host $strPowerSave.Count;
		#$strMonitor = . ($ScriptDir + "\Get-ScheduledTask.ps1") | where { $_.TaskName -match "Monitor" };
		##Write-Host $strMonitor.Count;
		#$strMonitor = . ($ScriptDir + "\Get-ScheduledTask.ps1") -Subfolders -Hidden | where { $_.TaskName -match "Monitor" };

		#$sMachCertTasks = . ($ScriptDir + "\Get-ScheduledTask.ps1") -Subfolders -Hidden | where { $_.TaskName -match "\\Microsoft\\Windows\\CertificateServicesClient\\*" };
		$sMachCertTasks = . ($ScriptDir + "\Get-ScheduledTask.ps1") -Subfolders -Hidden -TaskName "\Microsoft\Windows\CertificateServicesClient\*";



		#Get Local Machine Cert
		#http://mcpmag.com/articles/2014/11/04/expiring-certs-in-powershell.aspx
		#http://blogs.msdn.com/b/sonam_rastogi_blogs/archive/2014/08/18/request-export-and-import-certificate-using-powershell.aspx

		#https://social.technet.microsoft.com/Forums/scriptcenter/en-US/76f02650-e1fd-42fc-963e-c4ad86eeb65c/export-import-local-machine-certificate-using-powershell?forum=ITCG
		#https://gallery.technet.microsoft.com/scriptcenter/Certification-File-Manager-be4a6848
		#Get-ChildItem -Recurse "Cert:\LocalMachine\My" | FL
		# next requires PowerShell 4.0, 
		#Export-Certificate -Type CERT -FilePath C:\OrchCert.cer -Cert "Cert:\LocalMachine\My"
		#invoke-item Cert:\LocalMachine\My

		#Create a Zip file
		#Get-Childitem c:\fso -Recurse | Write-Zip -IncludeEmptyDirectories -OutputPath C:\fso_bu\fso.zip
		#That probably won't work, so here (requires .NET 4.5 and Powershell 3+):
		#http://www.adminarsenal.com/admin-arsenal-blog/powershell-zip-up-files-using-.net-and-add-type
		#This next one looks EASY.
		#http://blogs.msdn.com/b/jerrydixon/archive/2014/08/08/zipping-a-single-file-with-powershell.aspx




		<#
		#Create a Scheduled Task
		#Can use "schtasks.exe" to do the work.
		if ($strMonitor.Count -ne 1){
			#This is the block to run, to schedule the PowerShell Monitor.
			$TaskName = "PowerShell Monitor";
			$TaskRun = "PowerShell.exe -ExecutionPolicy ByPass -Command ""C:\Windows\ps-scripts\PS-Monitor.ps1""";
			#$TaskRun = "PowerShell.exe -ExecutionPolicy ByPass -Command ""$ScriptDir\PS-Monitor.ps1""";
			$TaskRunAs = "SYSTEM";			# /RP not used if /RU is the SYSTEM account.
			$TaskSched = "ONSTART";
			schtasks /Create /RU $TaskRunAs /TN $TaskName /TR $TaskRun /SC $TaskSched /F;
		}

		#This is the block to run, to schedule the PowerSave Tasks.
		$arrTimes = @("23", "23:30", "00", "01", "02", "03", "04", "05", "05:30");
		if ($strPowerSave.Count -lt $arrTimes.Count){
			foreach ($strTime in $arrTimes){
				#$TaskName = "PowerSave01";
				$TaskName = "PowerSave" + $strTime -Replace "\:", "";					#This uses Regular expressions;
				$TaskRun = "C:\Windows\System32\shutdown.exe /s /f /t 0";
				$TaskRunAs = "SYSTEM";			# /RP not used if /RU is the SYSTEM account.
				$TaskSched = "DAILY";
				#$TaskStart = "23:00";
				$TaskStart = $strTime;
				if (!$TaskStart.Contains(":")){
					$TaskStart = $TaskStart + ":00";
				}
				schtasks /Create /RU $TaskRunAs /TN $TaskName /TR $TaskRun /SC $TaskSched /ST $TaskStart /F;
			}
		}
		#>
	}

	function DoShutDown($strComp){
		if(($strComp -eq "") -OR ($strComp -eq $null)){
			#$strComp = "ALSDCP002656";		#Henry Laptop;
			$strComp = Read-Host 'What computer? (i.e. ALSDCP002656)';
		}

		#$Message = Stop-Computer -comp $strComp -force;
		$Message = Stop-Computer -comp $strComp -force -Credential $creds;
	}

	function GetCert{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$FilePath = "C:\SRM_Apps_N_Tools\", 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$FileName = ($Env:COMPUTERNAME + "_" + $Env:USERNAME)
		)
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters passed in.
			#$objReturn.Results		= $True or $False.  Was the Cert exported.
			#$objReturn.Message		= "Success" or the error message.
			#$objReturn.Returns		= The full path annd file name of the file created.

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

		if ($FileName.EndsWith(".adm")){
			$FileName = $FileName.SubString(0, ($FileName.Length - 4));
		}
		if ($FileName.EndsWith(".ad")){
			$FileName = $FileName.SubString(0, ($FileName.Length - 3));
		}

		$FileName = $FileName + ".cer";
		if ((Test-Path -Path ($FilePath + $FileName))){
			#File already exists.
			#Get file, and then add the 8 digit date code to the beginning of the file name.
			#Copy the original file renaming it.
			Copy-Item -Path ($FilePath + $FileName) -Destination ($FilePath + (Get-Date).ToString("yyyyMMdd_") + $FileName);
		}

		#https://jorgequestforknowledge.wordpress.com/2012/02/10/managing-certificates-on-a-windows-computer-with-powershell/
		# Find the Cert Based Upon Some Condition
		#$CertToExport = dir cert:\LocalMachine\My | where {$_.ThumbPrint -eq "EC9498B48CA4E48EB8D5BC557BCFBC09B5A02651"};
		$CertsToExport = dir cert:\LocalMachine\My;
		if (@($CertsToExport).Count -eq 1){
			if ((([System.DateTime]::Now - $CertsToExport.NotAfter).Days) -ge 0){
				#Write-Host "The Machine Cert has not been renewed yet.";
				$objReturn.Message = "Error.  The Machine Cert has not been renewed yet.  It expires '" + $CertsToExport.NotAfter.ToString() + "'.";
			}else{
				# Export The Targeted Cert In Bytes For The CER format
				$CertFileToExport = $CertsToExport.Export("Cert");

				# Write The Files Based Upon The Exported Bytes
				[system.IO.file]::WriteAllBytes(($FilePath + $FileName), $CertFileToExport);

				$objReturn.Results = $True;
				$objReturn.Message = "Success";
				$objReturn.Returns = ($FilePath + $FileName);
			}
		}else{
			#Write-Host "There is more than one Machine Cert, and don't know which one to export.";
			$objReturn.Message = "Error.  There is more than one Machine Cert, and don't know which one to export.";
			#$objReturn.Returns = "";
		}

		return $objReturn;
	}

	function GetCredentials($strComp){
		if (($strComp -eq "") -or ($strComp -eq $null)){
			$strComp = $Env:ComputerName;
		}

		$strUser = $Env:UserName;

		#$creds = Get-Credential $strComp\admin;				#domain\user
		$creds = Get-Credential ($strComp + "\" + $strUser);				#domain\user
		#$creds = Get-Credential -Credential $strComp\hyyjyg;
		$user = $creds.UserName;
		$password = $creds.GetNetworkCredential().Password;
		#Write-Host $user " -- " $password

		<#
		if($creds -ne $null){
			#$comp = Get-WmiObject Win32_ComputerSystem -ComputerName $strComp -Credential $creds -Authentication 6;
			#$comp = Get-WmiObject -Namespace "root\cimv2" Win32_ComputerSystem -ComputerName $strComp -Credential $creds -Authentication 6 -Impersonation Delegate;
			#$comp = Get-WmiObject -Namespace "root\cimv2" Win32_ComputerSystem -ComputerName $strComp -Credential $strComp\hyyjyg -Impersonation Impersonate;
			$comp = Get-WmiObject -Namespace "root\cimv2" Win32_ComputerSystem -ComputerName $strComp -Credential ($strComp + "\" + $strUser) -Impersonation 3;
		}else{
			#$comp = Get-WmiObject Win32_ComputerSystem -ComputerName $strComp -Authentication PacketPrivacy;
			#$comp = Get-WmiObject Win32_ComputerSystem -ComputerName $strComp -Authentication 6;
			#$comp = Get-WmiObject Win32_ComputerSystem -ComputerName $strComp -Authentication PacketPrivacy -Impersonation Delegate;
			$comp = Get-WmiObject Win32_ComputerSystem -ComputerName $strComp -Authentication PacketPrivacy -Impersonation Impersonate;
		}
		Write-Host "~~" $comp;
		#>
	}

	function LoggedInUser($strComp){
		if(($strComp -eq "") -OR ($strComp -eq $null)){
			#$strComp = "ALSDCP002656";		#Henry Laptop;
			$strComp = Read-Host 'What computer? (i.e. ALSDCP002656)';
		}

		#check if a user is logged in
		$strLoggedIn = gwmi -class win32_computerSystem -computer:$strComp | Select-Object username;
		if ($strLoggedIn.username.Length -gt 0){
			$Message = $strLoggedIn.username + " is currently logged in to " + $strComp + ".";
		}else{
			$Message = "No user is logged in locally to " + $strComp + ".";
		}

		#Write-Host $Message;

		return $Message;
	}

