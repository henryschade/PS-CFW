###########################################
# Updated Date:	5 May 2016
# Purpose:		Routines that require a Computer, or that interact w/ a Computer.
# Requirements: None
##########################################

	function CheckIfOnline{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strComp, 
			[ValidateNotNull()][Parameter(Mandatory=$False)]$bolUsers = $True
		)
		#$strComp = The computer to check.  ($Env:ComputerName)

		#if ([String]::IsNullOrEmpty($strComp)){
		#	#$strComp = "ALSDCP002656";		#Henry Laptop;
		#	#$strComp = "ALSDNI390014";		#Andrew Laptop;
		#	$strComp = Read-Host 'What computer? (i.e. ALSDCP002656)';
		#}

		$ErrorActionPreference = 'SilentlyContinue';
		$strIP = [System.Net.DNS]::GetHostAddresses($strComp);
		$ErrorActionPreference = 'Continue';
		#Write-Host $strIP;

		if ([String]::IsNullOrEmpty($strIP)){
			$strRet = "Host ($strComp) cannot be resolved, by DNS.";
		}else{
			#host is valid; now check if it is online by pinging it
			$objPing = New-Object System.Net.NetworkInformation.Ping;
			$Reply = $objPing.Send($strComp);

			if ($Reply.Status -eq "Success"){
				#Host is online
				$strRet = "Host ($strComp) is online.";

				if ($bolUsers -eq $True){
					$strRet = $strRet + "`r`n" + (LoggedInUser $strComp);
				}

				#ShutDown a computer.
				#$strRet = Stop-Computer -comp $strComp -force;
			}else{
				$strRet = "Host ($strComp) not online (ping failed).";
			}
		}

		#displays the results, in a popup.
		#$a = New-Object -comobject wscript.shell;
		#$b = $a.Popup($strRet,0,"Logged In User",1);

		return $strRet;
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

	function DoShutDown{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strComp, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][bool]$bReboot = $False, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][bool]$bAskCreds = $False
		)
		#$strComp = The computer to shutdown.
		#$bReboot = Should we Reboot instead of shutdown?
		#$bAskCreds = Prompt for credentials?

		if ([String]::IsNullOrEmpty($strComp)){
			#$strComp = "ALSDCP002656";		#Henry Laptop;
			$strComp = Read-Host 'What computer? (i.e. ALSDCP002656)';
		}

		if ($bAskCreds -eq $True){
			$objCreds = GetCredentials;
		}

		if ($bReboot -eq $True){
			if ($bAskCreds -eq $True){
				$strRet = Restart-Computer -comp $strComp -Force -Credential $objCreds;
			}
			else{
				$strRet = Restart-Computer -comp $strComp -Force;
			}
		}
		else{
			if ($bAskCreds -eq $True){
				$strRet = Stop-Computer -comp $strComp -Force -Credential $objCreds;
			}
			else{
				$strRet = Stop-Computer -comp $strComp -Force;
			}
		}
		
		return $strRet;
	}

	function EnableRPC{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strComp
		)
		#$strComp = The computer to Enable RPC on.  ($Env:ComputerName)

		#Chris has a way.
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
		if([String]::IsNullOrEmpty($strComp)){
			$strComp = $Env:ComputerName;
		}

		$strUser = $Env:UserName;

		#$objCreds = Get-Credential $strComp\admin;				#domain\user
		$objCreds = Get-Credential ($strComp + "\" + $strUser);				#domain\user
		#$objCreds = Get-Credential -Credential $strComp\hyyjyg;
		#$user = $objCreds.UserName;
		#$password = $objCreds.GetNetworkCredential().Password;
		##Write-Host $user " -- " $password

		<#
		if([String]::IsNullOrEmpty($objCreds)){
			#$comp = Get-WmiObject Win32_ComputerSystem -ComputerName $strComp -Authentication PacketPrivacy;
			#$comp = Get-WmiObject Win32_ComputerSystem -ComputerName $strComp -Authentication 6;
			#$comp = Get-WmiObject Win32_ComputerSystem -ComputerName $strComp -Authentication PacketPrivacy -Impersonation Delegate;
			$comp = Get-WmiObject Win32_ComputerSystem -ComputerName $strComp -Authentication PacketPrivacy -Impersonation Impersonate;
		}else{
			#$comp = Get-WmiObject Win32_ComputerSystem -ComputerName $strComp -Credential $objCreds -Authentication 6;
			#$comp = Get-WmiObject -Namespace "root\cimv2" Win32_ComputerSystem -ComputerName $strComp -Credential $objCreds -Authentication 6 -Impersonation Delegate;
			#$comp = Get-WmiObject -Namespace "root\cimv2" Win32_ComputerSystem -ComputerName $strComp -Credential $strComp\hyyjyg -Impersonation Impersonate;
			$comp = Get-WmiObject -Namespace "root\cimv2" Win32_ComputerSystem -ComputerName $strComp -Credential ($strComp + "\" + $strUser) -Impersonation 3;
		}
		Write-Host "~~" $comp;
		#>

		return $objCreds;
	}

	function GetRegistry{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strRegPath, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strProp, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strRemoteSys = "."
		)
		#Returns an array.
			#A "list" of Properties & "Directories" (Keys) under the "Path" provided.
			#Data from the Property in the format of: (Name, Type/Kind, Value).
		#$strRegPath = The Registry Path (Key) to get data from.  (i.e. "HKEY_LOCAL_MACHINE:\SOFTWARE\NMCI\ITSS-Tools\ASCII")
			#HKEY_CLASSES_ROOT
			#HKEY_CURRENT_CONFIG
			#HKEY_CURRENT_USER
			#HKEY_DYN_DATA
			#HKEY_LOCAL_MACHINE
			#HKEY_PERFORMANCE_DATA
			#HKEY_USERS
		#$strProp = The Property, if any, to get the data out of.  If omitted returns a list of available Properties & Keys.  (i.e. "(Default)" or "Version")
		#$strRemoteSys = A remote system to get the registry of.  "ALSDCP002656" ("." specifies local system, the default)

		#registry key = The Reg Path.
		#registry key property = The "attributes" / "fields" under the Key.

		#Remote Registry:
		#https://psremoteregistry.codeplex.com/		->	(http://blogs.microsoft.co.il/scriptfanatic/2010/01/10/remote-registry-powershell-module/)
		#http://stackoverflow.com/questions/1133335/openremotebasekey-credentials
		#https://social.technet.microsoft.com/Forums/windowsserver/en-US/14f33784-09a0-49be-8036-73921181fa3c/microsoftwin32registrykeyopenremotebasekey?forum=winserverpowershell

		$arrRet = $null;
		if (!([String]::IsNullOrEmpty($strRemoteSys))){
			#Write-Host "    Starting remote registry connection against: '$strRemoteSys'.";
			$strRoot = $strRegPath.SubString(0,$strRegPath.IndexOf(":"));
			switch ($strRoot){
				"HKEY_CLASSES_ROOT "{
					$strRoot = [Microsoft.Win32.RegistryHive]::ClassesRoot;
				}
				"HKEY_CURRENT_CONFIG"{
					$strRoot = [Microsoft.Win32.RegistryHive]::CurrentConfig;
				}
				"HKEY_CURRENT_USER"{
					$strRoot = [Microsoft.Win32.RegistryHive]::CurrentUser;
				}
				"HKEY_DYN_DATA"{
					$strRoot = [Microsoft.Win32.RegistryHive]::DynData;
				}
				"HKEY_LOCAL_MACHINE"{
					$strRoot = [Microsoft.Win32.RegistryHive]::LocalMachine;
				}
				"HKEY_PERFORMANCE_DATA"{
					$strRoot = [Microsoft.Win32.RegistryHive]::PerformanceData;
				}
				"HKEY_USERS"{
					$strRoot = [Microsoft.Win32.RegistryHive]::Users;
				}
				default{
					$arrRet = @("Error connecting to the (unknown) Registry root '$strRoot' of '$strRemoteSys'.", "-1");
					return $arrRet;
				}
			}
			#Write-Host "    Registry Hive is: [$strRoot]. `r`n";
			$strKey = $strRegPath.SubString($strRegPath.IndexOf(":") + 2);

			$Error.Clear();
			#Connect to "Root" of the Registry.
			$objReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($strRoot, $strRemoteSys);
			if ($Error){
				$strRoot = $strRegPath.SubString(0,$strRegPath.IndexOf(":"));
				$arrRet = @("Error connecting to the Registry '$strRoot' of '$strRemoteSys'.", $Error);
			}
			else{
				#Write-Host "    Open remote subkey: [$strKey].";
				$Error.Clear();
				$subKey = $objReg.OpenSubKey($strKey);
				#$subKey | GM;
				if (($Error) -or ([String]::IsNullOrEmpty($subKey))){
					$arrRet = @("The Key '$strRegPath' does NOT exist.", $Error);
				}
				else{
					if ([String]::IsNullOrEmpty($strProp)){
						[System.Collections.ArrayList]$arrRet = @();
						foreach ($objVal in $subKey.GetValueNames()){
							#Write-Host $objVal;
							if ([String]::IsNullOrEmpty($objVal)){
								$arrRet += "(Default)";
							}
							else{
								$arrRet += $objVal;
							}
						}
						foreach ($objKey in $subKey.GetSubKeyNames()){
							#Write-Host $objKey;
							$arrRet += "\" + $objKey;
						}
					}
					else{
						#Get prop info
						if ($strProp -eq "(Default)"){
							$strProp = "";
						}
						$Error.Clear();
						$ErrorActionPreference = 'SilentlyContinue';
						$arrRet = @($strProp, $subKey.GetValueKind($strProp), $subKey.GetValue($strProp));
						$ErrorActionPreference = 'Continue';
						if ($Error){
							$arrRet = @("The Property '$strProp' does NOT exist under Key '$strRegPath'.", $Error);
						}
					}
				}
			}
			#Write-Host "    Closing remote registry connection on: '$strRemoteSys'.";
			$objReg.close();
		}

		return $arrRet;
	}

	function GetRunningApp{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strAppName
		)
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters passed in.
			#$objReturn.Results		= # of instances of $strAppName found.
			#$objReturn.Message		= "Success" or the error message.
			#$objReturn.Returns		= A collection of $strAppName instances.

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
			Results = 0
			Message = "Error"
			Returns = "";
		}

		$objProcesses = (Get-Process $strAppName -ErrorAction SilentlyContinue);
		if ([String]::IsNullOrEmpty($objProcesses)){
			$objReturn.Message = "No instances of $strAppName found.";
		}
		else{
			$objReturn.Message = "Success";
			$objReturn.Results = @($objProcesses).Count;
			$objReturn.Returns = $objProcesses;
		}

		return $objReturn;
	}

	function LoggedInUser{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strComp
		)
		#$strComp = The computer to check.  ($Env:ComputerName)

		#if([String]::IsNullOrEmpty($strComp)){
		#	#$strComp = "ALSDCP002656";		#Henry Laptop;
		#	$strComp = Read-Host 'What computer? (i.e. ALSDCP002656)';
		#}

		#check if a user is logged in
		$Error.Clear();
		$ErrorActionPreference = 'SilentlyContinue';
		$strLoggedIn = Get-WmiObject -class win32_computerSystem -computer:$strComp | Select-Object username;
		$ErrorActionPreference = 'Continue';
		if ($strLoggedIn.username.Length -gt 0){
			$strRet = $strLoggedIn.username + " is currently logged in to " + $strComp + ".";
		}else{
			if ($Error){
				$strRet = "Error: " + [String]$Error + ". Trying to check " + $strComp + " for logged in users.";
			}
			else{
				$strRet = "No user is logged in locally to " + $strComp + ".";
			}
		}

		#Write-Host $strRet;
		return $strRet;
	}

	function SetRegistry{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strRegPath, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strProp, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strValue, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strType = "String", 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strRemoteSys = "."
		)
		#Returns an array with results info.
		#$strRegPath = The Registry Path (Key) to set data to, or to create.  (i.e. "HKEY_LOCAL_MACHINE:\SOFTWARE\NMCI\ITSS-Tools\ASCII")
			#HKEY_CLASSES_ROOT
			#HKEY_CURRENT_CONFIG
			#HKEY_CURRENT_USER
			#HKEY_DYN_DATA
			#HKEY_LOCAL_MACHINE
			#HKEY_PERFORMANCE_DATA
			#HKEY_USERS
		#$strProp = The Registry Key Property, if any, to update/create with $strValue. (i.e. "(Default)" or "Version").
		#$strValue = The Value to enter.  (i.e. "0.23")
		#$strType = The Type/Kind that strValue is to be stored as.  [Microsoft.Win32.RegistryValueKind]
			#DWord = Integer
			#String = String
		#$strRemoteSys = A remote system to set the registry of.  ("." specifies local system, the default)

		#registry key = The Reg Path.
		#registry key property = The "attributes" / "fields" under the Key.

		#Remote Registry:
		#https://psremoteregistry.codeplex.com/		->	(http://blogs.microsoft.co.il/scriptfanatic/2010/01/10/remote-registry-powershell-module/)
		#http://stackoverflow.com/questions/1133335/openremotebasekey-credentials
		#https://social.technet.microsoft.com/Forums/windowsserver/en-US/14f33784-09a0-49be-8036-73921181fa3c/microsoftwin32registrykeyopenremotebasekey?forum=winserverpowershell

		$arrRet = $null;
		if (!([String]::IsNullOrEmpty($strRemoteSys))){
			#Write-Host "    Starting remote registry connection against: '$strRemoteSys'.";
			$strRoot = $strRegPath.SubString(0,$strRegPath.IndexOf(":"));
			switch ($strRoot){
				"HKEY_CLASSES_ROOT "{
					$strRoot = [Microsoft.Win32.RegistryHive]::ClassesRoot;
				}
				"HKEY_CURRENT_CONFIG"{
					$strRoot = [Microsoft.Win32.RegistryHive]::CurrentConfig;
				}
				"HKEY_CURRENT_USER"{
					$strRoot = [Microsoft.Win32.RegistryHive]::CurrentUser;
				}
				"HKEY_DYN_DATA"{
					$strRoot = [Microsoft.Win32.RegistryHive]::DynData;
				}
				"HKEY_LOCAL_MACHINE"{
					$strRoot = [Microsoft.Win32.RegistryHive]::LocalMachine;
				}
				"HKEY_PERFORMANCE_DATA"{
					$strRoot = [Microsoft.Win32.RegistryHive]::PerformanceData;
				}
				"HKEY_USERS"{
					$strRoot = [Microsoft.Win32.RegistryHive]::Users;
				}
				default{
					$arrRet = @("Error connecting to the (unknown) Registry root '$strRoot' of '$strRemoteSys'.", "-1");
					return $arrRet;
				}
			}
			#Write-Host "    Registry Hive is: [$strRoot]. `r`n";
			$strKey = $strRegPath.SubString($strRegPath.IndexOf(":") + 2);

			$Error.Clear();
			#Connect to "Root" of the Registry.
			$objReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($strRoot, $strRemoteSys);
			if ($Error){
				$strRoot = $strRegPath.SubString(0,$strRegPath.IndexOf(":"));
				$arrRet = @("Error connecting to the Registry '$strRoot' of '$strRemoteSys'.", $Error);
			}
			else{
				#Write-Host "    Open remote subkey: [$strKey].";
				$Error.Clear();
				$subKey = $objReg.OpenSubKey($strKey, $True);
				#$subKey | GM;
				if ($Error){
					$arrRet = @("Error connecting to the Registry Key '$strKey' of '$strRemoteSys'.", $Error);
				}
				else{
					if (([String]::IsNullOrEmpty($subKey))){
						#Key does not exist.
						$subKey = $null;
						$Error.Clear();
						$subKey = $objReg.CreateSubKey($strKey);
						if (($Error) -or ([String]::IsNullOrEmpty($subKey))){
							$arrRet = @("Error creating the path '$strRegPath'.", $Error);
						}
						else{
							$arrRet = @("Created Key '$strRegPath' on '$strRemoteSys'.", $subKey);
						}
					}

					if (!($Error)){
						if ([String]::IsNullOrEmpty($strProp)){
							#Set/Update Key
						}
						else{
							#Create/Update Property
							$objResults = $null;
							if ([String]::IsNullOrEmpty($subKey)){
								$Error.Clear();
								$subKey = $objReg.OpenSubKey($strKey, $True);
								if ($Error){
									$arrRet = @("Error re-connecting to the Registry Key '$strKey' of '$strRemoteSys'.", $Error);
								}
							}
							if ($strProp -eq "(Default)"){
								$strProp = "";
							}
							$Error.Clear();
							#$objResults = $subKey.SetValue($strProp, $strValue, [Microsoft.Win32.RegistryValueKind]::$strType);
							$objResults = $subKey.SetValue($strProp, $strValue, $strType);
							if ($Error){
								$arrRet = @("Error creating/updating the Property '$strProp' with '$strValue', of type '$strType', on '$strRemoteSys'.", $Error);
							}
							else{
								$arrRet = @("Created Property '$strProp' with '$strValue', of type '$strType', on '$strRemoteSys'.", $strRegPath);
							}
						}
					}
				}
			}
			#Write-Host "    Closing remote registry connection on: '$strRemoteSys'.";
			$objReg.Close();
		}

		return $arrRet;
	}
