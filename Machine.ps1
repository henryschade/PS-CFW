###########################################
# Updated Date:	20 July 2016
# Purpose:		Routines that require a Computer, or that interact w/ a Computer.
# Requirements: None
##########################################

<# ---=== Change Log ===---
	#Changes for 28 June 2016:
		#Added Change Log.

#>



	function CheckIfOnline{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strComp, 
			[ValidateNotNull()][Parameter(Mandatory=$False)]$bolUsers = $True, 
			[ValidateNotNull()][Parameter(Mandatory=$False)]$bolRetIP = $False
		)
		#Returns a string with the results.
			#"Host ($strComp) cannot be resolved, by DNS."
			#"Host ($strComp) is online."
				#"Host ($strComp) is online." (with a list of logged in users)
			#"Host ($strComp) not online (ping failed)."
		#$strComp = The computer to check.  ($Env:ComputerName)
		#$bolUsers = True or False.  Check if users are logged in.
		#$bolRetIP = True False.  Return an array, instead of a string.  @($strRet, $strIP).

		#if ([String]::IsNullOrEmpty($strComp)){
		#	#$strComp = "ALSDCP002656";		#Henry Laptop;
		#	#$strComp = "ALSDNI390014";		#Andrew Laptop;
		#	$strComp = Read-Host 'What computer? (i.e. ALSDCP002656)';
		#}

		#$strIP = "x.x.x.x";
		$Error.Clear();
		$ErrorActionPreference = 'SilentlyContinue';
		$strIP = [System.Net.DNS]::GetHostAddresses($strComp);
		if ($Error){
			$Error.Clear();
			[String]$strIP = [String](Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $strComp -Namespace "root\CIMV2" | WHERE{$_.IPEnabled -eq "True"}).IPAddress;
			if (($Error) -or ([String]::IsNullOrEmpty($strIP))){
				$strIP = "x.x.x.x";
			}
		}
		$ErrorActionPreference = 'Continue';
		#Write-Host $strIP;
		if (($strIP.GetType()).Name -eq "IPAddress[]"){
			$strIP = $strIP[0].IPAddressToString;
		}

		if (([String]::IsNullOrEmpty($strIP)) -or ($strIP -eq "x.x.x.x")){
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

		if ($bolRetIP){
			return @($strRet, $strIP);
		}
		else{
			return $strRet;
		}
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

	function DeleteRegistry{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strRegPath, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strProp, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strRemoteSys = "."
		)
		#Returns an array with results info.
		#$strRegPath = The Registry Path (Key) to Delete.  (i.e. "HKEY_LOCAL_MACHINE:\SOFTWARE\NMCI\ITSS-Tools\ASCII" or "HKEY_LOCAL_MACHINE\SOFTWARE\NMCI\ITSS-Tools\ASCII")
			#HKEY_CLASSES_ROOT
			#HKEY_CURRENT_CONFIG
			#HKEY_CURRENT_USER
			#HKEY_DYN_DATA
			#HKEY_LOCAL_MACHINE
			#HKEY_PERFORMANCE_DATA
			#HKEY_USERS
		#$strProp = The Registry Key Property, if any, to Delete. (i.e. "Version").
		#$strRemoteSys = A remote system to Delete the registry key of.  ("." specifies local system, the default)

		#https://social.technet.microsoft.com/Forums/scriptcenter/en-US/daae2835-bd41-4c7d-81c9-7a48e43f0b44/deleting-registry-keys?forum=winserverpowershell
		#registry key = The Reg Path.
		#registry key property = The "attributes" / "fields" under the Key.

		$arrRet = $null;
		if (!([String]::IsNullOrEmpty($strRemoteSys))){
			#Write-Host "    Starting remote registry connection against: '$strRemoteSys'.";
			if ($strRegPath.IndexOf(":") -gt 0){
				#Has a ":" seperator.
				$strRoot = $strRegPath.SubString(0,$strRegPath.IndexOf(":"));
			}
			else{
				$strRoot = $strRegPath.SubString(0,$strRegPath.IndexOf("\"));
			}
			switch ($strRoot){
				"HKEY_CLASSES_ROOT"{
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
				{($strRoot -eq "HKEY_LOCAL_MACHINE") -or ($strRoot -eq "HKLM")}{
					$strRoot = [Microsoft.Win32.RegistryHive]::LocalMachine;
				}
				"HKEY_PERFORMANCE_DATA"{
					$strRoot = [Microsoft.Win32.RegistryHive]::PerformanceData;
				}
				{($strRoot -eq "HKEY_USERS") -or ($strRoot -eq "HKU")}{
					$strRoot = [Microsoft.Win32.RegistryHive]::Users;
				}
				default{
					$arrRet = @("Error connecting to the (unknown) Registry root '$strRoot' of '$strRemoteSys'.", "-1");
					return $arrRet;
				}
			}
			#Write-Host "    Registry Hive is: [$strRoot]. `r`n";
			if ($strRegPath.IndexOf(":") -gt 0){
				#Has a ":" seperator.
				$strKey = $strRegPath.SubString($strRegPath.IndexOf(":") + 2);
			}
			else{
				$strKey = $strRegPath.SubString($strRegPath.IndexOf("\") + 1);
			}

			$Error.Clear();
			#Connect to "Root" of the Registry.
			$objReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($strRoot, $strRemoteSys);
			if ($Error){
				if ($strRegPath.IndexOf(":") -gt 0){
					#Has a ":" seperator.
					$strRoot = $strRegPath.SubString(0,$strRegPath.IndexOf(":"));
				}
				else{
					$strRoot = $strRegPath.SubString(0,$strRegPath.IndexOf("\"));
				}
				$arrRet = @("Error connecting to the Registry '$strRoot' of '$strRemoteSys'.", $Error);
			}
			else{
				#Do remove here
				if ([String]::IsNullOrEmpty($strProp)){
					$Error.Clear();
					$objKey = $objReg.DeleteSubKey($strKey);
					if ($Error){
						$arrRet = @("Error Deleting the Registry Key '$strRegPath' on '$strRemoteSys'.", $Error);
					}
					else{
						$arrRet = @("Deleted Registry Key '$strRegPath' on '$strRemoteSys'.", $strRegPath);
					}
				}
				else{
					$Error.Clear();
					$subKey = $objReg.OpenSubKey($strKey, $True);
					if ($Error){
						$arrRet = @("Error opening the Key '$strRegPath', on '$strRemoteSys'.", $Error);
					}
					else{
						$Error.Clear();
						$objResults = $subKey.DeleteValue($strProp)
						if ($Error){
							$arrRet = @("Error deleting the Property '$strProp', under Key '$strRegPath', on '$strRemoteSys'.", $Error);
						}
						else{
							$arrRet = @("Deleted '$strProp', under Key '$strRegPath', on '$strRemoteSys'.", $strRegPath);
						}
					}
				}
			}

			#Write-Host "    Closing remote registry connection on: '$strRemoteSys'.";
			$objReg.Close();
		}

		return $arrRet;
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

	function EnableWinRM{
		param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strComp
		)
		#a modified copy of Chris's original code.
		#$strComp = The computer to enable WinRM on.

		$bolReturn = $True;
		$ESDWorkingFolder = "\\$strComp\C$\Program Files\NMCI\ESD";
		#$ESDTasksPath = "\\nawesdnifs101v.nadsuswe.nads.navy.mil\NMCIISF02$\ITSS-Tools\PS-CFW\ESDTasks.XML";

		if (Test-Path $ESDWorkingFolder){
			$strXML = "<?xml version=`"1.0`" encoding=`"utf-8`"?>`r`n";
			$strXML = $strXML + "<Root xmlns:xsi=`"http://www.w3.org/2001/XMLSchema-instance`" xsi:noNamespaceSchemaLocation=`"\\10.70.13.17\Advanced Publisher\In Process\EMT-Tools\EMT Starter\Includes\EMTTasksSchema.xsd`">`r`n";
			$strXML = $strXML + "  <Task xsi:type=`"RunCommandLineTaskType`" name=`"EnableWinRM`">`r`n";
			$strXML = $strXML + "    <Applet>RunCommandLine</Applet>`r`n";
			$strXML = $strXML + "    <LogName>[ESD]ESDTasks</LogName>`r`n";
			$strXML = $strXML + "    <LogFolder>%IDMLOG%</LogFolder>`r`n";
			$strXML = $strXML + "    <Parameters>`r`n";
			$strXML = $strXML + "      <Enum index=`"2`">WINRM quickconfig -quiet</Enum>`r`n";
			$strXML = $strXML + "    </Parameters>`r`n";
			$strXML = $strXML + "  </Task>`r`n";
			$strXML = $strXML + "</Root>`r`n";

			$strResults = mkdir "C:\Users\Public\ITSS-Tools\" -ErrorAction SilentlyContinue;
			#$strResults = cmd /c "mkdir `"C:\Users\Public\ITSS-Tools\`"" '2>&1';			#This allows something.......

			$Error.Clear();
			$strXML | Out-File ("C:\Users\Public\ITSS-Tools\ESDTasks.XML");
			if ($Error){
				$ESDTasksPath = "\\nawesdnifs101v.nadsuswe.nads.navy.mil\NMCIISF02$\ITSS-Tools\PS-CFW\ESDTasks.XML";
			}
			else{
				$ESDTasksPath = "C:\Users\Public\ITSS-Tools\ESDTasks.XML";
			}
			#$ErrorActionPreference = 'SilentlyContinue';
			$strResults = Copy-Item -Path $ESDTasksPath -Destination $ESDWorkingFolder -Force;
			#$ErrorActionPreference = 'Continue';
			$strResults = (Start-Process -FilePath 'C:\Progra~2\Hewlett-Packard\HPCA\Agent\radntfyc.exe' -ArgumentList $([String]::Format("{0} -p 1 EMT -File=`"C:\Program Files\NMCI\ESD\ESDTasks.xml`" EnableWinRM", $strComp)));
		}
		else{
			#Can NOT UNC to Machine.
			$bolReturn = $False;
		}

		return $bolReturn;
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
		#$strRegPath = The Registry Path (Key) to get data from.  (i.e. "HKEY_LOCAL_MACHINE:\SOFTWARE\NMCI\ITSS-Tools\ASCII" or "HKEY_LOCAL_MACHINE\SOFTWARE\NMCI\ITSS-Tools\ASCII")
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
			if ($strRegPath.IndexOf(":") -gt 0){
				#Has a ":" seperator.
				$strRoot = $strRegPath.SubString(0,$strRegPath.IndexOf(":"));
			}
			else{
				$strRoot = $strRegPath.SubString(0,$strRegPath.IndexOf("\"));
			}
			switch ($strRoot){
				"HKEY_CLASSES_ROOT"{
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
				{($strRoot -eq "HKEY_LOCAL_MACHINE") -or ($strRoot -eq "HKLM")}{
					$strRoot = [Microsoft.Win32.RegistryHive]::LocalMachine;
				}
				"HKEY_PERFORMANCE_DATA"{
					$strRoot = [Microsoft.Win32.RegistryHive]::PerformanceData;
				}
				{($strRoot -eq "HKEY_USERS") -or ($strRoot -eq "HKU")}{
					$strRoot = [Microsoft.Win32.RegistryHive]::Users;
				}
				default{
					$arrRet = @("Error connecting to the (unknown) Registry root '$strRoot' of '$strRemoteSys'.", "-1");
					return $arrRet;
				}
			}
			#Write-Host "    Registry Hive is: [$strRoot]. `r`n";
			if ($strRegPath.IndexOf(":") -gt 0){
				#Has a ":" seperator.
				$strKey = $strRegPath.SubString($strRegPath.IndexOf(":") + 2);
			}
			else{
				$strKey = $strRegPath.SubString($strRegPath.IndexOf("\") + 1);
			}

			$Error.Clear();
			#Connect to "Root" of the Registry.
			$objReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($strRoot, $strRemoteSys);
			if ($Error){
				if ($strRegPath.IndexOf(":") -gt 0){
					#Has a ":" seperator.
					$strRoot = $strRegPath.SubString(0,$strRegPath.IndexOf(":"));
				}
				else{
					$strRoot = $strRegPath.SubString(0,$strRegPath.IndexOf("\"));
				}
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

	function GetSoftwareInfo{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strComp, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strSoftware = $null, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bWriteScreen = $False
		)
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters passed in.
			#$objReturn.Results		= # of $strSoftware installations found.
			#$objReturn.Message		= "Success" or the error message.
			#$objReturn.Returns		= An array.  @($strSoftwareName, $strSoftwareLastUsed)
		#$strComp = Computer to query.
		#$strSoftware = Software to look for, partial names are OK too.
		#$bWriteScreen = Output "Errors" to the screen.

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

		$arrSoftwareName = @();
		$arrSoftwareLastUsed = @();
		$arrSoftware = $null;
		$strSoftwareName = $strSoftware;
		$strSoftwareLastUsed = "-";

		$ErrorActionPreference = 'SilentlyContinue';
		$Error.Clear();
		$arrSoftware = Get-WMIObject Win32_SoftwareFeature -ComputerName $strComp | Select-Object ProductName,LastUse -unique | Sort-Object LastUse;
		#Sorting by "LastUse" makes it so that the most last used date is the last one.
		$ErrorActionPreference = 'Continue';

		#if ($Error){
		if ([String]::IsNullOrEmpty($arrSoftware)){
			if ($bWriteScreen){
				Write-Host "  Error connecting to Machine $strComp. `r`n $Error";
			}
			$strSoftwareLastUsed = "WMI Error: " + $Error;
			$objReturn.Message = $strSoftwareLastUsed;
		}
		else{
			$objReturn.Message = "Success";
			if (!([String]::IsNullOrEmpty($strSoftware))){
				if ($bWriteScreen){
					Write-Host "  Looking for $strSoftware";
				}
				$strSoftwareInfo = $arrSoftware -Match $strSoftware;
				if ((!([String]::IsNullOrEmpty($strSoftwareInfo))) -and ($strSoftwareInfo -ne $False)){
					if ($strSoftwareInfo.Count -ne 1){
						#Sorting by "LastUse" makes it so that the most last used date is the last one.
						$strSoftwareName = $strSoftwareInfo[$strSoftwareInfo.Count - 1].ProductName;
						$strSoftwareLastUsed = $strSoftwareInfo[$strSoftwareInfo.Count - 1].LastUse;
					}
					else{
						$strSoftwareName = $strSoftwareInfo[0].ProductName;
						$strSoftwareLastUsed = $strSoftwareInfo[0].LastUse;
					}
					$strSoftwareLastUsed = $strSoftwareLastUsed.SubString(0, 8);
					$strSoftwareLastUsed = $strSoftwareLastUsed.SubString(0,4) + "-" + $strSoftwareLastUsed.SubString(4,2) + "-" + $strSoftwareLastUsed.SubString(6,2);
				}
				else{
					$strSoftwareLastUsed = "Not installed/found.";
				}
			}
			else{
				if ($bWriteScreen){
					Write-Host "  No criteria provided";
				}
				foreach ($strPackage in $arrSoftware){
					$arrSoftwareName += $strPackage.ProductName;
					$arrSoftwareLastUsed += $strPackage.LastUse;
				}

				$strSoftwareName = $arrSoftwareName;
				$strSoftwareLastUsed = $arrSoftwareLastUsed;
			}

			$objReturn.Returns = @($strSoftwareName, $strSoftwareLastUsed);
			$objReturn.Results = @(($objReturn.Returns)[0]).Count;
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
		$strLoggedIn = Get-WmiObject -class Win32_ComputerSystem -computer:$strComp | Select-Object username;
		$ErrorActionPreference = 'Continue';
		if ($strLoggedIn.username.Length -gt 0){
			$strRet = $strLoggedIn.username + " is currently logged in to " + $strComp + ".";
		}else{
			if ($Error){
				$strRet = "Error trying to check " + $strComp + " for logged in users.  Error: " + [String]$Error + ".";
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
		#$strRegPath = The Registry Path (Key) to set data to, or to create.  (i.e. "HKEY_LOCAL_MACHINE:\SOFTWARE\NMCI\ITSS-Tools\ASCII" or "HKEY_LOCAL_MACHINE\SOFTWARE\NMCI\ITSS-Tools\ASCII")
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
			if ($strRegPath.IndexOf(":") -gt 0){
				#Has a ":" seperator.
				$strRoot = $strRegPath.SubString(0,$strRegPath.IndexOf(":"));
			}
			else{
				$strRoot = $strRegPath.SubString(0,$strRegPath.IndexOf("\"));
			}
			switch ($strRoot){
				{($strRoot -eq "HKEY_CLASSES_ROOT") -or ($strRoot -eq "HKCR")}{
					$strRoot = [Microsoft.Win32.RegistryHive]::ClassesRoot;
				}
				"HKEY_CURRENT_CONFIG"{
					$strRoot = [Microsoft.Win32.RegistryHive]::CurrentConfig;
				}
				{($strRoot -eq "HKEY_CURRENT_USER") -or ($strRoot -eq "HKCU")}{
					$strRoot = [Microsoft.Win32.RegistryHive]::CurrentUser;
				}
				"HKEY_DYN_DATA"{
					$strRoot = [Microsoft.Win32.RegistryHive]::DynData;
				}
				{($strRoot -eq "HKEY_LOCAL_MACHINE") -or ($strRoot -eq "HKLM")}{
					$strRoot = [Microsoft.Win32.RegistryHive]::LocalMachine;
				}
				"HKEY_PERFORMANCE_DATA"{
					$strRoot = [Microsoft.Win32.RegistryHive]::PerformanceData;
				}
				{($strRoot -eq "HKEY_USERS") -or ($strRoot -eq "HKU")}{
					$strRoot = [Microsoft.Win32.RegistryHive]::Users;
				}
				default{
					$arrRet = @("Error connecting to the (unknown) Registry root '$strRoot' of '$strRemoteSys'.", "-1");
					return $arrRet;
				}
			}
			#Write-Host "    Registry Hive is: [$strRoot]. `r`n";
			if ($strRegPath.IndexOf(":") -gt 0){
				#Has a ":" seperator.
				$strKey = $strRegPath.SubString($strRegPath.IndexOf(":") + 2);
			}
			else{
				$strKey = $strRegPath.SubString($strRegPath.IndexOf("\") + 1);
			}

			$Error.Clear();
			#Connect to "Root" of the Registry.
			$objReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($strRoot, $strRemoteSys);
			if ($Error){
				if ($strRegPath.IndexOf(":") -gt 0){
					#Has a ":" seperator.
					$strRoot = $strRegPath.SubString(0,$strRegPath.IndexOf(":"));
				}
				else{
					$strRoot = $strRegPath.SubString(0,$strRegPath.IndexOf("\"));
				}
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
							#$objResults = $subKey.SetValue($strKey, $strValue);
							#if ($Error){
							#	$arrRet = @("Error updating the Key '$strKey' to '$strValue' on '$strRemoteSys'.", $Error);
							#}
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




	function RunRemotly{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strCommand, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strRemoteSys = ".", 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bInteractive
		)

		#http://stackoverflow.com/questions/4147821/start-a-windows-service-and-launch-cmd
		#https://blogs.msdn.microsoft.com/alejacma/2007/12/20/how-to-call-createprocesswithlogonw-createprocessasuser-in-net/


		#http://serverfault.com/questions/637743/execute-command-on-remote-computer-and-show-ui-to-logged-on-user
			#https://msdn.microsoft.com/en-us/library/windows/desktop/aa379608(v=vs.85).aspx
		#http://poshcode.org/1856
		#http://stackoverflow.com/questions/7162604/get-cached-credentials-in-powershell-from-windows-7-credential-manager

		#http://stackoverflow.com/questions/13290296/createprocessasuser-error-5

		#Get primary access token of a logged on user....
		#WTSQueryUserToken((uint)session.SessionId, out userPrimaryAccessToken)
		#execute application in the security context of a logged on user....
		#CreateProcessAsUser(userPrimaryAccessToken, null, cmdLine, ref saProcessAttributes, ref saThreadAttributes, false, 0, IntPtr.Zero, null, ref si, out pi)


		#try {
		#	[CreateProcessUtility.CreateProcessCaller]::modifyEnvParamWrapper2($command, $strDomain, $strName, $strPassword)
		#	return $True
		#} catch {
		#	write-host "Unable to modify regestry entry: " $_
		#	return $False
		#}



		#Testing...
		#[IntPtr]$userToken = [Security.Principal.WindowsIdentity]::GetCurrent().Token
		#$identity = New-Object Security.Principal.WindowsIdentity $userToken
		#$context = $identity.Impersonate()
		#([Security.Principal.WindowsIdentity]::GetCurrent() | Format-Table Name, Token, User, Groups -Auto | Out-String)



		#Following does work, on local system, so maybe w/ PS-Session (Invoke-Command) could work?
			#http://stackoverflow.com/questions/3705321/enter-pssession-is-not-working-in-my-powershell-script
				#Enter-PSSession -computerName ALSDCP002656;
				#cd c:\
				#Exit-PSSession;
			#or
				#$s = New-PSSession -computerName ALSDCP002656;
				#Invoke-Command -Session $s -Scriptblock {mkdir c:\Testing-2-See;};
				#Remove-PSSession $s;
		#http://stackoverflow.com/questions/16686122/calling-createprocess-from-powershell
		$si = New-Object STARTUPINFO;
		$pi = New-Object PROCESS_INFORMATION;

		$si.cb = [System.Runtime.InteropServices.Marshal]::SizeOf($si);
		$si.wShowWindow = [ShowWindow]::SW_SHOW;

		$pSec = New-Object SECURITY_ATTRIBUTES;
		$tSec = New-Object SECURITY_ATTRIBUTES;
		$pSec.Length = [System.Runtime.InteropServices.Marshal]::SizeOf($pSec);
		$tSec.Length = [System.Runtime.InteropServices.Marshal]::SizeOf($tSec);

		$strCommand = "c:\windows\notepad.exe";
		$bolRet = [Kernel32]::CreateProcess($strCommand, $null, [ref] $pSec, [ref] $tSec, $false, [CreationFlags]::NONE, [IntPtr]::Zero, "c:", [ref] $si, [ref] $pi);
		#$bolRet = [Advapi32]::CreateProcessAsUser($userPrimaryAccessToken, $strCommand, $null, [ref] $pSec, [ref] $tSec, $false, [CreationFlags]::NONE, [IntPtr]::Zero, "c:", [ref] $si, [ref] $pi);

		$intErr = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error();
		if ([String]::IsNullOrEmpty($pi)){
			return $intErr;
		}
		else{
			#$pi;
			return $pi;
		}
	}

	#http://stackoverflow.com/questions/13290296/createprocessasuser-error-5
	if ($True -eq $False){
		$createProcess = @'
			using System;
			using System.Collections.Generic;
			using System.Text;
			using System.Runtime.InteropServices;
			using Microsoft.Win32;
			using System.IO;
			using System.Security.Principal;

			namespace CreateProcessUtility{
				class Win32{
					#region "CONTS"
					const UInt32 INFINITE = 0xFFFFFFFF;
					const UInt32 WAIT_FAILED = 0xFFFFFFFF;

					#endregion

					#region "ENUMS"

					[Flags]
					public enum LogonType
					{
						LOGON32_LOGON_INTERACTIVE = 2,
						LOGON32_LOGON_NETWORK = 3,
						LOGON32_LOGON_BATCH = 4,
						LOGON32_LOGON_SERVICE = 5,
						LOGON32_LOGON_UNLOCK = 7,
						LOGON32_LOGON_NETWORK_CLEARTEXT = 8,
						LOGON32_LOGON_NEW_CREDENTIALS = 9
					}

					[Flags]
					public enum LogonProvider
					{
						LOGON32_PROVIDER_DEFAULT = 0,
						LOGON32_PROVIDER_WINNT35,
						LOGON32_PROVIDER_WINNT40,
						LOGON32_PROVIDER_WINNT50
					}

					#endregion

					#region "STRUCTS"

					[StructLayout(LayoutKind.Sequential)]
					public struct STARTUPINFO
					{
						public Int32 cb;
						public String lpReserved;
						public String lpDesktop;
						public String lpTitle;
						public Int32 dwX;
						public Int32 dwY;
						public Int32 dwXSize;
						public Int32 dwYSize;
						public Int32 dwXCountChars;
						public Int32 dwYCountChars;
						public Int32 dwFillAttribute;
						public Int32 dwFlags;
						public Int16 wShowWindow;
						public Int16 cbReserved2;
						public IntPtr lpReserved2;
						public IntPtr hStdInput;
						public IntPtr hStdOutput;
						public IntPtr hStdError;
					}

					[StructLayout(LayoutKind.Sequential)]
					public struct PROCESS_INFORMATION
					{
						public IntPtr hProcess;
						public IntPtr hThread;
						public Int32 dwProcessId;
						public Int32 dwThreadId;
					}

					#endregion

					#region "FUNCTIONS (P/INVOKE)"

					[StructLayout(LayoutKind.Sequential)]
					public struct ProfileInfo {
						public int dwSize; 
						public int dwFlags;
						public String lpUserName; 
						public String lpProfilePath; 
						public String lpDefaultPath; 
						public String lpServerName; 
						public String lpPolicyPath; 
						public IntPtr hProfile; 
					}

					[DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
					public static extern Boolean LogonUser 
					(
						String lpszUserName,
						String lpszDomain,
						String lpszPassword,
						LogonType dwLogonType,
						LogonProvider dwLogonProvider,
						out IntPtr phToken
					);

					[DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
					public static extern Boolean CreateProcessAsUser 
					(
						IntPtr hToken,
						String lpApplicationName,
						String lpCommandLine,
						IntPtr lpProcessAttributes,
						IntPtr lpThreadAttributes,
						Boolean bInheritHandles,
						Int32 dwCreationFlags,
						IntPtr lpEnvironment,
						String lpCurrentDirectory,
						ref STARTUPINFO lpStartupInfo,
						out PROCESS_INFORMATION lpProcessInformation
					);

					[DllImport("kernel32.dll", SetLastError = true)]
					public static extern UInt32 WaitForSingleObject 
					(
						IntPtr hHandle,
						UInt32 dwMilliseconds
					);

					[DllImport("kernel32", SetLastError=true)]
					public static extern Boolean CloseHandle (IntPtr handle);

					[DllImport("userenv.dll", SetLastError = true, CharSet = CharSet.Auto)]
					public static extern bool LoadUserProfile(IntPtr hToken, ref ProfileInfo lpProfileInfo);

					[DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
					public static extern int DuplicateToken(IntPtr hToken, int impersonationLevel, ref IntPtr hNewToken);

					#endregion

					#region "FUNCTIONS"

					public static void LaunchCommand2(string strCommand, string strDomain, string strName, string strPassword)
					{
						// Variables
						WindowsIdentity m_ImpersonatedUser;
						IntPtr tokenDuplicate = IntPtr.Zero;
						PROCESS_INFORMATION processInfo = new PROCESS_INFORMATION();
						STARTUPINFO startInfo = new STARTUPINFO();
						Boolean bResult = false;
						IntPtr hToken = IntPtr.Zero;
						UInt32 uiResultWait = WAIT_FAILED;
						string executableFile = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe";
						const int SecurityImpersonation = 2;

						try 
						{
							// Logon user
							bResult = Win32.LogonUser(
								strName,
								strDomain,
								strPassword,
								Win32.LogonType.LOGON32_LOGON_INTERACTIVE,
								Win32.LogonProvider.LOGON32_PROVIDER_DEFAULT,
								out hToken
							);
							if (!bResult) { throw new Exception("Logon error #" + Marshal.GetLastWin32Error()); }

							 #region LoadUserProfile
								ProfileInfo currentProfile = new ProfileInfo();
								currentProfile.dwSize = Marshal.SizeOf(currentProfile);
								currentProfile.lpUserName = strName;
								currentProfile.dwFlags = 1;                        
								Boolean bResult2 = LoadUserProfile(hToken, ref currentProfile);
								Console.WriteLine(bResult2);

								if (!bResult2) { throw new Exception("LoadUserProfile error #" + Marshal.GetLastWin32Error()); }
							   Console.WriteLine(currentProfile.hProfile + "----"+IntPtr.Zero);


							   if (currentProfile.hProfile == IntPtr.Zero){
									Console.WriteLine("LoadUserProfile() failed - HKCU handle was not loaded. Error code: " + Marshal.GetLastWin32Error());
									throw new Exception("LoadUserProfile error #" + Marshal.GetLastWin32Error());
								}
							 #endregion

							// Create process
							startInfo.cb = Marshal.SizeOf(startInfo);
							startInfo.lpDesktop = "winsta0\\default";

							Console.WriteLine("Before impersonation: " + WindowsIdentity.GetCurrent().Name);

							if (DuplicateToken(hToken, SecurityImpersonation, ref tokenDuplicate) != 0){
							 m_ImpersonatedUser = new WindowsIdentity(tokenDuplicate);

								if(m_ImpersonatedUser.Impersonate() != null){
									Console.WriteLine("After Impersonation succeeded: " + Environment.NewLine + "User Name: " + WindowsIdentity.GetCurrent(TokenAccessLevels.MaximumAllowed).Name + Environment.NewLine + "SID: " + WindowsIdentity.GetCurrent(TokenAccessLevels.MaximumAllowed).User.Value);
									Console.WriteLine(m_ImpersonatedUser);
								}

								bResult = Win32.CreateProcessAsUser(
								tokenDuplicate, 
								executableFile, 
								strCommand, 
								IntPtr.Zero,
								IntPtr.Zero,
								false,
								0,
								IntPtr.Zero,
								null,
								ref startInfo,
								out processInfo
							);
							if (!bResult) { throw new Exception("CreateProcessAsUser error #" + Marshal.GetLastWin32Error()); }
						}

							// Wait for process to end
							uiResultWait = WaitForSingleObject(processInfo.hProcess, INFINITE);
							if (uiResultWait == WAIT_FAILED) { throw new Exception("WaitForSingleObject error #" + Marshal.GetLastWin32Error()); }
						}
						finally 
						{
							// Close all handles
							CloseHandle(hToken);
							CloseHandle(processInfo.hProcess);
							CloseHandle(processInfo.hThread);
						}
					}

					#endregion
				}

				// Interface between powershell and C#    
				public class CreateProcessCaller
				{
					public static void modifyEnvParamWrapper2(string strCommand, string strDomain, string strName, string strPassword)
					{
						Win32.LaunchCommand2(strCommand, strDomain, strName, strPassword);
					}
				}
			}
'@;		#This MUST end w/ no leading spaces.

		Add-Type -TypeDefinition $createProcess -Language CSharp -IgnoreWarnings;
	}

	#http://stackoverflow.com/questions/16686122/calling-createprocess-from-powershell
	if ($True -eq $False){
		$createProcess2 = @"
			using System;
			using System.Diagnostics;
			using System.Runtime.InteropServices;

			[StructLayout(LayoutKind.Sequential)]
			public struct PROCESS_INFORMATION
			{
				public IntPtr hProcess;
				public IntPtr hThread;
				public uint dwProcessId;
				public uint dwThreadId;
			}

			[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
			public struct STARTUPINFO
			{
				public uint cb;
				public string lpReserved;
				public string lpDesktop;
				public string lpTitle;
				public uint dwX;
				public uint dwY;
				public uint dwXSize;
				public uint dwYSize;
				public uint dwXCountChars;
				public uint dwYCountChars;
				public uint dwFillAttribute;
				public STARTF dwFlags;
				public ShowWindow wShowWindow;
				public short cbReserved2;
				public IntPtr lpReserved2;
				public IntPtr hStdInput;
				public IntPtr hStdOutput;
				public IntPtr hStdError;
			}

			[StructLayout(LayoutKind.Sequential)]
			public struct SECURITY_ATTRIBUTES
			{
				public int length;
				public IntPtr lpSecurityDescriptor;
				public bool bInheritHandle;
			}

			[Flags]
			public enum CreationFlags : int
			{
				NONE = 0,
				DEBUG_PROCESS = 0x00000001,
				DEBUG_ONLY_THIS_PROCESS = 0x00000002,
				CREATE_SUSPENDED = 0x00000004,
				DETACHED_PROCESS = 0x00000008,
				CREATE_NEW_CONSOLE = 0x00000010,
				CREATE_NEW_PROCESS_GROUP = 0x00000200,
				CREATE_UNICODE_ENVIRONMENT = 0x00000400,
				CREATE_SEPARATE_WOW_VDM = 0x00000800,
				CREATE_SHARED_WOW_VDM = 0x00001000,
				CREATE_PROTECTED_PROCESS = 0x00040000,
				EXTENDED_STARTUPINFO_PRESENT = 0x00080000,
				CREATE_BREAKAWAY_FROM_JOB = 0x01000000,
				CREATE_PRESERVE_CODE_AUTHZ_LEVEL = 0x02000000,
				CREATE_DEFAULT_ERROR_MODE = 0x04000000,
				CREATE_NO_WINDOW = 0x08000000,
			}

			[Flags]
			public enum STARTF : uint
			{
				STARTF_USESHOWWINDOW = 0x00000001,
				STARTF_USESIZE = 0x00000002,
				STARTF_USEPOSITION = 0x00000004,
				STARTF_USECOUNTCHARS = 0x00000008,
				STARTF_USEFILLATTRIBUTE = 0x00000010,
				STARTF_RUNFULLSCREEN = 0x00000020,  // ignored for non-x86 platforms
				STARTF_FORCEONFEEDBACK = 0x00000040,
				STARTF_FORCEOFFFEEDBACK = 0x00000080,
				STARTF_USESTDHANDLES = 0x00000100,
			}

			public enum ShowWindow : short
			{
				SW_HIDE = 0,
				SW_SHOWNORMAL = 1,
				SW_NORMAL = 1,
				SW_SHOWMINIMIZED = 2,
				SW_SHOWMAXIMIZED = 3,
				SW_MAXIMIZE = 3,
				SW_SHOWNOACTIVATE = 4,
				SW_SHOW = 5,
				SW_MINIMIZE = 6,
				SW_SHOWMINNOACTIVE = 7,
				SW_SHOWNA = 8,
				SW_RESTORE = 9,
				SW_SHOWDEFAULT = 10,
				SW_FORCEMINIMIZE = 11,
				SW_MAX = 11
			}

			public static class Kernel32
			{
				[DllImport("kernel32.dll", SetLastError=true)]
				public static extern bool CreateProcess(
					string lpApplicationName, 
					string lpCommandLine, 
					ref SECURITY_ATTRIBUTES lpProcessAttributes, 
					ref SECURITY_ATTRIBUTES lpThreadAttributes,
					bool bInheritHandles, 
					CreationFlags dwCreationFlags, 
					IntPtr lpEnvironment,
					string lpCurrentDirectory, 
					ref STARTUPINFO lpStartupInfo, 
					out PROCESS_INFORMATION lpProcessInformation);
			}

			public static class Advapi32
			{
				[DllImport("Advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
				public static extern bool CreateProcessAsUser(
					IntPtr hToken,
					String lpApplicationName,
					String lpCommandLine,
					ref SECURITY_ATTRIBUTES lpProcessAttributes,
					ref SECURITY_ATTRIBUTES lpThreadAttributes,
					bool bInheritHandles,
					Int32 dwCreationFlags,
					IntPtr lpEnvironment,
					String lpCurrentDirectory,
					ref STARTUPINFO lpStartupInfo,
					out PROCESS_INFORMATION lpProcessInformation);
			}
"@

		Add-Type -TypeDefinition $createProcess2 -Language CSharp -IgnoreWarnings;
	}

	#http://stackoverflow.com/questions/4147821/start-a-windows-service-and-launch-cmd
	if ($True -eq $False){
		$createProcess3 = @'
			public static class ProcessAsCurrentUser
			{
				/// <summary>
				/// Connection state of a session.
				/// </summary>
				public enum ConnectionState
				{
					/// <summary>
					/// A user is logged on to the session.
					/// </summary>
					Active,
					/// <summary>
					/// A client is connected to the session.
					/// </summary>
					Connected,
					/// <summary>
					/// The session is in the process of connecting to a client.
					/// </summary>
					ConnectQuery,
					/// <summary>
					/// This session is shadowing another session.
					/// </summary>
					Shadowing,
					/// <summary>
					/// The session is active, but the client has disconnected from it.
					/// </summary>
					Disconnected,
					/// <summary>
					/// The session is waiting for a client to connect.
					/// </summary>
					Idle,
					/// <summary>
					/// The session is listening for connections.
					/// </summary>
					Listening,
					/// <summary>
					/// The session is being reset.
					/// </summary>
					Reset,
					/// <summary>
					/// The session is down due to an error.
					/// </summary>
					Down,
					/// <summary>
					/// The session is initializing.
					/// </summary>
					Initializing
				}

				[StructLayout(LayoutKind.Sequential)]
				class SECURITY_ATTRIBUTES
				{
					public int nLength;
					public IntPtr lpSecurityDescriptor;
					public int bInheritHandle;
				}

				[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
				struct STARTUPINFO
				{
					public Int32 cb;
					public string lpReserved;
					public string lpDesktop;
					public string lpTitle;
					public Int32 dwX;
					public Int32 dwY;
					public Int32 dwXSize;
					public Int32 dwYSize;
					public Int32 dwXCountChars;
					public Int32 dwYCountChars;
					public Int32 dwFillAttribute;
					public Int32 dwFlags;
					public Int16 wShowWindow;
					public Int16 cbReserved2;
					public IntPtr lpReserved2;
					public IntPtr hStdInput;
					public IntPtr hStdOutput;
					public IntPtr hStdError;
				}

				[StructLayout(LayoutKind.Sequential)]
				internal struct PROCESS_INFORMATION
				{
					public IntPtr hProcess;
					public IntPtr hThread;
					public int dwProcessId;
					public int dwThreadId;
				}

				enum LOGON_TYPE
				{
					LOGON32_LOGON_INTERACTIVE = 2,
					LOGON32_LOGON_NETWORK,
					LOGON32_LOGON_BATCH,
					LOGON32_LOGON_SERVICE,
					LOGON32_LOGON_UNLOCK = 7,
					LOGON32_LOGON_NETWORK_CLEARTEXT,
					LOGON32_LOGON_NEW_CREDENTIALS
				}

				enum LOGON_PROVIDER
				{
					LOGON32_PROVIDER_DEFAULT,
					LOGON32_PROVIDER_WINNT35,
					LOGON32_PROVIDER_WINNT40,
					LOGON32_PROVIDER_WINNT50
				}

				[Flags]
				enum CreateProcessFlags : uint
				{
					CREATE_BREAKAWAY_FROM_JOB = 0x01000000,
					CREATE_DEFAULT_ERROR_MODE = 0x04000000,
					CREATE_NEW_CONSOLE = 0x00000010,
					CREATE_NEW_PROCESS_GROUP = 0x00000200,
					CREATE_NO_WINDOW = 0x08000000,
					CREATE_PROTECTED_PROCESS = 0x00040000,
					CREATE_PRESERVE_CODE_AUTHZ_LEVEL = 0x02000000,
					CREATE_SEPARATE_WOW_VDM = 0x00000800,
					CREATE_SHARED_WOW_VDM = 0x00001000,
					CREATE_SUSPENDED = 0x00000004,
					CREATE_UNICODE_ENVIRONMENT = 0x00000400,
					DEBUG_ONLY_THIS_PROCESS = 0x00000002,
					DEBUG_PROCESS = 0x00000001,
					DETACHED_PROCESS = 0x00000008,
					EXTENDED_STARTUPINFO_PRESENT = 0x00080000,
					INHERIT_PARENT_AFFINITY = 0x00010000
				}

				[StructLayout(LayoutKind.Sequential)]
				public struct WTS_SESSION_INFO
				{
					public int SessionID;
					[MarshalAs(UnmanagedType.LPTStr)]
					public string WinStationName;
					public ConnectionState State;
				}

				[DllImport("wtsapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
				public static extern Int32 WTSEnumerateSessions(IntPtr hServer, int reserved, int version,
																ref IntPtr sessionInfo, ref int count);

				[DllImport("advapi32.dll", EntryPoint = "CreateProcessAsUserW", SetLastError = true, CharSet = CharSet.Auto)]
				static extern bool CreateProcessAsUser(
					IntPtr hToken,
					string lpApplicationName,
					string lpCommandLine,
					IntPtr lpProcessAttributes,
					IntPtr lpThreadAttributes,
					bool bInheritHandles,
					UInt32 dwCreationFlags,
					IntPtr lpEnvironment,
					string lpCurrentDirectory,
					ref STARTUPINFO lpStartupInfo,
					out PROCESS_INFORMATION lpProcessInformation);

				[DllImport("wtsapi32.dll")]
				public static extern void WTSFreeMemory(IntPtr memory);

				[DllImport("kernel32.dll")]
				private static extern UInt32 WTSGetActiveConsoleSessionId();

				[DllImport("wtsapi32.dll", SetLastError = true)]
				static extern int WTSQueryUserToken(UInt32 sessionId, out IntPtr Token);

				[DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
				public extern static bool DuplicateTokenEx(
					IntPtr hExistingToken,
					uint dwDesiredAccess,
					IntPtr lpTokenAttributes,
					int ImpersonationLevel,
					int TokenType,
					out IntPtr phNewToken);

				private const int TokenImpersonation = 2;
				private const int SecurityIdentification = 1;
				private const int MAXIMUM_ALLOWED = 0x2000000;
				private const int TOKEN_DUPLICATE = 0x2;
				private const int TOKEN_QUERY = 0x00000008;

				/// <summary>
				/// Launches a process for the current logged on user if there are any.
				/// If none, return false as well as in case of 
				/// 
				/// ##### !!! BEWARE !!! ####  ------------------------------------------
				/// This code will only work when running in a windows service (where it is really needed)
				/// so in case you need to test it, it needs to run in the service. Reason
				/// is a security privileg which only services have (SE_??? something, cant remember)!
				/// </summary>
				/// <param name="processExe"></param>
				/// <returns></returns>
				public static bool CreateProcessAsCurrentUser(string processExe)
				{

					IntPtr duplicate = new IntPtr();
					STARTUPINFO info = new STARTUPINFO();
					PROCESS_INFORMATION procInfo = new PROCESS_INFORMATION();

					Debug.WriteLine(string.Format("CreateProcessAsCurrentUser. processExe: " + processExe));

					IntPtr p = GetCurrentUserToken();

					bool result = DuplicateTokenEx(p, MAXIMUM_ALLOWED | TOKEN_QUERY | TOKEN_DUPLICATE, IntPtr.Zero, SecurityIdentification, SecurityIdentification, out duplicate);
					Debug.WriteLine(string.Format("DuplicateTokenEx result: {0}", result));
					Debug.WriteLine(string.Format("duplicate: {0}", duplicate));

					if (result)
					{
						result = CreateProcessAsUser(duplicate, processExe, null,
							IntPtr.Zero, IntPtr.Zero, false, (UInt32)CreateProcessFlags.CREATE_NEW_CONSOLE, IntPtr.Zero, null,
							ref info, out procInfo);
						Debug.WriteLine(string.Format("CreateProcessAsUser result: {0}", result));

					}

					if (p.ToInt32() != 0)
					{
						Marshal.Release(p);
						Debug.WriteLine(string.Format("Released handle p: {0}", p));
					}

					if (duplicate.ToInt32() != 0)
					{
						Marshal.Release(duplicate);
						Debug.WriteLine(string.Format("Released handle duplicate: {0}", duplicate));
					}

					return result;
				}

				public static int GetCurrentSessionId()
				{
					uint sessionId = WTSGetActiveConsoleSessionId();
					Debug.WriteLine(string.Format("sessionId: {0}", sessionId));

					if (sessionId == 0xFFFFFFFF)
						return -1;
					else
						return (int)sessionId;
				}

				public static bool IsUserLoggedOn()
				{
					List<WTS_SESSION_INFO> wtsSessionInfos = ListSessions();
					Debug.WriteLine(string.Format("Number of sessions: {0}", wtsSessionInfos.Count));
					return wtsSessionInfos.Where(x => x.State == ConnectionState.Active).Count() > 0;
				}

				private static IntPtr GetCurrentUserToken()
				{
					List<WTS_SESSION_INFO> wtsSessionInfos = ListSessions();
					int sessionId = wtsSessionInfos.Where(x => x.State == ConnectionState.Active).FirstOrDefault().SessionID;
					//int sessionId = GetCurrentSessionId();

					Debug.WriteLine(string.Format("sessionId: {0}", sessionId));
					if (sessionId == int.MaxValue)
					{
						return IntPtr.Zero;
					}
					else
					{
						IntPtr p = new IntPtr();
						int result = WTSQueryUserToken((UInt32)sessionId, out p);
						Debug.WriteLine(string.Format("WTSQueryUserToken result: {0}", result));
						Debug.WriteLine(string.Format("WTSQueryUserToken p: {0}", p));

						return p;
					}
				}

				public static List<WTS_SESSION_INFO> ListSessions()
				{
					IntPtr server = IntPtr.Zero;
					List<WTS_SESSION_INFO> ret = new List<WTS_SESSION_INFO>();

					try
					{
						IntPtr ppSessionInfo = IntPtr.Zero;

						Int32 count = 0;
						Int32 retval = WTSEnumerateSessions(IntPtr.Zero, 0, 1, ref ppSessionInfo, ref count);
						Int32 dataSize = Marshal.SizeOf(typeof(WTS_SESSION_INFO));

						Int64 current = (int)ppSessionInfo;

						if (retval != 0)
						{
							for (int i = 0; i < count; i++)
							{
								WTS_SESSION_INFO si = (WTS_SESSION_INFO)Marshal.PtrToStructure((System.IntPtr)current, typeof(WTS_SESSION_INFO));
								current += dataSize;

								ret.Add(si);
							}

							WTSFreeMemory(ppSessionInfo);
						}
					}
					catch (Exception exception)
					{
						Debug.WriteLine(exception.ToString());
					}

					return ret;
				}

			}
'@

		Add-Type -TypeDefinition $createProcess3 -Language CSharp -IgnoreWarnings;

	}


