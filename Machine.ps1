###########################################
# Updated Date:	19 April 2016
# Purpose:		Routines that require a Computer, or that interact w/ a Computer.
# Requirements: None
##########################################

	function CheckIfOnline{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strComp, 
			[ValidateNotNull()][Parameter(Mandatory=$False)]$bolUsers = $True
		)
		#$strComp = The computer to check.  ($Env:ComputerName)

		#if ([String]::IsNullOrWhiteSpace($strComp)){
		#	#$strComp = "ALSDCP002656";		#Henry Laptop;
		#	#$strComp = "ALSDNI390014";		#Andrew Laptop;
		#	$strComp = Read-Host 'What computer? (i.e. ALSDCP002656)';
		#}

		$ErrorActionPreference = 'SilentlyContinue';
		$strIP = [System.Net.DNS]::GetHostAddresses($strComp);
		$ErrorActionPreference = 'Continue';
		#Write-Host $strIP;

		if ([String]::IsNullOrWhiteSpace($strIP)){
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

	function DoShutDown($strComp){
		if ([String]::IsNullOrWhiteSpace($strComp)){
			#$strComp = "ALSDCP002656";		#Henry Laptop;
			$strComp = Read-Host 'What computer? (i.e. ALSDCP002656)';
		}

		#$strRet = Stop-Computer -comp $strComp -force;
		#$strRet = Stop-Computer -comp $strComp -force -Credential $objCreds;
		$objCreds = GetCredentials;
		$strRet = Stop-Computer -comp $strComp -force -Credential $objCreds;

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
		if([String]::IsNullOrWhiteSpace($strComp)){
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
		if([String]::IsNullOrWhiteSpace($objCreds)){
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
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strProp
		)
		#Returns an array.
			#A "list" of Properties & "Directories" (Keys) under the "Path" provided.
			#Data from the Property in the format of: (Name, Type/Kind, Value).
		#$strRegPath = The Registry Path (Key) to get data from.  (i.e. "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion")
		#$strProp = The Property, if any, to get the data out of.  If omitted returns a list of available Properties & Keys.  (i.e. "DevicePath")

		#registry key = The Reg Path.
		#registry key property = The "attributes" / "fields" under the Key.

		#http://powershell.com/cs/blogs/tips/archive/2015/04/15/getting-registry-values-and-value-types.aspx
		#https://4sysops.com/archives/interacting-with-the-registry-in-powershell/
		#http://stackoverflow.com/questions/15511809/how-do-i-get-the-value-of-a-registry-key-and-only-the-value-using-powershell
		#Remote Registry:
		#http://blogs.microsoft.co.il/scriptfanatic/2010/01/10/remote-registry-powershell-module/
		#https://psremoteregistry.codeplex.com/


		#$strRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion";
		#$strProp = "DevicePath";

		if (Test-Path $strRegPath){
			if ([String]::IsNullOrWhiteSpace($strProp)){
				if ($strRegPath.EndsWith("\")){
					$strRegPath = $strRegPath.SubString(0, $strRegPath.Length - 1);
				}

				$arrRet = (Get-Item -Path $strRegPath -Force -ErrorAction SilentlyContinue).Property;
				#$arrRet now has all the Keys/Properties under $strRegPath, BUT NOT any "child" Pathing.
				#$arrRet is an array.

				$objTemp = ((Get-ChildItem -Path $strRegPath -Force -ErrorAction SilentlyContinue) | SELECT Name);
				for ($intX = 0; $intX -lt $objTemp.Count; $intX++){
					$objTemp[$intX] = $objTemp[$intX].Name;
					$objTemp[$intX] = $objTemp[$intX].Replace($strRegPath.Replace("HKLM:", "HKEY_LOCAL_MACHINE"), "");
				}

				$arrRet = $arrRet + $objTemp;
			}
			else{
				$bolExist = $False;
				$arrInfo = GetRegistry $strRegPath;
				foreach ($strEntry in $arrInfo){
					if ($strEntry -eq $strProp){
						$bolExist = $True;
						break;
					}
				}
				if ($bolExist -eq $True){
					$arrRet = @($strProp, (Get-Item -Path $strRegPath -Force -ErrorAction SilentlyContinue).GetValueKind($strProp), (Get-Item -Path $strRegPath -Force -ErrorAction SilentlyContinue).GetValue($strProp));
					#$arrRet now has the Value and Type stored in/at $strProp.
					#Although RegEdit shows "%SystemRoot%\inf", and the value returned is "C:\Windows\inf".
					#$arrRet is an array of (Name, Type/Kind, Value).
				}
				else{
					$arrRet = @("The Property '$strProp' does NOT exist under Key '$strRegPath'.", "-1");
				}
			}
		}
		else{
			$arrRet = @("The Key '$strRegPath' does NOT exist.", "-1");
		}

		return $arrRet;
	}

	function LoggedInUser{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strComp
		)
		#$strComp = The computer to check.  ($Env:ComputerName)

		#if([String]::IsNullOrWhiteSpace($strComp)){
		#	#$strComp = "ALSDCP002656";		#Henry Laptop;
		#	$strComp = Read-Host 'What computer? (i.e. ALSDCP002656)';
		#}

		#check if a user is logged in
		$Error.Clear();
		$strLoggedIn = Get-WmiObject -class win32_computerSystem -computer:$strComp | Select-Object username;
		if ($strLoggedIn.username.Length -gt 0){
			$strRet = $strLoggedIn.username + " is currently logged in to " + $strComp + ".";
		}else{
			$strRet = "No user is logged in locally to " + $strComp + ".";
		}

		#Write-Host $strRet;
		return $strRet;
	}

	function SetRegistry{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strRegPath, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strProp, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strValue, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strType = "String"
		)
		#Returns an array with results info.
		#$strRegPath = The Registry Path (Key) to set data to, or to create.  (i.e. "HKLM:\SOFTWARE\ITSS-Tools\ASCII")
		#$strProp = The Registry Key Property, if any, to update/create with $strValue. (i.e. "Version").
		#$strValue = The Value to enter.  (i.e. "0.23")
		#$strType = The Type/Kind that strValue is to be stored as.
			#DWord = Integer
			#String = String

		#registry key = The Reg Path.
		#registry key property = The "attributes" / "fields" under the Key.

		#Remote Registry:
		#http://blogs.microsoft.co.il/scriptfanatic/2010/01/10/remote-registry-powershell-module/

		#$strRegPath = "HKLM:\SOFTWARE\ITSS-Tools\ASCII";
		#$strProp = "Version";
		#$strValue = "0.23";

		$arrRet = $null;

		$Error.Clear();
		if (Test-Path $strRegPath){
			#Path exists, set/update Key.
		}
		else{
			#Create/update a new Path.  Creates any neccessary parent pathing.
			$objResults = $null;
			#$strResults = New-Item -Path $strRegPath -Force | Out-Null;
			$objResults = New-Item -Path $strRegPath -Force -ErrorAction SilentlyContinue;
			if (($Error) -or ([String]::IsNullOrWhiteSpace($objResults))){
				$arrRet = @("Error creating the path '$strRegPath'.", $Error);
			}
			else{
				$arrRet = @("Created Key '$strRegPath'.", $objResults);
			}
		}

		if ((!([String]::IsNullOrWhiteSpace($strProp))) -and (!($Error))){
			#Now we can create/update the Key.
			$objResults = $null;
			$arrInfo = GetRegistry $strRegPath $strProp;
			$Error.Clear();
			if ($arrInfo[0].Contains("does NOT exist")){
				#$objResults = New-ItemProperty -Path $strRegPath -Name $strProp -Value $strValue -PropertyType $strType -Force | Out-Null;
				$objResults = New-ItemProperty -Path $strRegPath -Name $strProp -Value $strValue -PropertyType $strType -Force -ErrorAction SilentlyContinue;
				if (($Error) -or ([String]::IsNullOrWhiteSpace($objResults))){
					$arrRet = @("Error creating Property '$strProp', under Key '$strRegPath'.", $Error);
				}
				else{
					$arrInfo = GetRegistry $strRegPath $strProp;
					#$strMessage = "Created Property '$arrInfo[0]', under Key '$strRegPath', with a Value of '$arrInfo[2]' (of type '$arrInfo[1]').";
					$strMessage = "Created Property '" + $arrInfo[0] + "', under Key '$strRegPath', with a Value of '" + $arrInfo[2] + "'";
					if ($strValue -ne $arrInfo[2]){
						$strMessage = $strMessage + " (NOT the provided: '$strValue')";
					}
					$strMessage = $strMessage + " (of type '" + $arrInfo[1] + "').";

					if ([String]::IsNullOrWhiteSpace($arrRet)){
						#$arrRet = @("Created Property '$strProp', under Key '$strRegPath', with a Value of '$strValue' (of type '$strType').", $objResults);
						$arrRet = @($strMessage, $objResults);
					}
					else{
						#$arrRet[0] = $arrRet[0] + "`r`n" + "Created Property '$strProp', under Key '$strRegPath', with a Value of '$strValue' (of type '$strType').";
						$arrRet[0] = $arrRet[0] + "`r`n" + $strMessage;
						$arrRet = $arrRet + @($objResults)
					}
				}
			}
			else{
				$objResults = Set-ItemProperty -Path $strRegPath -Name $strProp -Value $strValue -Type $strType -ErrorAction SilentlyContinue;
				if (($Error) -or (!([String]::IsNullOrWhiteSpace($objResults)))){
					$arrRet = @("Error updating Property '$strProp', under Key '$strRegPath'.", $Error);
				}
				else{
					$arrInfo = GetRegistry $strRegPath $strProp;
					#$strMessage = "Created Property '$arrInfo[0]', under Key '$strRegPath', with a Value of '$arrInfo[2]' (of type '$arrInfo[1]').";
					$strMessage = "Updated Property '" + $arrInfo[0] + "', under Key '$strRegPath', with a Value of '" + $arrInfo[2] + "'";
					if ($strValue -ne $arrInfo[2]){
						$strMessage = $strMessage + " (NOT the provided: '$strValue')";
					}
					$strMessage = $strMessage + " (of type '" + $arrInfo[1] + "').";

					#$arrRet = @("Updated Property '$strProp', under Key '$strRegPath', with a Value of '$strValue' (of type '$strType').", $objResults);
					$arrRet = @($strMessage , $arrInfo);
				}
			}
		}

		return $arrRet;
	}
