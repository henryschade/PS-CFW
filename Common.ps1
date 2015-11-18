###########################################
# Updated Date:	18 November 2015
# Purpose:		Common routines to all/most projects.
# Requirements: Documents.ps1 for the CreateZipFile() routine.
##########################################

	##How to include/use this file in other projects:
	##Include following Scripts/Files.
	#$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;
	#. ($ScriptDir + "\Common.ps1")

	function CleanDir{
		#Cleans files out of directories based on the DateLastModified.  
		#Checks the "NumDays2KeepLogs" entry in MiscSettings.txt file, if $HowOld is -2, blank, or null.
		#   (180 days) if error reading NumDays2KeepLogs.
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$Directory, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$DoSubs = $False, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$TypesToSkip = "", 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Int]$HowOld = -2
		)
		#$Directory = Folder/Directory path to clean.  i.e. "C:\SRM_Apps_N_Tools" or "\\Nawesdnifs08.nadsuswe.nads.navy.mil\NMCIISF\NMCIISF-SDCP-MAC\MAC\Boise SRM Team\Jay_Nielson\Reports"
		#$DoSubs = True/False. (defult = False) Check/Clean sub folders too.
		#$TypesToSkip = file types NOT to delete/clean. 
		#	i.e. ".mdb" or ".ps1" or ".zip"
		#	Supports "!" (not) (as the first char).  i.e. "!.tmp" (it will only delete these file types).
		#		I want to make this support a ; seperated list of file types too.   i.e. ".mdb; .zip; .xlsx"
		#$HowOld = How many days old does the file need to be, to be deleted.

		$strSettingFile = "\\nawesdnifs08.nadsuswe.nads.navy.mil\NMCIISF\NMCIISF-SDCP-MAC\MAC\Entr_SRM\Support Files\MiscSettings.txt";

		if (($HowOld -lt -1) -or ($HowOld -eq "") -or ($HowOld -eq $null)){
			if ((Test-Path $strSettingFile)){
				$Error.Clear();
				foreach ($strLine in [System.IO.File]::ReadAllLines("\\nawesdnifs08.nadsuswe.nads.navy.mil\NMCIISF\NMCIISF-SDCP-MAC\MAC\Entr_SRM\Support Files\MiscSettings.txt")) {
					if ($strLine.StartsWith("--") -ne $True){
						if ($strLine.Contains("NumDays2KeepLogs")){
							$HowOld = $strLine.SubString($strLine.IndexOf("=") + 1).Trim();
							break;
						}
					}
				}
			}

			##http://stackoverflow.com/questions/10928030/in-powershell-how-can-i-test-if-a-variable-holds-a-numeric-value
			#Add-Type -Assembly Microsoft.VisualBasic;
			#if (!([Microsoft.VisualBasic.Information]::IsNumeric($HowOld))){
			#	$HowOld = 180;
			#}
			#I have since wrote a PS equivelent of the VB IsNumeric() routine.  (Part of Common.ps1)
			if ((IsNumeric $HowOld) -ne $True){
				$HowOld = 180;
			}
		}

		if (Test-Path $Directory){
			$objDirectory = Get-ChildItem -Path $Directory -ErrorAction SilentlyContinue -Force;			#force is necessary to get hidden files/folders;
			foreach ($objItem in $objDirectory){
				if ($objItem -ne $null){
					$objSubItem = Get-Item -LiteralPath $objItem.Fullname -Force -ErrorAction Stop;			#force is necessary to get hidden files/folders
					if ($objSubItem.PSIsContainer){
						#Directory
						if ($DoSubs -eq $True){
							CleanDir $objSubItem.Fullname $DoSubs $TypesToSkip $HowOld;

							$Error.Clear();
							$objTest = Get-ChildItem -Path $objSubItem.Fullname -ErrorAction SilentlyContinue -Force;			#force is necessary to get hidden files/folders;
							if (($objTest -eq $null) -and (!($Error))){
								Remove-Item $objSubItem.Fullname -Force;
							}
						}
					}else{
						#File
						#Write-Host $objSubItem.Fullname;
						#Write-Host $objSubItem.LastWriteTime;
						#Write-Host $objSubItem.CreationTime;
						#Write-Host [IO.File]::GetLastWriteTime($objItem.Fullname);
						$dteNow = [System.DateTime]::Now;
						#if (($objSubItem.LastWriteTime -lt ($dteNow.AddDays([Int]($HowOld * -1)))) -and ($objSubItem.CreationTime -lt ($dteNow.AddDays(-1)))){
						if ($objSubItem.LastWriteTime -lt ($dteNow.AddDays([Int]($HowOld * -1)))){
							if ($TypesToSkip -eq ""){
								#Delete it
								Remove-Item $objSubItem.Fullname -Force;
							}else{
								#See if the file is one to NOT delete.
								if ($TypesToSkip.Contains(";")){
									#multiple file types specified.  Still coding this feature.
								}else{
									if ($TypesToSkip.StartsWith("!")){
										if ($objSubItem.Name.EndsWith($TypesToSkip.SubString(1))){
											#Delete it
											Remove-Item $objSubItem.Fullname -Force;
										}
									}else{
										if ($objSubItem.Name.EndsWith($TypesToSkip) -eq $False){
											#Delete it
											Remove-Item $objSubItem.Fullname -Force;
										}
									}
								}
							}
						}
					}
				}
			}
		}
	}

	function ConvertUTCToLocal{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True, HelpMessage = "UTC / GMT time to convert to Local time.")][String]$strTime
		)

		return [System.TimeZone]::CurrentTimeZone.ToLocalTime($strTime);
	}

	function CreateZipFile{
		#Should use ZipCreateFile() in Documents.ps1.
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True, HelpMessage = "Zip file path to create.")][String]$ZipFile, 
			[ValidateNotNull()][Parameter(Mandatory=$True, HelpMessage = "Array of file path (full) to add.")][Array]$Files
		)
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters passed in.
			#$objReturn.Results		= $True or $False.  Was a zip file created.
			#$objReturn.Message		= "Success" or the error message.
			#$objReturn.Returns		= The full path annd file name of the file created.
		#$ZipFile = The zip file to create. (Full path) [i.e. "c:\path\file.zip"]
		#$Files = An array of the files to add to the zip file. (Full paths) [i.e. @("c:\path\file.one", "c:\path\file.two")]

		$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;
		. ($ScriptDir + "\Documents.ps1")

		$objReturn = ZipCreateFile $ZipFile $Files;

		return $objReturn;
	}

	function EnableDotNet4{
		#Checks if .NET 4 is enabled, and if NOT then creates the *.xml config file to enable .NET4 support.
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bISE2 = $False
		)
		#$bISE2 = $True or $False.  Create the "*\powershell_ise.exe.config" files along with the "*\powershell.exe.config" files.

		#Returns $True if created config files (at least tried to).
		#Returns $False if Config Files did NOT have to be created (.NET 4.x already enabled).

		$bReturn = $False;

		if ($PSVersionTable.CLRVersion.Major -lt 4){
			$bReturn = $True;

			#http://tfl09.blogspot.com/2010/08/using-newer-versions-of-net-with.html
			#http://w3facility.org/question/powershell-load-dll-got-error-add-type-could-not-load-file-or-assembly-webdriver-dll-or-one-of-its-dependencies-operation-is-not-supported/
			#http://www.adminarsenal.com/admin-arsenal-blog/powershell-running-net-4-with-powershell-v2
			#http://www.bonusbits.com/wiki/HowTo:Enable_.NET_4_Runtime_for_PowerShell_and_Other_Applications

			$strXML = "<?xml version=`"1.0`"?>`r`n<configuration>`r`n";
			$strXML = $strXML + "`t<startup useLegacyV2RuntimeActivationPolicy=`"true`">`r`n";
			$strXML = $strXML + "`t`t<supportedRuntime version=`"v4.0.30319`"/>`r`n`t`t<supportedRuntime version=`"v2.0.50727`"/>`r`n";
			$strXML = $strXML + "`t</startup>`r`n`t<runtime>`r`n";
			$strXML = $strXML + "`t`t<loadFromRemoteSources enabled=`"true`"/>`r`n";
			$strXML = $strXML + "`t</runtime>`r`n";
			$strXML = $strXML + "</configuration>`r`n";

			#Ideally we would use $pshome, but it is not always right, so doing both potential directories manually.
			$arrConfigFiles = @("C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe.config", "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe.config");
			if ($bISE2){
				$arrConfigFiles += "C:\Windows\System32\WindowsPowerShell\v1.0\powershell_ise.exe.config";
				$arrConfigFiles += "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell_ise.exe.config";
			}

			$Error.Clear();
			#The following may have issues creating/copying the files due to permissions.
			foreach ($strConfigFile in $arrConfigFiles){
				if (Test-Path -Path ($strConfigFile + ".bak")){
					#Have a backup copy already
					$strXML | Out-File ($strConfigFile);
				}else{
					#Don't have a backup file yet.
					if (Test-Path -Path ($strConfigFile)){
						#Copy the original config file to *.bak.
						Copy-Item -Path ($strConfigFile) -Destination ($strConfigFile + ".bak");
						if ((Test-Path -Path ($strConfigFile + ".bak"))){
							#Update/overwrite the config file
							$strXML | Out-File ($strConfigFile);
						}
					}else{
						#Config file does not exist, so create both
						$strXML | Out-File ($strConfigFile + ".bak");
						$strXML | Out-File ($strConfigFile);
					}
				}
			}

			if ($Error){
				$bReturn = $False;
			}
		}

		return $bReturn;
	}

	function GetFolderPathing{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$sName
		)
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters passed in.
			#$objReturn.Results		= $True or $False.  Was a path found/gotten.
			#$objReturn.Message		= "Success" or the error message.
			#$objReturn.Returns		= A DataTable of Path(s).
		#$sName = The name of the path(s) to get.

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

		#1) Look for a local config file.
		#2) Read the desired path out of the file.
		#3) If path not exist, query DB for path info.
		#4) Update local config file.
		#5) Return the path info.
		
		#Andrew wants to do step in this order instead:
		# 3, 1, 2, 4, 5

	}

	function GetUTC{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True, HelpMessage = "Local time to convert to UTC / GMT time.")][String]$strTime
		)

		return ((Get-Date $strTime).ToUniversalTime()).ToString();
	}

	function isADInstalled{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bEnable = $False, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bDisable = $False
		)
		#Check if have AD Installed and Enabled.
		#$bEnable = $True, $False.  Turn on the AD Features (that are part of the NMCI SRM default set) ONLY if RSAT is installed.
		#$bDisable = $True, $False.  Turn off the AD Features (that are NOT part of the NMCI SRM default set) ONLY if RSAT is installed.

		#https://technet.microsoft.com/en-us/library/ee449483(v=ws.10).aspx
			#Here are the settings from my system:
			#RemoteServerAdministrationTools-Features-Wsrm -- Disabled
			#RemoteServerAdministrationTools-Features-StorageManager -- Disabled
			#RemoteServerAdministrationTools-Features-StorageExplorer -- Disabled
			#RemoteServerAdministrationTools-Features-SmtpServer -- Disabled
			#RemoteServerAdministrationTools-Features-LoadBalancing -- Disabled
			#RemoteServerAdministrationTools-Features-GP -- Disabled
			#RemoteServerAdministrationTools-Features-Clustering -- Disabled
			#RemoteServerAdministrationTools-Features-BitLocker -- Disabled
			#RemoteServerAdministrationTools-Features -- Disabled
			#RemoteServerAdministrationTools-Roles-RDS -- Enabled
			#RemoteServerAdministrationTools-Roles-HyperV -- Disabled
			#RemoteServerAdministrationTools-Roles-FileServices-StorageMgmt -- Disabled
			#RemoteServerAdministrationTools-Roles-FileServices-Fsrm -- Disabled
			#RemoteServerAdministrationTools-Roles-FileServices-Dfs -- Disabled
			#RemoteServerAdministrationTools-Roles-FileServices -- Disabled
			#RemoteServerAdministrationTools-Roles-DNS -- Enabled
			#RemoteServerAdministrationTools-Roles-DHCP -- Disabled
			#RemoteServerAdministrationTools-Roles-AD-Powershell -- Enabled
			#RemoteServerAdministrationTools-Roles-AD-LDS -- Enabled
			#RemoteServerAdministrationTools-Roles-AD-DS-NIS -- Disabled
			#RemoteServerAdministrationTools-Roles-AD-DS-AdministrativeCenter -- Enabled
			#RemoteServerAdministrationTools-Roles-AD-DS-SnapIns -- Enabled
			#RemoteServerAdministrationTools-Roles-AD-DS -- Enabled
			#RemoteServerAdministrationTools-Roles-AD -- Enabled
			#RemoteServerAdministrationTools-Roles-CertificateServices-OnlineResponder -- Disabled
			#RemoteServerAdministrationTools-Roles-CertificateServices-CA -- Disabled
			#RemoteServerAdministrationTools-Roles-CertificateServices -- Disabled
			#RemoteServerAdministrationTools-Roles -- Enabled
			#RemoteServerAdministrationTools-ServerManager -- Disabled
			#RemoteServerAdministrationTools -- Enabled

		$bInstalled = $False;

		#To get a list of all "RemoteServerAdministrationTools" and if they are enabled or disabled:
		#[System.Collections.ArrayList]$arrResults = DISM /online /get-features | Select-String -Pattern ":";
		$arrResults = DISM /online /get-features | Select-String -Pattern ":";
		#if (($arrResults.GetType().Name -eq "ArrayList") -or ($arrResults.GetType().BaseType.Name -eq "Array")){
		#if (($arrResults.GetType().Name -eq "ArrayList") -or ($arrResults.GetType().IsArray) -or ($arrResults.GetType().BaseType.Name -eq "Array")){
		if (($arrResults.GetType().IsArray) -or ($arrResults.GetType().BaseType.Name -eq "Array")){
			[System.Collections.ArrayList]$arrResults = $arrResults;
		}else{
			[System.Collections.ArrayList]$arrResults = @($arrResults.ToString());
		}

		$arrFiltered = @();
		for ($intX = $arrResults.Count; $intX -ge 0; $intX--){
			if ($arrResults[$intX] -Match "RemoteServerAdministrationTools"){
				$strEntry = $arrResults[$intX].ToString();
				$strEntry = ($strEntry.Replace("Feature Name :", "")).Trim();
				$strEnabled = $arrResults[$intX+1].ToString();
				$strEnabled = ($strEnabled.Replace("State :", "")).Trim();
				#Write-Host "Entry-- " $strEntry;
				#Write-Host "Enabled-- " $strEnabled;
				#Write-Host $strEntry " -- " $strEnabled;
				$arrFiltered += $strEntry + " -- " + $strEnabled;
			}
		}

		if (($arrFiltered.Count -eq 0) -or ($arrFiltered.Count -eq $null)){
			#NOT installed.
			$bInstalled = $False;
			if (($arrResults -match "Error:") -ne ""){
				#To catch the "Error: 740" error if no permissions to "read" the installed "Windows Features".
				#Should translate to:    "Error(740): The requested operation requires elevation."
				#Write-Host "Error"
				$bInstalled = $arrResults;
			}
		}else{
			#Installed
			if ((($arrFiltered -Match "RemoteServerAdministrationTools-Roles-AD-Powershell -- Enabled").Count -eq 0) -or (($arrFiltered -Match "RemoteServerAdministrationTools-Roles-AD -- Enabled").Count -eq 0)){
				#AD Checkboxes are NOT Checked.
				$bInstalled = $False;
			}else{
				#AD Checkboxes are Checked.
				$bInstalled = $True;
			}

			#if (($bInstalled -eq $False) -and ($bEnable)){
			if ($bEnable){
				#Turn ON these features
				$strResults = DISM /online /enable-feature /featurename:RemoteServerAdministrationTools /featurename:RemoteServerAdministrationTools-Roles /featurename:RemoteServerAdministrationTools-Roles-AD /featurename:RemoteServerAdministrationTools-Roles-AD-DS /featurename:RemoteServerAdministrationTools-Roles-AD-DS-SnapIns /featurename:RemoteServerAdministrationTools-Roles-AD-DS-AdministrativeCenter /featurename:RemoteServerAdministrationTools-Roles-AD-LDS /featurename:RemoteServerAdministrationTools-Roles-AD-Powershell /featurename:RemoteServerAdministrationTools-Roles-DNS /featurename:RemoteServerAdministrationTools-Roles-RDS;
				if ([String]($strResults -Match "completed successfully") -Like "*successfully*"){
					$bInstalled = $True;
				}
			}

			if ($bDisable){
				#Turn OFF these features
				$strResults = DISM /online /disable-feature /featurename:RemoteServerAdministrationTools-Features /featurename:RemoteServerAdministrationTools-Features-BitLocker /featurename:RemoteServerAdministrationTools-Features-Clustering /featurename:RemoteServerAdministrationTools-Features-GP /featurename:RemoteServerAdministrationTools-Features-LoadBalancing /featurename:RemoteServerAdministrationTools-Features-SmtpServer /featurename:RemoteServerAdministrationTools-Features-StorageExplorer /featurename:RemoteServerAdministrationTools-Features-StorageManager /featurename:RemoteServerAdministrationTools-Features-Wsrm /featurename:RemoteServerAdministrationTools-Roles-AD-DS-NIS /featurename:RemoteServerAdministrationTools-Roles-DHCP /featurename:RemoteServerAdministrationTools-Roles-HyperV /featurename:RemoteServerAdministrationTools-Roles-FileServices /featurename:RemoteServerAdministrationTools-Roles-FileServices-Dfs /featurename:RemoteServerAdministrationTools-Roles-FileServices-Fsrm /featurename:RemoteServerAdministrationTools-Roles-FileServices-StorageMgmt /featurename:RemoteServerAdministrationTools-Roles-CertificateServices /featurename:RemoteServerAdministrationTools-Roles-CertificateServices-CA /featurename:RemoteServerAdministrationTools-Roles-CertificateServices-OnlineResponder /featurename:RemoteServerAdministrationTools-ServerManager;
			}
		}
		return $bInstalled;
	}

	function isNumeric($intX){
		#IsNumeric() equivelant is -> [Boolean]([String]($x -as [int]))

		#http://rosettacode.org/wiki/Determine_if_a_string_is_numeric
		try {
			0 + $intX | Out-Null;
			return $true;
		} catch {
			return $false;
		}
	}

	function WriteLogFile{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$Message, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$LogDir, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$LogFile, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$Header = ""
		)
		#Uses Out-File to append $Message to the $LogFile, in the path $LogDir.
		#$LogFile get updated to a format of "yyyymmdd_"$LogFile.  (i.e. 20150513_AscII.log)
		#$Message gets PrePended with a "Header":
		#Default Header is (but a different/custom one can be provided INSTEAD) ("False" for no header at all. [NOT boolean]):
			#Date Time - Domain\User - MachineName (MAC) - IP - Ticket# -- $Message
			#i.e.:
			#5/13/2015 9:23:15 - NMCI-ISF\henry.schade - ADIDBO226572 (00:24:81:21:CA:CC) - 10.12.21.80 - 8989765 -- $Message

		#Make sure the log directory exist.
		if (!(Test-Path -Path $LogDir)){
			#Need to create the directory
			#PS mkdir, will create any parent folders needed.
			$strResults = mkdir $LogDir;
		}

		if (($Message.Trim() -ne "`r`n") -and ($Message.Trim() -ne "")){
			#$strDateCode = (Get-Date -format "yyyy") + (Get-Date -format "MM") + (Get-Date -format "dd");			# + "." + (Get-Date -format "HH") + (Get-Date -format "mm");
			$strDateCode = (Get-Date).ToString("yyyyMMdd");
			$LogFile = $strDateCode + "_" + $LogFile;
			$Message = $Message.Trim().Replace("`r`n", "   ");

			if (($Header -eq "") -or ($Header -eq $null)){
				if (($txbTicketNum -ne $null)){
					if (($txbTicketNum.Text -eq "") -or ($txbTicketNum.Text -eq $null)){
						$strTicketNum = "none";
					}else{
						$strTicketNum = $txbTicketNum.Text;
					}
				}else{
					$strTicketNum = "none";
				}
				$strLogHeader = (([System.DateTime]::Now).ToString() + " - " + ([System.Security.Principal.WindowsIdentity]::GetCurrent()).Name + " - " + $env:computername + " (" + (Get-WmiObject Win32_NetworkAdapterConfiguration -Namespace "root\CIMV2" | WHERE{$_.IPEnabled -eq "True"}).MACAddress + ") - " + (Get-WmiObject Win32_NetworkAdapterConfiguration -Namespace "root\CIMV2" | WHERE{$_.IPEnabled -eq "True"}).IPAddress + " - " + $strTicketNum + " -- ");
				$Message = $strLogHeader + $Message;
			}else{
				if ($Header -ne "False"){
					$Message = $Header + $Message;
				}
			}

			#Write to log file
			$Message | Out-File -filepath ($LogDir + $LogFile) -Encoding Default -Append;
		}
	}

