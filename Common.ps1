###########################################
# Updated Date:	3 February 2017
# Purpose:		Common routines to all/most projects.
# Requirements: Core.ps1 [will try to load it automatically].
#				DB-Routines.ps1 for the CheckVer() routine [it will try to load it automatically].
#				...\PS-CFW\MiscSettings.txt.
##########################################

<# ---=== Change Log ===---
	#Changes for 28 June 2016:
		#Added Change Log.
		#Fixed bug with UpdateLocalFiles(), make sure share path exists.
	#Changes for 4 Oct 2016
		#Add CDR SIPR info to GetPathing().
		#Add ECMD SIPR info to GetPathing().
		#Update MiscSettings.txt default path to be CFW instead of SupportFiles.
		#Added GetCurrentFiles() back in from the 20160614b backup copy, and updated it.
	#Changes for 14 Oct 2016
		#fixed a bug in SaveConfig().  Would only save one config setting.
	#Changes for 24 Oct 2016
		#Bug fix in BackUpDir().  If $strBackUpDir was provided did not check to make sure directory existed.
	#Changes for 27 Oct 2016
		#Improve the EnableDotNet4() message about running/restarting as admin.
	#Changes for 10 Nov 2016
		#Update isADInstalled() to better check Servers for AD installed and enabled.
	#Changes for 21 Nov 2016
		#Update LoadRequired() to have a progress bar, so one can tell if things are still running.
	#Changes for 5 December 2016
		#Remove routines from Common.ps1, and create Core.ps1 with them.
	#Changes for 8 December 2016
		#Add "#Returns: " to functions, for routine documentation.
		#Added documentation to isNumeric().
	#Changes for 15 December 2016
		#Updated isADInstalled() to help accomodate Servers.
	#Changes for 26 January 2017
		#Updated $ScriptDir entries to $ScriptDirectory
	#Changes for 27 January 2017
		#Moved isNumeric() to Core.ps1 as it makes use of it too.
	#Changes for 30 January 2017
		#Added GetChangeLog() routine.
	#Changes for 3 February 2017
		#Fixed bug with checking for Core routines.
		#Added CopyUpLocalLogs().
#>


	#$global:LoadedFiles that CheckVer() uses is in Core.ps1.

	#Make sure the Core routines in Core.ps1 are loaded.
	if ((!(Get-Command "EnableDotNet4" -ErrorAction SilentlyContinue)) -or ((!(Get-Command "LoadRequired" -ErrorAction SilentlyContinue)) -or (!(Get-Command "isNumeric" -ErrorAction SilentlyContinue)))){
		if ([String]::IsNullOrEmpty($MyInvocation.MyCommand.Path)){
			$ScriptDirectory = (Get-Location).ToString();
		}
		else{
			$ScriptDirectory = Split-Path $MyInvocation.MyCommand.Path;
		}

		if ((Test-Path (".\Core.ps1"))){
			. (".\Core.ps1");
		}
		else{
			if ((Test-Path (".\..\PS-CFW\Core.ps1"))){
				. (".\..\PS-CFW\Core.ps1");
			}
		}
	}

	function CheckVer{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$Project, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$RunningVer
		)
		#Checks the running version of $Project against the posted Production version.
		#Can also checks that the files in $global:LoadedFiles are up to date.
		#Returns: a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters passed in.
			#$objReturn.Results		= $True or $False.  Running correct Production\Beta version.
			#$objReturn.Message		= "Success", "Disable", or the error message.
			#$objReturn.Returns		= An array of the Production version number and the Beta version number.  i.e. @("2.2", "2.5b")
		#$Project = The Project name to check.  (i.e. "WILE", "ASCII", etc)
		#$RunningVer = The version currently being run, Beta versions MUST end in "b", "B", "Beta", "beta".
			#To Update/Check the Required/Included files, need to dot source this function and set $Project = "Includes" and $RunningVer = "".
			#i.e.   $objRet = (. CheckVer "Includes" "");

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
			Returns = "0.0";
		}

		if (($Project -eq "Includes") -and ($RunningVer -eq "")){
			#Check if the Required files have been updated since initial load.
			$objReturn.Results = $True;
			$objReturn.Message = "Success";

			$Error.Clear();
			foreach ($strFile in $global:LoadedFiles.Keys){
				$strPath = $global:LoadedFiles.$strFile.Path;
				$objFile = Get-Item -LiteralPath ($strPath + $strFile);
				$Date = $objFile.LastWriteTime;

				#if ((((Get-Date $objFile.LastWriteTime).ToUniversalTime()).ToString()) -gt $global:LoadedFiles.$strFile.Date){
				if ((((Get-Date $Date).ToUniversalTime()).ToString()) -gt $global:LoadedFiles.$strFile.Date){
					#The included file has been updated sincel last loaded.
					#Reload the file
					. ($global:LoadedFiles.$strFile.Path + $strFile);

					#Update the $global:LoadedFiles entry
					$strDateVer = (((Get-Date $Date).ToString("yyyyMMdd.hhmmss")));
					$Date = (((Get-Date $Date).ToUniversalTime()).ToString());
					#$global:LoadedFiles.($objFile.Name) = (@{"Ver" = $strDateVer; "Date" = $Date; "Path" = ($objFile.FullName).Replace(($objFile.Name), "")});
					$global:LoadedFiles.($objFile.Name) = (@{"Ver" = $strDateVer; "Date" = $Date; "Path" = $strPath});
				}
			}
			if ($Error){
				$objReturn.Message = "Error `r`n" + $Error;
			}
		}
		else{
			#Make sure the DB routines that are in DB-Routines.ps1 are loaded.
			if ((!(Get-Command "GetDBInfo" -ErrorAction SilentlyContinue)) -or (!(Get-Command "QueryDB" -ErrorAction SilentlyContinue))){
				if ([String]::IsNullOrEmpty($MyInvocation.MyCommand.Path)){
					$ScriptDirectory = (Get-Location).ToString();
				}else{
					$ScriptDirectory = Split-Path $MyInvocation.MyCommand.Path;
				}
				if ((Test-Path (".\DB-Routines.ps1"))){
					. (".\DB-Routines.ps1");
				}
				else{
					if ((Test-Path (".\..\PS-CFW\DB-Routines.ps1"))){
						. (".\..\PS-CFW\DB-Routines.ps1");
					}
				}
			}

			#Query DB.
			#$arrDBInfo = GetDBInfo "AgentActivity";
			#$strSQL = "GetSP_spGetNetPath '" + $Project + "';";
			#$Project = "CA";		#TEAC
			$arrDBInfo = GetDBInfo "SRMDB";
			$strSQL = "";
			#$strSQL = $strSQL + "SELECT * FROM AppChanges ";
			#$strSQL = $strSQL + "INNER JOIN AppInfo on AppChanges.AppInitials = AppInfo.AppInitials ";
			#$strSQL = $strSQL + "INNER JOIN AppReference on AppReference.AppInitials = AppInfo.AppInitials ";
			#$strSQL = $strSQL + "WHERE (AppChanges.AppInitials = '" + $Project + "')";
			$strSQL = $strSQL + "SELECT Ver_Num_P, Ver_Num_B, Allow_Old_Ver ";
			$strSQL = $strSQL + "FROM ((SourceDesc ";
			$strSQL = $strSQL + "LEFT JOIN SourceChanges ON SourceDesc.GUID = SourceChanges.SourceDesc_GUID) ";
			$strSQL = $strSQL + "LEFT JOIN SourceFiles ON SourceDesc.GUID = SourceFiles.SourceDesc_GUID) ";
			$strSQL = $strSQL + "LEFT JOIN SourceUses ON SourceDesc.GUID = SourceUses.SourceDesc_GUID ";
			$strSQL = $strSQL + "WHERE ((App_Name = '" + $Project + "') OR (App_Name_Short = '" + $Project + "'));";
			$objTable = $null;
			$Error.Clear();
			#$objTable = QueryDB $arrDBInfo[1] $arrDBInfo[2] $strSQL $True;
			$objTable = QueryDB $arrDBInfo[1] $arrDBInfo[2] $strSQL $False $arrDBInfo[3] $arrDBInfo[4];

			if (!(($objTable.Rows[0].Message -eq "Error") -or ($Error) -or ($objTable -eq $null) -or ($objTable.Rows.Count -eq 0))){
				$objReturn.Message = "Success";
				#Return the Prod and Beta ver #'s
				#$objReturn.Returns = $objTable.Rows[0].UpdatedDate;
				$strBDVerP = [String]$objTable.Rows[0].Ver_Num_P;
				$strBDVerB = [String]$objTable.Rows[0].Ver_Num_B;
				#$objReturn.Returns = @($objTable.Rows[0].Ver_Num_P, $objTable.Rows[0].Ver_Num_B);
				$objReturn.Returns = @($strBDVerP, $strBDVerB);

				if (($strBDVerP -gt $strBDVerB) -and ($strBDVerP -ne $RunningVer)){
					#Not running production, and Production ver is greater than Beta ver.
					$objReturn.Message = "Disable";
				}
				else{
					#Allow old version to run?
					#if (($objTable.Rows[0].DisableOld -eq "yes") -or ($objTable.Rows[0].DisableOld -eq $True)){
					if (($objTable.Rows[0].Allow_Old_Ver -eq "no") -or ($objTable.Rows[0].Allow_Old_Ver -eq $False) -or ($objTable.Rows[0].Allow_Old_Ver -eq 0)){
						#NO old versions allowed.
						if (($RunningVer.EndsWith("B")) -or ($RunningVer.EndsWith("b")) -or ($RunningVer.EndsWith("Beta")) -or ($RunningVer.EndsWith("beta"))){
							#$strBDVer = [String]$objTable.Rows[0].Ver_Num_B;
							if (($strBDVerB -ne $RunningVer) -or ($strBDVerB -eq $null) -or ($strBDVerB -eq "")){
								$objReturn.Message = "Disable";
							}
							else{
								$objReturn.Results = $True;
							}
						}
						else{
							#$strBDVer = [String]$objTable.Rows[0].Ver_Num_P;
							if ($strBDVerP -ne $RunningVer){
								$objReturn.Message = "Disable";
							}
							else{
								$objReturn.Results = $True;
							}
						}
					}
					else{
						#Old versions are allowed.
						#Are they running the latest Production version?
						if (($RunningVer.EndsWith("B")) -or ($RunningVer.EndsWith("b")) -or ($RunningVer.EndsWith("Beta")) -or ($RunningVer.EndsWith("beta"))){
							#$strBDVer = [String]$objTable.Rows[0].Ver_Num_B;
							if (($strBDVerB -eq $RunningVer)){
								#$objReturn.Message = "Disable";
								$objReturn.Results = $True;
							}
						}
						else{
							#$strBDVer = [String]$objTable.Rows[0].Ver_Num_P;
							if ($strBDVerP -eq $RunningVer){
								#$objReturn.Message = "Disable";
								$objReturn.Results = $True;
							}
						}
					}
				}
			}
			else{
				$objReturn.Results = $True;
				if ($Error){
					$strMessage = "Error getting version info.`r`n" + $Error;
					$strMessage = "`r`n" + ("-" * 100) + "`r`n" + $strMessage + "`r`n`r`n";
					$strMessage = $strMessage + $strSQL + "`r`n";
				}
				else{
					if (($objTable -eq $null) -or ($objTable.Rows.Count -eq 0)){
						$strMessage = "No Results getting version info.";
					}else{
						$strMessage = $objTable.Rows[0].Message + " getting version info.`r`n" + $objTable.Rows[0].Results;
					}
				}
				$objReturn.Message = "Error " + $strMessage;
			}
		}

		return $objReturn;
	}

	function CleanDir{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$Directory, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$DoSubs = $False, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$TypesToSkip = "", 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Int]$HowOld = -2
		)
		#Cleans files out of directories based on the DateLastModified.  
		#Checks the "NumDays2KeepLogs" entry in MiscSettings.txt file, if $HowOld is -2, blank, or null.
		#   (180 days) if error reading NumDays2KeepLogs.
		#Returns: 
		#$Directory = Folder/Directory path to clean.  i.e. "C:\SRM_Apps_N_Tools" or "\\Server.Name.FQDN\Path1\Path2\Path3"
		#$DoSubs = True/False. (defult = False) Check/Clean sub folders too.
		#$TypesToSkip = file types NOT to delete/clean. 
		#	i.e. ".mdb" or ".ps1" or ".zip"
		#	Supports "!" (not) (as the first char).  i.e. "!.tmp" (it will only delete these file types).
		#		I want to make this support a ; seperated list of file types too.   i.e. ".mdb; .zip; .xlsx"
		#$HowOld = How many days old does the file need to be, to be deleted.

		$strSettingFile = (GetPathing "SupportFiles").Returns.Rows[0]['Path'];
		$strSettingFile = $strSettingFile + "MiscSettings.txt";

		if (($HowOld -lt -1) -or ($HowOld -eq "") -or ($HowOld -eq $null)){
			if ((Test-Path $strSettingFile)){
				$Error.Clear();
				foreach ($strLine in [System.IO.File]::ReadAllLines($strSettingFile)) {
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

	function CreateZipFile{
		#Should use ZipCreateFile() in Documents.ps1.
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True, HelpMessage = "Zip file path to create.")][String]$ZipFile, 
			[ValidateNotNull()][Parameter(Mandatory=$True, HelpMessage = "Array of file path (full) to add.")][Array]$Files
		)
		#Should use ZipCreateFile() in Documents.ps1.
		#Returns: a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters passed in.
			#$objReturn.Results		= $True or $False.  Was a zip file created.
			#$objReturn.Message		= "Success" or the error message.
			#$objReturn.Returns		= The full path annd file name of the file created.
		#$ZipFile = The zip file to create. (Full path) [i.e. "c:\path\file.zip"]
		#$Files = An array of the files to add to the zip file. (Full paths) [i.e. @("c:\path\file.one", "c:\path\file.two")]

		if ([String]::IsNullOrEmpty($MyInvocation.MyCommand.Path)){
			$ScriptDirectory = (Get-Location).ToString();
		}else{
			$ScriptDirectory = Split-Path $MyInvocation.MyCommand.Path;
		}
		$strInclude = "Documents.ps1";
		if (Test-Path -Path ($ScriptDirectory + "\..\PS-CFW\" + $strInclude)){
			. ($ScriptDirectory + "\..\PS-CFW\" + $strInclude)
		}
		else{
			. ($ScriptDirectory + "\" + $strInclude)
		}

		$objReturn = ZipCreateFile $ZipFile $Files;

		return $objReturn;
	}

	function CopyUpLocalLogs{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strLogDirL, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strLogDirS, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strLogFile
		)
		#Copies local log files to the Share, appending to any existing.
		#Returns: 
		#$strLogDirL = Local log Directory path.  i.e. "C:\SRM_Apps_N_Tools\Logs\"
		#$strLogDirS = Share log Directory path.  i.e. "\\Server.Name.FQDN\Path1\Path2\Path3\Logs\"
		#$strLogFile = Name of the default log file.

		if (Test-Path -Path ($strLogDirL)){
			$strLogDir = $strLogDirL;
		}
		else{
			$strLogDir = $strLogDirS;
		}

		if ((Test-Path -Path $strLogDirS)){
			if (!($env:UserDomain.toLower().Contains("snmci-isf"))){
				$strMessage = "Moving local log file(s) to the File Server.";
				#UpdateResults $strMessage $True;
				WriteLogFile $strMessage $strLogDir $strLogFile;

				#[System.DateTime]::Now;
				if (Test-Path -Path ($strLogDirL)){
					$SubItems = Get-ChildItem -Path $strLogDirL -Force;
					if ($SubItems -ne $null){
						$strMessage = "Moving local log file(s) to the File Server.";
						#UpdateResults $strMessage $True;
						WriteLogFile $strMessage $strLogDir $strLogFile;

						foreach ($SubItem in $SubItems){
							if (($SubItem -ne "") -and ($SubItem -ne $null)){
								#Write-Host $SubItem;
								#Write-Host $SubItem.Fullname;
								$strFileDate = "";
								$strFileName = $SubItem.Fullname;
								if ($strFileName -Match $strLogFile){
									#$strTodaysFile = ($strLogDirL + (Get-Date).ToString("yyyyMMdd") + "_" + $strLogFile);
									$strFileDate = "";
									$strFileName = $SubItem.Name;
									$strFileDate = $strFileName.SubString(0, $strFileName.IndexOf("_"));
									if ($strFileDate -ne ""){
										$Error.Clear();
										Add-Content ($strLogDirS + $strFileName) (Get-Content ($strLogDirL + $strFileName) -ReadCount 0);
										if ((!($Error)) -and (Test-Path -Path ($strLogDirS + $strFileName))){
											$intX = 0;
											do {
												$intX++;
												$Error.Clear();
												$strResults = Remove-Item ($strLogDirL + $strFileName);
											} while (($Error) -and ($intX -lt 10))
										}
									}
								}
								else{
									if (($strFileName.EndsWith(".log")) -or ($strFileName.EndsWith(".txt"))){
										$strFileName = $SubItem.Name;
										$strFileDate = $strFileName.SubString(0, $strFileName.IndexOf("_"));
										if ($strFileDate -ne ""){
											Add-Content ($strLogDirS + $strFileName) (Get-Content ($strLogDirL + $strFileName) -ReadCount 0);
										}
									}
								}
							}
						}
					}
				}
				#[System.DateTime]::Now;
				#The above processed a 252 MB log file in about (I gave up after 30 min), the system was basically unusable.
				#So added the "-ReadCount 0" paramater, and reduced the log file to 131 MB; The above processed the log file in about 52 minutes  But at least the system did not basically lock up.
				#With the "-ReadCount 0" paramater, and reduced the log file to 10.5 MB; The above processed the log file in about 4 minutes  And the system still did not lock up.
			}
		}
	}

	function GetChangeLog{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strFile
		)
		#Gets the Change Log of the File/Project ($strFile) requested.
			#See the top of this file for an example of how the Change Log entries should be entered.
		#Returns: A String of the Change Log.
		#$strFile = The file (full path) to check/get the change log info of.

		$strLogFile = "";
		if (Test-Path -Path $strFile){
			$bReadingLog = $False;
			$arrFile = [System.IO.File]::ReadAllLines($strFile);
			for ($intX = 0; $intX -lt $arrFile.Count; $intX++){
				if (($bReadingLog -eq $False) -and ($arrFile[$intX] -Match "Change Log")){
					$bReadingLog = $True
				}
				if ($bReadingLog -eq $True){
					if (($arrFile[$intX] -Match "Change Log") -and (!($arrFile[$intX] -Match ($strFile.split("\")[-1].Replace(".ps1", "").Replace("PS-", ""))))){
						$strLogFile = $strLogFile + ($arrFile[$intX].Replace("Change Log", (($strFile.split("\")[-1].Replace(".ps1", "").Replace("PS-", "")) + " Change Log"))) + "`r`n";
					}
					else{
						$strLogFile = $strLogFile + $arrFile[$intX] + "`r`n";
					}
				}
				if (($bReadingLog -eq $True) -and (($arrFile[$intX] -Match "#>") -or ((($arrFile[$intX] -Match "function") -and ($arrFile[$intX] -Match "{"))))){
					$bReadingLog = $False;
					break;
				}
			}
		}

		return $strLogFile;
	}

	function GetCurrentFiles{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strLocalDir, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strProjName, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bolDoPrompts = $False
		)
		#A place holder.  Should be using Core.UpdateLocalFiles() instead of this one.
		#Returns: 

		$strResults = UpdateLocalFiles $strLocalDir $strProjName $bolDoPrompts;

		return $strResults;
	}

	function isADInstalled{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bEnable = $False, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bDisable = $False
		)
		#Check if have AD Installed and Enabled.
		#Returns: 
		#$bEnable = $True, $False.  Turn on the AD Features (that are part of the NMCI SRM default set) ONLY if RSAT is installed.
		#$bDisable = $True, $False.  Turn off the AD Features (that are NOT part of the NMCI SRM default set) ONLY if RSAT is installed.

		#Here are the settings from my system:
		<#
		#https://technet.microsoft.com/en-us/library/ee449483(v=ws.10).aspx
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
		#>

		$bInstalled = $False;

		#To get a list of all "RemoteServerAdministrationTools" and if they are enabled or disabled:
		#[System.Collections.ArrayList]$arrResults = DISM /online /get-features | Select-String -Pattern ":";
		$arrResults = DISM /online /get-features | Select-String -Pattern ":";
		if (($arrResults.GetType().IsArray) -or ($arrResults.GetType().BaseType.Name -eq "Array")){
			[System.Collections.ArrayList]$arrResults = $arrResults;
		}else{
			[System.Collections.ArrayList]$arrResults = @($arrResults.ToString());
		}

		$arrFiltered = @();
		for ($intX = $arrResults.Count; $intX -ge 0; $intX--){
			if (($arrResults[$intX] -Match "RemoteServerAdministrationTools") -or ($arrResults[$intX] -Match "DirectoryServices") -or ($arrResults[$intX] -Match "ActiveDirectory")){
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

		<# $arrFiltered results:
			#My Desktop:
				RemoteServerAdministrationTools-Features-Wsrm -- Disabled
				RemoteServerAdministrationTools-Features-StorageManager -- Disabled
				RemoteServerAdministrationTools-Features-StorageExplorer -- Disabled
				RemoteServerAdministrationTools-Features-SmtpServer -- Disabled
				RemoteServerAdministrationTools-Features-LoadBalancing -- Disabled
				RemoteServerAdministrationTools-Features-GP -- Disabled
				RemoteServerAdministrationTools-Features-Clustering -- Disabled
				RemoteServerAdministrationTools-Features-BitLocker -- Disabled
				RemoteServerAdministrationTools-Features -- Disabled
				RemoteServerAdministrationTools-Roles-RDS -- Enabled
				RemoteServerAdministrationTools-Roles-HyperV -- Disabled
				RemoteServerAdministrationTools-Roles-FileServices-StorageMgmt -- Disabled
				RemoteServerAdministrationTools-Roles-FileServices-Fsrm -- Disabled
				RemoteServerAdministrationTools-Roles-FileServices-Dfs -- Disabled
				RemoteServerAdministrationTools-Roles-FileServices -- Disabled
				RemoteServerAdministrationTools-Roles-DNS -- Enabled
				RemoteServerAdministrationTools-Roles-DHCP -- Disabled
				RemoteServerAdministrationTools-Roles-AD-Powershell -- Enabled
				RemoteServerAdministrationTools-Roles-AD-LDS -- Enabled
				RemoteServerAdministrationTools-Roles-AD-DS-NIS -- Disabled
				RemoteServerAdministrationTools-Roles-AD-DS-AdministrativeCenter -- Enabled
				RemoteServerAdministrationTools-Roles-AD-DS-SnapIns -- Enabled
				RemoteServerAdministrationTools-Roles-AD-DS -- Enabled
				RemoteServerAdministrationTools-Roles-AD -- Enabled
				RemoteServerAdministrationTools-Roles-CertificateServices-OnlineResponder -- Disabled
				RemoteServerAdministrationTools-Roles-CertificateServices-CA -- Disabled
				RemoteServerAdministrationTools-Roles-CertificateServices -- Disabled
				RemoteServerAdministrationTools-Roles -- Enabled
				RemoteServerAdministrationTools-ServerManager -- Disabled
				RemoteServerAdministrationTools -- Enabled

			#Server #1:
				DirectoryServices-ISM-Smtp -- Disabled
				DirectoryServices-ADAM-Tools -- Enabled
				DirectoryServices-AdministrativeCenter -- Enabled
				ActiveDirectory-PowerShell -- Enabled
				DirectoryServices-ADAM -- Disabled
				DirectoryServices-DomainController -- Disabled
				DirectoryServices-DomainController-Tools -- Enabled

			#Server #2:
				DirectoryServices-ISM-Smtp -- Disabled
				DirectoryServices-ADAM-Tools -- Enabled
				DirectoryServices-AdministrativeCenter -- Enabled
				ActiveDirectory-PowerShell -- Enabled
				DirectoryServices-ADAM -- Disabled
				DirectoryServices-DomainController -- Disabled
				DirectoryServices-DomainController-Tools -- Enabled

		#>

		if (($arrFiltered.Count -eq 0) -or ($arrFiltered.Count -eq $null)){
			#NOT installed.
			$bInstalled = $False;
			if (($arrResults -match "Error:") -ne ""){
				#To catch the "Error: 740" error if no permissions to "read" the installed "Windows Features".
				#Should translate to:    "Error(740): The requested operation requires elevation."
				#Write-Host "Error"
				$bInstalled = $arrResults;
			}
		}
		else{
			#Installed
			#Client Checks:
			if ((($arrFiltered -Match "RemoteServerAdministrationTools-Roles-AD-Powershell -- Enabled").Count -eq 0) -or (($arrFiltered -Match "RemoteServerAdministrationTools-Roles-AD -- Enabled").Count -eq 0)){
				#Client check failed.
				#Server Checks:  (I think "ActiveDirectory-PowerShell" is the important one, but not 100% sure still)
				if ((($arrFiltered -Match "ActiveDirectory-PowerShell -- Enabled").Count -eq 1) -or ((($arrFiltered -Match "DirectoryServices-ADAM -- Enabled").Count -eq 1) -and (($arrFiltered -Match "DirectoryServices-ADAM-Tools -- Enabled").Count -eq 1))){
					#AD Checkboxes are Checked.
					$bInstalled = $True;
				}
				else{
					#AD Checkboxes are NOT Checked.
					$bInstalled = $False;
				}
			}
			else{
				#AD Checkboxes are Checked.
				$bInstalled = $True;
			}

			#if (($bInstalled -eq $False) -and ($bEnable)){
			if ($bEnable){
				if (($arrFiltered -Match "RemoteServerAdministrationTools").Count -gt 0){
					#Turn ON these features
					$strResults = DISM /online /enable-feature /featurename:RemoteServerAdministrationTools /featurename:RemoteServerAdministrationTools-Roles /featurename:RemoteServerAdministrationTools-Roles-AD /featurename:RemoteServerAdministrationTools-Roles-AD-DS /featurename:RemoteServerAdministrationTools-Roles-AD-DS-SnapIns /featurename:RemoteServerAdministrationTools-Roles-AD-DS-AdministrativeCenter /featurename:RemoteServerAdministrationTools-Roles-AD-LDS /featurename:RemoteServerAdministrationTools-Roles-AD-Powershell /featurename:RemoteServerAdministrationTools-Roles-DNS /featurename:RemoteServerAdministrationTools-Roles-RDS;
					if ([String]($strResults -Match "completed successfully") -Like "*successfully*"){
						$bInstalled = $True;
					}
				}
			}

			if ($bDisable){
				if (($arrFiltered -Match "RemoteServerAdministrationTools").Count -gt 0){
					#Turn OFF these features
					$strResults = DISM /online /disable-feature /featurename:RemoteServerAdministrationTools-Features /featurename:RemoteServerAdministrationTools-Features-BitLocker /featurename:RemoteServerAdministrationTools-Features-Clustering /featurename:RemoteServerAdministrationTools-Features-GP /featurename:RemoteServerAdministrationTools-Features-LoadBalancing /featurename:RemoteServerAdministrationTools-Features-SmtpServer /featurename:RemoteServerAdministrationTools-Features-StorageExplorer /featurename:RemoteServerAdministrationTools-Features-StorageManager /featurename:RemoteServerAdministrationTools-Features-Wsrm /featurename:RemoteServerAdministrationTools-Roles-AD-DS-NIS /featurename:RemoteServerAdministrationTools-Roles-DHCP /featurename:RemoteServerAdministrationTools-Roles-HyperV /featurename:RemoteServerAdministrationTools-Roles-FileServices /featurename:RemoteServerAdministrationTools-Roles-FileServices-Dfs /featurename:RemoteServerAdministrationTools-Roles-FileServices-Fsrm /featurename:RemoteServerAdministrationTools-Roles-FileServices-StorageMgmt /featurename:RemoteServerAdministrationTools-Roles-CertificateServices /featurename:RemoteServerAdministrationTools-Roles-CertificateServices-CA /featurename:RemoteServerAdministrationTools-Roles-CertificateServices-OnlineResponder /featurename:RemoteServerAdministrationTools-ServerManager;
				}
			}
		}
		return $bInstalled;
	}

	#function isNumeric($intX){
	#	#Check if passed in value is a number.
	#		#IsNumeric() equivelant is -> [Boolean]([String]($x -as [int]))
	#	#Returns: True or False
	#	#$intX = The value to check to see if it is a number.

	#	#http://rosettacode.org/wiki/Determine_if_a_string_is_numeric
	#	try {
	#		0 + $intX | Out-Null;
	#		return $True;
	#	}
	#	catch {
	#		return $False;
	#	}
	#}

	function LocalToUTC{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True, HelpMessage = "Local time to convert to UTC / GMT time.")][String]$strTime
		)
		#Converts passed in time, local time, to UTC.
		#Returns: 

		return ((Get-Date $strTime).ToUniversalTime()).ToString();
	}

	function SaveConfig{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strProject, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][Hashtable]$hSettings, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strWhatSet = "Personal"
		)
		#Save config/ini info/file.
		#Returns: True or False.
		#$strProject = The Project/file name.
		#$hSettings = A HashTable/Array of the settings to save.  MUST provide at least one key/setting.  Providing 0 Keys triggers a config file reset.
		#$strWhatSet = What settings/file info to set/save.  "Personal", "Global".
			#Global settings file is in Project root dir.
			#Personal settings file is in users "My Documents" dir.

		$bolComplete = $False;
		$strConfigFile = $strProject + ".ini";

		if ($strWhatSet -eq "Global"){
			#Get Global path
			#https://blogs.technet.microsoft.com/heyscriptingguy/2014/08/03/weekend-scripter-a-hidden-gem-in-the-powershell-ocean-get-pscallstack/
			$objCallStack = Get-PSCallStack | Select-Object -Property *;
			#$strLastCmd = $objCallStack[0].Command;
			#$strFirstCmd = $objCallStack[($objCallStack.Count - 1)].Command;
			$strFirstCmd = $objCallStack[-1].Command;
			if (($strFirstCmd -eq "prompt") -and ($objCallStack.Count -ge 2)){
				$strPathG = Split-Path $objCallStack[-2].ScriptName;
			}
			else{
				$strPathG = Split-Path $objCallStack[-1].ScriptName;
			}
			if (!($strPathG.EndsWith("\"))){
				$strPathG = $strPathG + "\";
			}

			#Verify $strPathG.
			if (!([String]::IsNullOrEmpty($strPathG))){
				if (Test-Path -Path $strPathG){
					$strConfigFile = $strPathG + $strConfigFile;
					#if (Test-Path -Path $strConfigFile){
					#	#Read file into Hash Array
					#	$arrFile = [System.IO.File]::ReadAllLines($strConfigFile);
					#	for ($intX = 0; $intX -lt $arrFile.Count; $intX++){
					#		$strKey = "";
					#		$strVal = "";
					#		if ($arrFile[$intX].Contains(" : ")){
					#			$strKey = $arrFile[$intX].SubString(0, $arrFile[$intX].IndexOf(" : ")).Trim();
					#			$strVal = $arrFile[$intX].SubString($arrFile[$intX].IndexOf(" : ") + 2).Trim();
					#		}

					#		if (!([String]::IsNullOrEmpty($strKey))){
					#			foreach ($strEntry in $hSettings.Keys){
					#				if ($hSettings.ContainsKey($strKey)){
					#					$arrFile[$intX] = $strKey + " : " + $hSettings.$strKey;
					#					break;
					#				}
					#			}
					#		}
					#	}
					#}
				}
			}
		}

		if ($strWhatSet -eq "Personal"){
			#Get Personal path
			$strPathP = [Environment]::GetFolderPath("MyDocuments");
			if (!($strPathP.EndsWith("\"))){
				$strPathP = $strPathP + "\";
			}

			#Verify $strPathP.
			if (!([String]::IsNullOrEmpty($strPathP))){
				if (Test-Path -Path $strPathP){
					$strConfigFile = $strPathP + $strConfigFile;
					#if (Test-Path -Path $strConfigFile){
					#	#Read file into Hash Array
					#	$arrFile = [System.IO.File]::ReadAllLines($strConfigFile);
					#	for ($intX = 0; $intX -lt $arrFile.Count; $intX++){
					#		$strKey = "";
					#		$strVal = "";
					#		if ($arrFile[$intX].Contains(" : ")){
					#			$strKey = $arrFile[$intX].SubString(0, $arrFile[$intX].IndexOf(" : ")).Trim();
					#			$strVal = $arrFile[$intX].SubString($arrFile[$intX].IndexOf(" : ") + 2).Trim();
					#		}

					#		if (!([String]::IsNullOrEmpty($strKey))){
					#			foreach ($strEntry in $hSettings.Keys){
					#				if ($hSettings.ContainsKey($strKey)){
					#					$arrFile[$intX] = $strKey + " : " + $hSettings.$strKey;
					#					break;
					#				}
					#			}
					#		}
					#	}
					#}
				}
			}
		}

		if ($strConfigFile -ne ($strProject + ".ini")){
			if (Test-Path -Path $strConfigFile){
				if ($hSettings.Count -gt 0){
					#Read file into an Array
					$arrFile = [System.IO.File]::ReadAllLines($strConfigFile);
					for ($intX = 0; $intX -lt $arrFile.Count; $intX++){
						$strKey = "";
						$strVal = "";
						if (!([String]::IsNullOrEmpty($arrFile[$intX]))){
							if ($arrFile[$intX].Contains(" : ")){
								$strKey = $arrFile[$intX].SubString(0, $arrFile[$intX].IndexOf(" : ")).Trim();
								$strVal = $arrFile[$intX].SubString($arrFile[$intX].IndexOf(" : ") + 2).Trim();
							}

							if (!([String]::IsNullOrEmpty($strKey))){
								#Check if the provided info is already in the config file, update it if so.
								if ($hSettings.ContainsKey($strKey)){
									#$arrFile[$intX] = $strKey + " : " + $hSettings.$strKey;
									#$hSettings.Remove($strKey);
								}
								else{
									$hSettings.Add($strKey, $strVal);
								}
							}
						}
					}
				}
				else{
					#$hSettings.Count is 0, so delete the config file.
					$intX = 0;
					do {
						$intX++;
						$Error.Clear();
						$strResults = Remove-Item $strConfigFile;
					} while (($Error) -and ($intX -lt 10));
					if (!($Error)){
						$bolComplete = $True;
					}
				}
			}
			#else{
			#	#Currently no existing config file, so will need to create one.
			#}

			if ($hSettings.Count -gt 0){
				$arrFile = @();
				foreach ($strKey in $hSettings.Keys){
					$arrFile += $strKey + " : " + $hSettings.$strKey;
				}

				$Error.Clear();
				#$arrFile | Out-File -Append -filepath $strConfigFile -Encoding Default;
				$arrFile | Out-File -filepath $strConfigFile -Encoding Default;
				if (!($Error)){
					$bolComplete = $True;
				}
			}
		}

		return $bolComplete;
	}

	function UpdateResults{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strText, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bolClear = $False, 
			[ValidateNotNull()][Parameter(Mandatory=$False)]$objControl, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strLogDir, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strLogFile
		)
		#Returns: 
		#$strText = The text to put in $objControl ($objControl ideally should to be a TextBox).  ($txbResults by default)
		#$bolClear = True or False.  Clear the Control B4 entering $strText into it.
		#$objControl = The control to put $strText into.
		#$strLogDir = (Only needed if not a "global" variable.) The path to the directory where log files for the running project are stored.
		#$strLogFile = (Only needed if not a "global" variable.) The name of the log file to write info to.

		#Write to local log file
		if ((!([String]::IsNullOrEmpty($strLogDir))) -and (!([String]::IsNullOrEmpty($strLogFile)))){
			if ((!([String]::IsNullOrEmpty($strText.Trim()))) -and ($strText.Trim() -ne "`r`n")){
				WriteLogFile (" " + $strText.Replace("`r`n", " ")) $strLogDir $strLogFile;
			}
		}

		if ([String]::IsNullOrEmpty($objControl)){
			if ([String]::IsNullOrEmpty($txbResults)){
				return;
			}
			$objControl = $txbResults;
		}

		if ($bolClear -eq $True){
			$objControl.Text = "";
		}

		#$objControl.Text = $objControl.Text + $strText;		#Does not show new messages that are appended.
		$objControl.AppendText($strText);						#Scrolls to the bottom so the appended messages are visible.

		#$frmAScIIGUI.Update();
		#$frmAScIIGUI.Refresh();
		[System.Windows.Forms.Application]::DoEvents();
	}

	function UTCToLocal{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True, HelpMessage = "UTC / GMT time to convert to Local time.")][String]$strTime
		)
		#Convert passed in time, UTC time, to local time.
		#Returns: 

		return [System.TimeZone]::CurrentTimeZone.ToLocalTime($strTime);
	}

	function VerifyPathing{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$sLocalPath, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$sSharePath

		)
		#Checks if both paths exist, and tries creating them if not, returns Share path unless it does not exist.
		#Returns: 
		#$sLocalPath = The local path for the program.
		#$sSharePath = The Share path for the program. 

		if ((!(Test-Path -Path $sSharePath)) -and ($sSharePath -ne "")){
			#Need to create the directory
			if ((Test-Path -Path ("\\" + $sSharePath.Split("\")[2] + "\" + $sSharePath.Split("\")[3]))){
				#PS mkdir, will create any parent folders needed.
				$strResults = mkdir $sSharePath;
			}
		}
		if ((!(Test-Path -Path $sLocalPath)) -and ($sLocalPath -ne "")){
			#Need to create the directory
			#PS mkdir, will create any parent folders needed.
			$strResults = mkdir $sLocalPath;
		}
		#Set logging path
		if ((!(Test-Path -Path $sSharePath)) -or ($sSharePath -eq "")){
			$sLogDir = $sLocalPath;
		}
		else{
			$sLogDir = $sSharePath;
		}

		return , $sLogDir;
	}

