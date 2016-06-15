###########################################
# Updated Date:	14 June 2016
# Purpose:		Common routines to all/most projects.
# Requirements: DB-Routines.ps1 for the CheckVer() routine.
#				.\MiscSettings.txt
##########################################

	#For use with CheckVer() and LoadRequired().
	if ($global:LoadedFiles -eq $null){
		#($global:LoadedFiles.GetType().Name -ne "Hashtable")
		$global:LoadedFiles = @{};
	}

	function AsAdmin{
		#Checks if the loged in user of the PowerShell session has admin privileges.

		$bolAsAdmin = $False;

		#Next little block is based off the info found in the following URL:
			#http://blogs.msdn.com/b/virtual_pc_guy/archive/2010/09/23/a-self-elevating-powershell-script.aspx

		# Get the ID and security principal of the current user account
		$myWindowsID = [System.Security.Principal.WindowsIdentity]::GetCurrent();
		$myWindowsPrincipal = new-object System.Security.Principal.WindowsPrincipal($myWindowsID);

		# Get the security principal for the Administrator role
		$adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator;

		# Check to see if we are currently running "as Administrator"
		if ($myWindowsPrincipal.IsInRole($adminRole)){
			#Write-Host "Your Admin";
			$bolAsAdmin = $True;
		}else{
			#Write-Host "NOT Admin";
			$bolAsAdmin = $False;
		}

		return $bolAsAdmin;
	}

	function BackUpDir{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strSourceDir, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strDestDir, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bSubs = $False, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bPrompts = $True, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bSkip = $True, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strBackUpDir
		)
		#Copies/Backups Source Directory files to Destination Directory.
			#The SrcFile.LastWriteTime MUST be greater than the DestFileLastWriteTime, or the file is NOT copied/backedup.
			#(If a file starts with a 2 digit #, it is assumed to be a backup file, and is NOT copied/backedup.)
		#Returns a string, if work was done in the format of:    "Copied # of # files."
		#$strSourceDir = The Source Directory.
		#$strDestDir = The Destination Directory.
		#$bSubs = $True or $False.  Also backup subdirectories.
		#$bPrompts = $True or $False.  Prompt before copying missing files.  ($False will copy all missing files.)
		#$bSkip = $True or $False.  Skip "Special" files.  ("*.lnk", "*.db", "*.md", "*.sln", "*.pssproj", ".git*", "*Test*")
		#$strBackUpDir = The Directory to create BackUp copies in. [Defaults to ($strDestDir + "..\BackUps\").]

		$arrSkipEndings = @(".lnk", ".db", ".md", ".sln", ".pssproj");
		$arrSkipStarts = @(".git");
		$arrSkipContains = @("Test");
		$bDoAll = $False;
		$strMessage = "";

		if ($bPrompts -eq $True){
			if ([String]::IsNullOrEmpty($strSourceDir.Trim())){
				$strTempPath = "";
				$strTempPath = (GetPathing "Dev").Returns.Rows[0]['Path'];
				$strTempPath = $strTempPath + "PS-Scripts\";
				$strSourceDir = Read-Host "What is the Source Directory? `r`n  (defaults to: $strTempPath) `r`n";
				if ([String]::IsNullOrEmpty($strSourceDir.Trim())){
					$strSourceDir = $strTempPath + "PS-Scripts\";
				}
			}
			if ([String]::IsNullOrEmpty($strDestDir.Trim())){
				$strTempPath = "";
				$strTempPath = (GetPathing "Scripts").Returns.Rows[0]['Path'];
				$strTempPath = Read-Host "What is the Destination Directory? `r`n  (defaults to: $strDestDir) `r`n";
				if ([String]::IsNullOrEmpty($strDestDir.Trim())){
					$strDestDir = strTempPath;
				}
			}
		}
		else{
			if (([String]::IsNullOrEmpty($strSourceDir)) -or ([String]::IsNullOrEmpty($strDestDir))){
				return $False;
			}

			#If no prompting, then copy all.
			$bDoAll = $True;
		}

		#Setup $strBackUpDir.
		if ([String]::IsNullOrEmpty($strBackUpDir)){
			#$strBackUpDir = (GetPathing "BackUps").Returns.Rows[0]['Path'];
			$strBackUpDir = $strDestDir + "..\";
			$objDir = Get-Item $strBackUpDir -Force;
			#This block is not right, but the idea is there.
				#$strDestDirEnd = $strDestDir.Split("\");
				#if ([String]::IsNullOrEmpty($strDestDirEnd[-1])){
				#	$strDestDirEnd = $strDestDirEnd[-2];
				#}
				#else{
				#	$strDestDirEnd = $strDestDirEnd[-1];
				#}
				#if (($objDir.Name -eq "Beta") -or ($objDir.Name -eq $strDestDirEnd)){
			if ($objDir.Name -eq "Beta"){
				$strBackUpDir = $strBackUpDir + "..\";
			}
			$strBackUpDir = $strBackUpDir + "BackUps\";
			if (!(Test-Path -Path $strBackUpDir)){
				$strUpdate = "Creating BackUp Directory: " + $strBackUpDir;
				if ($bPrompts -eq $True){
					Write-Host $strUpdate;
				}
				else{
					$strMessage = $strMessage + $strUpdate + "`r`n";
				}
				mkdir $strBackUpDir | Out-Null;
			}
		}

		#Make sure directory paths end with "\".
		if (!($strSourceDir.EndsWith("\"))){
			$strSourceDir = $strSourceDir + "\";
		}
		if (!($strDestDir.EndsWith("\"))){
			$strDestDir = $strDestDir + "\";
		}
		if (!($strBackUpDir.EndsWith("\"))){
			$strBackUpDir = $strBackUpDir + "\";
		}
		#Write-Host "strSourceDir $strSourceDir";
		#Write-Host "strDestDir $strDestDir";
		#Write-Host "strBackUpDir $strBackUpDir";
		#Write-Host "";

		$objSrcSubItems = Get-ChildItem $strSourceDir -Force;		#force is necessary to get hidden files/folders
		$objDestSubItems = Get-ChildItem $strDestDir -Force;		#force is necessary to get hidden files/folders

		$intFileCount = 0;
		$intCount = 0;
		foreach ($objSrcItem in $objSrcSubItems){
			if (!([String]::IsNullOrEmpty($objSrcItem))){
				if ($objSrcItem.PSIsContainer -eq $False){
					#A File
					$intFileCount ++;
					$bolSkipFile = $False;
					if ($bSkip){
						foreach ($strCheckSkip in $arrSkipEndings){
							if ($bolSkipFile){
								break;
							}
							if ($objSrcItem.Name.EndsWith($strCheckSkip)){
								$bolSkipFile = $True;
							}
						}
						foreach ($strCheckSkip in $arrSkipStarts){
							if ($bolSkipFile){
								break;
							}
							if ($objSrcItem.Name.StartsWith($strCheckSkip)){
								$bolSkipFile = $True;
							}
						}
						foreach ($strCheckSkip in $arrSkipContains){
							if ($bolSkipFile){
								break;
							}
							if ($objSrcItem.Name.Contains($strCheckSkip)){
								$bolSkipFile = $True;
							}
						}
						if ((isNumeric $objSrcItem.Name.SubString(0, 2)) -eq $True){
							$bolSkipFile = $True;
						}
					}

					#if (($bolSkipFile -eq $False) -and ((isNumeric $objSrcItem.Name.SubString(0, 2)) -eq $False)){
					if ($bolSkipFile -eq $False){
						<#
							#Write-Host $objSrcItem;			#Same as .Name
							#Write-Host $objSrcItem.Name;
							#Write-Host $objSrcItem.FullName;
							#Write-Host $objSrcItem.Attributes;
							#Write-Host $objSrcItem.Length;
							#Write-Host $objSrcItem.CreationTime;
							#Write-Host $objSrcItem.LastWriteTime;
							#Write-Host $objSrcItem.LastAccessTime;
							#Write-Host $objSrcItem.VersionInfo;

							#(Get-Date -format "MM/dd/yyyy")
							#(Get-Date).ToString("yyyyMMdd")
							#$intTime = ([System.DateTime]::Now - $dteStart).TotalMinutes;
							#$intTime = [Math]::Round($intTime, 2);
						#>

						$bolFoundFile = $False;
						#Check if the SrcFile is in the Dest Dir already.
						foreach ($objDestItem in $objDestSubItems){
							if ($objSrcItem.Name -eq $objDestItem.Name){
								#Found the file in the Destination Directory.
								$bolFoundFile = $True;
								if (($objSrcItem.LastWriteTime -gt $objDestItem.LastWriteTime) -and ($objSrcItem.LastWriteTime -ne $objDestItem.LastWriteTime)){
									#Source file is newer
									$strUpdate = "`r`n" + $objSrcItem.Name + " (" + $objSrcItem.LastWriteTime + ") is newer than " + $objDestItem.Name + " (" + $objDestItem.LastWriteTime + ")";
									if ($bPrompts -eq $True){
										Write-Host $strUpdate;
									}
									else{
										$strMessage = $strMessage + $strUpdate + "`r`n";
									}

									#Check if have a backup file, for today.
									if (!([String]::IsNullOrEmpty($strBackUpDir))){
										$strDateCode = (Get-Date).ToString("yyyyMMdd");
										if (($strDestDir.Contains("\beta")) -or ($strDestDir.Contains("\Beta")) -or ($strDestDir.Contains("\BETA")) -or ($strDestDir -Match "\beta") -or ($strDestDir -Match "\Beta") -or ($strDestDir -Match "\BETA")){
											$strDateCode = $strDateCode + "b";
										}

										if (!(Test-Path -Path ($strBackUpDir + $strDateCode + "_" + $objSrcItem.Name))){
											$strUpdate = "    Creating a backup copy of " + $objDestItem.Name;
											if ($bPrompts -eq $True){
												Write-Host $strUpdate;
											}
											else{
												$strMessage = $strMessage + $strUpdate + "`r`n";
											}
											Copy-Item -Path $objDestItem.FullName -Destination ($strBackUpDir + $strDateCode + "_" + $objDestItem.Name);
										}
									}
									$strUpdate = "    Copying " + $objSrcItem.Name;
									if ($bPrompts -eq $True){
										Write-Host $strUpdate;
									}
									else{
										$strMessage = $strMessage + $strUpdate + "`r`n";
									}
									Copy-Item -Path $objSrcItem.FullName -Destination $objDestItem.FullName;
									$intCount ++;
								}
								else{
									$strUpdate = "      File " + $objDestItem.Name + " is up to date.";
									if ($bPrompts -eq $True){
										Write-Host $strUpdate;
									}
									else{
										$strMessage = $strMessage + $strUpdate + "`r`n";
									}
								}
							}

							if ($bolFoundFile -eq $True){
								break;
							}
						}

						if ($bolFoundFile -eq $False){
							#SrcFile was NOT found in Dest Dir.
							if ($bDoAll -eq $False){
								Write-Host "`r`n" $objSrcItem.Name "was not found in the destination directory.";
								Write-Host " Do you want to copy this file? `r`n [Y]es or [N]o or [A]ll `r`n"
								$objResponse = $host.UI.RawUI.ReadKey("NoEcho, IncludeKeyDown")
								if (($objResponse.Character -eq "A") -or ($objResponse.Character -eq "a")){
									$bDoAll = $True;
								}
							}
							if (($objResponse.Character -eq "Y") -or ($objResponse.Character -eq "y") -or ($bDoAll -eq $True)){
								$strUpdate = "    Copying " + $objSrcItem.Name;
								if ($bPrompts -eq $True){
									Write-Host $strUpdate;
								}
								else{
									$strMessage = $strMessage + $strUpdate + "`r`n";
								}
								Copy-Item -Path $objSrcItem.FullName -Destination ($strDestDir + $objSrcItem.Name);
								$intCount ++;
							}
						}
					}
					else{
						#Files to skip, or that start with 2 digit #'s.
						$strUpdate = "      Skipping " + $objSrcItem.Name;
						if ($bPrompts -eq $True){
							Write-Host $strUpdate;
						}
						else{
							$strMessage = $strMessage + $strUpdate + "`r`n";
						}
					}
				}
				else{
					#A Directory
					if ($bSubs -eq $True){
						$strNewSrc = $strSourceDir + $objSrcItem.Name;
						$strNewDest = $strDestDir + $objSrcItem.Name;

						if (!(Test-Path -Path $strNewDest)){
							$strUpdate = "Creating Directory: " + $strNewDest;
							if ($bPrompts -eq $True){
								Write-Host $strUpdate;
							}
							else{
								$strMessage = $strMessage + $strUpdate + "`r`n";
							}
							mkdir $strNewDest | Out-Null;
						}
						$strMessage = $strMessage + (BackUpDir $strNewSrc $strNewDest $bSubs $bPrompts $bSkip $strBackUpDir);
						$strMessage = $strMessage + "`r`n";
					}
				}
			}
		}

		if (($bolFoundFile -eq $True) -or ($intFileCount -gt 0)){
			$strUpdate = "Copied " + $intCount + " of " + $intFileCount + " files.";
			#if ($intCount -lt 1){
			#	$strUpdate = "";
			#}
			if ($bPrompts -eq $True){
				Write-Host $strUpdate;
			}
			else{
				$strMessage = $strMessage + $strUpdate + "`r`n";
			}
		}

		return $strMessage;
	}

	function CheckVer{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$Project, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$RunningVer
		)
		#Checks the running version of $Project against the posted Production version.
		#Can also checks that the files in $global:LoadedFiles are up to date.
		#Returns a PowerShell object.
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
					$ScriptDir = (Get-Location).ToString();
				}else{
					$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;
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
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters passed in.
			#$objReturn.Results		= $True or $False.  Was a zip file created.
			#$objReturn.Message		= "Success" or the error message.
			#$objReturn.Returns		= The full path annd file name of the file created.
		#$ZipFile = The zip file to create. (Full path) [i.e. "c:\path\file.zip"]
		#$Files = An array of the files to add to the zip file. (Full paths) [i.e. @("c:\path\file.one", "c:\path\file.two")]

		if ([String]::IsNullOrEmpty($MyInvocation.MyCommand.Path)){
			$ScriptDir = (Get-Location).ToString();
		}else{
			$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;
		}
		$strInclude = "Documents.ps1";
		if (Test-Path -Path ($ScriptDir + "\..\PS-CFW\" + $strInclude)){
			. ($ScriptDir + "\..\PS-CFW\" + $strInclude)
		}
		else{
			. ($ScriptDir + "\" + $strInclude)
		}

		$objReturn = ZipCreateFile $ZipFile $Files;

		return $objReturn;
	}

	function EnableDotNet4{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bISE2 = $False, 
			[ValidateNotNull()][Parameter(Mandatory=$False)]$strProjPath
		)
		#Checks if .NET 4 is enabled, and if NOT then creates the *.xml config file to enable .NET4 support.
		#$bISE2 = $True or $False.  Create the "*\powershell_ise.exe.config" files along with the "*\powershell.exe.config" files.
		#Returns $True if created config files, or .NET 4.x already enabled.
		#Returns $False if Config Files were NOT created.

		$bReturn = $False;
		$bolAsAdmin = $False;

		if ($PSVersionTable.CLRVersion.Major -lt 4){
			$bReturn = $True;
			$bolAsAdmin = AsAdmin;
			if ($bolAsAdmin -ne $True){
				$strMessage = "You should run this PS Script with admin permissions." + "`r`n" + "Want us to restart this PS Script for you?";
				if ((!(Get-Command "MsgBox" -ErrorAction SilentlyContinue))){
					Write-Host "`r`n$strMessage ([Y]es, [N]o)";
					#Write-Host "Press any key to continue ..."
					$strResponse = $host.UI.RawUI.ReadKey("NoEcho, IncludeKeyDown")
				}
				else{
					$strResponse = MsgBox $strMessage "Not running with Admin perms" 4;
				}

				if (($strResponse -eq "yes") -or ($strResponse -eq "y") -or ($strResponse -eq "Y") -or ($strResponse.Character -eq "yes") -or ($strResponse.Character -eq "y") -or ($strResponse.Character -eq "Y")){
					$strCommand = $MyInvocation.MyCommand.Path;
					if ([String]::IsNullOrEmpty($strCommand)){
					#if (($strCommand -eq "") -or ($strCommand -eq $Null)){
						if (!([String]::IsNullOrEmpty($strProjPath))){
						#if (($strProjPath -ne "") -and ($strProjPath -ne $Null)){
							$strCommand = $strProjPath;
						}
					}

					if (!([String]::IsNullOrEmpty($strCommand))){
					#if (($strCommand -ne "") -and ($strCommand -ne $Null)){
						$strCommand = "& '" + $strCommand + "'";

						$strMessage = "Restarting as Admin.";
						WriteLogFile $strMessage $strLogDirL $strLogFile;

						#method 1.  Uses Windows UAC to get creds.
						#Start-Process $PSHOME\powershell.exe -verb runas -Wait -ArgumentList "-Command $strCommand";				#When I use " -Wait" kicks "Access Denied" error.
						Start-Process ($PSHOME + "\powershell.exe") -verb runas -ArgumentList "-ExecutionPolicy ByPass -Command $strCommand";
						#Start-Process ($PSHOME + "\powershell.exe") -verb runas -ArgumentList "-STA -ExecutionPolicy ByPass -Command $strCommand";
						exit;

						#http://powershell.com/cs/blogs/tobias/archive/2012/05/09/managing-child-processes.aspx
						$objProcess = (Get-WmiObject -Class Win32_Process -Filter "ParentProcessID=$PID").ProcessID;
						Stop-Process -Id $PID;
					}
				}else{
					$bolAsAdmin = $True;
				}
			}

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
		else{
			$bReturn = $True;
		}

		return $bReturn;
	}

	function GetConfig{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strProject, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strWhatGet = "Personal"
		)
		#Load/get config/ini info/file.
			#Returns a HashArray.
		#$strProject = The Project/file name.
		#$strWhatGet = What settings/file info to get.  "Personal", "Global", "Both".
			#Global settings file is in Project root dir.
			#Personal settings file is in users "My Documents" dir.

		$hashSettings = @{};

		if (($strWhatGet -eq "Global") -or ($strWhatGet -eq "Both")){
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
			#Write-Host $strPathG;
			#Verify $strPathG.
			if (!([String]::IsNullOrEmpty($strPathG))){
				if (Test-Path -Path $strPathG){
					$strConfigFile = $strPathG + $strProject + ".ini";
					if (Test-Path -Path $strConfigFile){
						#Read file into Hash Array
						foreach ($strLine in [System.IO.File]::ReadAllLines($strConfigFile)){
							if ($strLine.Contains(" : ")){
								$strKey = $strLine.SubString(0, $strLine.IndexOf(" : ")).Trim();
								$strVal = $strLine.SubString($strLine.IndexOf(" : ") + 2).Trim();
								if ($hashSettings.ContainsKey($strKey)){
									$hashSettings.$strKey = $strVal
								}
								else{
									$hashSettings.Add($strKey, $strVal);
								}
							}
						}
					}
					else{
						#Write-Host "No config file at: $strConfigFile";
					}
				}
				else{
					#Write-Host "Invalid path: $strPathG";
				}
			}
		}

		if (($strWhatGet -eq "Personal") -or ($strWhatGet -eq "Both")){
			#Get Personal path
			$strPathP = [Environment]::GetFolderPath("MyDocuments");
			if (!($strPathP.EndsWith("\"))){
				$strPathP = $strPathP + "\";
			}
			#Write-Host $strPathP;
			#Verify $strPathP.
			if (!([String]::IsNullOrEmpty($strPathP))){
				if (Test-Path -Path $strPathP){
					$strConfigFile = $strPathP + $strProject + ".ini";
					if (Test-Path -Path $strConfigFile){
						#Read file into Hash Array
						foreach ($strLine in [System.IO.File]::ReadAllLines($strConfigFile)){
							if ($strLine.Contains(" : ")){
								$strKey = $strLine.SubString(0, $strLine.IndexOf(" : ")).Trim();
								$strVal = $strLine.SubString($strLine.IndexOf(" : ") + 2).Trim();
								if ($hashSettings.ContainsKey($strKey)){
									$hashSettings.$strKey = $strVal
								}
								else{
									$hashSettings.Add($strKey, $strVal);
								}
							}
						}
					}
					else{
						#Write-Host "No config file at: $strConfigFile";
					}
				}
				else{
					#Write-Host "Invalid path: $strPathP";
				}
			}
		}

		return $hashSettings;
	}

	function GetPathing{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$sName = "all"
		)
		#Querys a DB for Pathing info, so that can update pathing info w/out having to release new code versions.
		#Has default values incase DB is unreachable.
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters passed in.
			#$objReturn.Results		= # of Rows of data returning.
			#$objReturn.Message		= "Success" or the error message.
			#$objReturn.Returns		= A DataTable of Path(s).
		#$sName = The name of the path(s) to get, (All of them by default).

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

		#Write-Host "Calling routine is: " (Get-PSCallStack)[1].Command;

		$strConfigFile = "MiscSettings.txt";

		$arrDefaults = @{};
		#DB Info
		$arrDefaults.Add("AgentActivity", "c3RyREJUeXBlID0gbXNzcWwNCnN0ckRCU2VydmVyID0gTkFXRVNETklTUTcyVkJcU1E3MlZCSU5TVDAxDQpzdHJEQk5hbWUgPSBBZ2VudEFjdGl2aXR5DQpzdHJEQkxvZ2luUiA9IGFpb2RhdGFyZWFkZXINCnN0ckRCUGFzc1IgPSBDTVc2MTE2MWRhdGFyZWFkZXINCnN0ckRCTG9naW5XID0gYWlvZGF0YQ0Kc3RyREJQYXNzVyA9IENNVzYxMTYxZGF0YQ==");
		$arrDefaults.Add("AssMan", "c3RyREJTZXJ2ZXIgPSBubWNpbnJma2FzMDEubmFkc3VzZWEubmFkcy5uYXZ5Lm1pbA0Kc3RyREJTZXJ2ZXIyID0gbm1jaXNkbmlhczAxLm5hZHN1c3dlLm5hZHMubmF2eS5taWwNCnN0clBvcnQgPSAxNTIxDQpzdHJEQlR5cGUgPSBPcmFjbGUNCnN0ckRCTmFtZSA9IEFDUFJPRA0Kc3RyREJMb2dpblIgPSBpYnVsaw0Kc3RyREJQYXNzUiA9IGdCMjAlNGt1bGEyMyFBQw0Kc3RyREJMb2dpblcgPSBub25lDQpzdHJEQlBhc3NXID0gbm9uZQ==");
		$arrDefaults.Add("CDR", "c3RyREJUeXBlID0gbXNzcWwNCnN0ckRCU2VydmVyID0gbmFlYW5yZmt0bTAyDQpzdHJEQk5hbWUgPSBkYnBob2VuaXg1NTENCnN0ckRCTG9naW5SID0gaXNmdXNlcg0Kc3RyREJQYXNzUiA9IG4vYQ0Kc3RyREJMb2dpblcgPSBpc2Z1c2VyDQpzdHJEQlBhc3NXID0gbi9h");
		$arrDefaults.Add("ECMD", "c3RyREJUeXBlID0gbXNzcWwNCnN0ckRCU2VydmVyID0gTkFFQU5SRktTUTUzXFNRNTNJTlNUMDENCnN0ckRCTmFtZSA9IEVDTUQNCnN0ckRCTG9naW5SID0ga2JTaXRlQ29kZURCVXNlcg0Kc3RyREJQYXNzUiA9IEtCU2l0QENvZEBVc2VyMQ0Kc3RyREJMb2dpblcgPSBub25lDQpzdHJEQlBhc3NXID0gbm9uZQ==");
		$arrDefaults.Add("Score", $arrDefaults.AgentActivity);
		$arrDefaults.Add("Sites", "c3RyREJUeXBlID0gbXNzcWwNCnN0ckRCU2VydmVyID0gTkFFQU5SRktTUTc1VkFcU1E3NVZBSU5TVDAxDQpzdHJEQk5hbWUgPSBTaXRlQ29kZXMNCnN0ckRCTG9naW5SID0gS0J1c2VyDQpzdHJEQlBhc3NSID0ga2M1JHNxMDI=");
		$arrDefaults.Add("SRMDB", "c3RyREJUeXBlID0gbXNzcWwNCnN0ckRCU2VydmVyID0gTkFXRVNETklTUTcyVkJcU1E3MlZCSU5TVDAxDQpzdHJEQk5hbWUgPSBTUk1fQXBwc19Ub29scw0Kc3RyREJMb2dpblIgPSBTUk1fQXBwc19Ub29sc19XRk0NCnN0ckRCUGFzc1IgPSAhU1JNX0FwcHNfVG9vbHNfV0ZNNjkNCnN0ckRCTG9naW5XID0gU1JNX0FwcHNfVG9vbHMNCnN0ckRCUGFzc1cgPSAhU1JNX0FwcHNfVG9vbHM2OQ==");
		#File Share Info ( MUST end in \ )
		$arrDefaults.Add("BackUps", "\\nawesdnifs101v.nadsuswe.nads.navy.mil\NMCIISF02$\ITSS-Tools\BackUps\");
		$arrDefaults.Add("Beta", "\\nawesdnifs101v.nadsuswe.nads.navy.mil\NMCIISF02$\ITSS-Tools\Beta\");
		$arrDefaults.Add("CFW", "\\nawesdnifs101v.nadsuswe.nads.navy.mil\NMCIISF02$\ITSS-Tools\PS-CFW\");
		$arrDefaults.Add("Dev", "C:\Projects\");
		$arrDefaults.Add("Local", "C:\Users\Public\ITSS-Tools\");
		$arrDefaults.Add("Logs", "\\NAWESPSCFS101V.NADSUSWE.NADS.NAVY.MIL\ISF-IOS$\IOS-LOGS\");
		$arrDefaults.Add("Logs_ITSS", "\\NAWESPSCFS101V.NADSUSWE.NADS.NAVY.MIL\ISF-IOS$\ITSS-Tools\Logs\");
		$arrDefaults.Add("Reports", "\\nawesdnifs101v.nadsuswe.nads.navy.mil\NMCIISF02$\ITSS-Tools\Reports\");
		$arrDefaults.Add("Root", "\\nawesdnifs101v.nadsuswe.nads.navy.mil\NMCIISF02$\ITSS-Tools\");
		$arrDefaults.Add("Scripts", "\\nawesdnifs101v.nadsuswe.nads.navy.mil\NMCIISF02$\ITSS-Tools\PS-Scripts\");
		$arrDefaults.Add("SupportFiles", "\\nawesdnifs101v.nadsuswe.nads.navy.mil\NMCIISF02$\ITSS-Tools\SupportFiles\");
		$arrDefaults.Add("ITSS-Tools", $arrDefaults.Root);

		$strConfigFile = $arrDefaults.SupportFiles + $strConfigFile;

		#Config file  (takes about 1 sec)
		if (($objReturn.Results -eq $False) -or ($objReturn.Results -lt 1)){
			#Write-Host "Check if the config file exist";
			if (!(Test-Path -Path $strConfigFile)){
				if (Test-Path -Path (".\..\PS-CFW\" + $strConfigFile)){
					$strConfigFile = ".\..\PS-CFW\" + $strConfigFile;
				}
			}
			if (Test-Path -Path $strConfigFile){
				#Write-Host "Config file exists";

				#Create the DataTable to return
				$objTable = New-Object System.Data.DataTable;
				$col1 = New-Object System.Data.DataColumn Name,([String]);
				$col2 = New-Object System.Data.DataColumn Path,([String]);
				$col3 = New-Object System.Data.DataColumn Description,([String]);
				$objTable.columns.add($col1);
				$objTable.columns.add($col2);
				$objTable.columns.add($col3);

				$Error.Clear();
				foreach ($strLine in [System.IO.File]::ReadAllLines($strConfigFile)){
					if (($strLine.StartsWith($sName)) -or ($sName -eq "all")){
						#Need to accomodate commented lines, especially for when $sName is "all".
						if (!($strLine.StartsWith("--"))){
							$strRawName = $strLine.SubString(0, $strLine.IndexOf("=") - 1).Trim();
							$strRawPath = $strLine.SubString($strLine.IndexOf("=") + 1).Trim();

							$row = $objTable.NewRow();
							$row.Name = $strRawName;
							$row.Path = $strRawPath;
							$row.Description = $null;
							$objTable.Rows.Add($row);
						}
					}
				}

				if ($objTable.Rows.Count -gt 0){
					$objReturn.Message = "Success";
					$objReturn.Results = $objTable.Rows.Count;
				}
			}
		}

		#DB  (takes 2 to 3 sec)
		if ((($objReturn.Results -eq $False) -or ($objReturn.Results -lt 1)) -and (((Get-PSCallStack)[1].Command -ne "GetDBInfo") -and ((Get-PSCallStack)[1].Command -ne "GetPathing"))){
			#??  Update Share config file (used by apps/tools NOT PowerShell).  ??

			#No config file, or no entry, so check DB.
			#Make sure the DB routines that are in DB-Routines.ps1 are loaded.
			if ((!(Get-Command "GetDBInfo" -ErrorAction SilentlyContinue)) -or (!(Get-Command "QueryDB" -ErrorAction SilentlyContinue))){
				if ([String]::IsNullOrEmpty($MyInvocation.MyCommand.Path)){
					$ScriptDir = (Get-Location).ToString();
				}else{
					$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;
				}
				if ((Test-Path (".\DB-Routines.ps1"))){
					. (".\DB-Routines.ps1");
				}
				else{
					if ((Test-Path (".\..\PS-CFW\DB-Routines.ps1"))){
						. (".\..\PS-CFW\DB-Routines.ps1");
					}
					else{
						#When calling from a DOS command prompt, the PS executes from a ?Random? directory, and so the above "relative" checks fail.
						$strMyLoc = (Get-PSCallStack | Select-Object -Property *)[0].ScriptName;
							#Now $strMyLoc = "C:\Projects\PS-CFW\Common.ps1"
						$strMyLoc = $strMyLoc.Replace($strMyLoc.Split("\")[-1], "");
							#Now $strMyLoc = "C:\Projects\PS-CFW\"
						if ((Test-Path ($strMyLoc + "DB-Routines.ps1"))){
							. ($strMyLoc + "DB-Routines.ps1");
						}
						else{
							if ((Test-Path ($arrDefaults."Dev" + "PS-CFW\DB-Routines.ps1"))){
								. ($arrDefaults."Dev" + "PS-CFW\DB-Routines.ps1");
							}
							else{
								if ((Test-Path ($arrDefaults."Local" + "PS-CFW\DB-Routines.ps1"))){
									. ($arrDefaults."Local" + "PS-CFW\DB-Routines.ps1");
								}
								else{
									if ((Test-Path ($arrDefaults."Root" + "PS-CFW\DB-Routines.ps1"))){
										. ($arrDefaults."Root" + "PS-CFW\DB-Routines.ps1");
									}
								}
							}
						}
					}
				}
			}

			#Write-Host "Query DB";
			$arrDBInfo = GetDBInfo "AgentActivity";
			#$strSQL = "SELECT * FROM NetPath WHERE Name like '" + $sName + "'";
			if ($sName -eq "all"){
				$strSQL = "GetSP_spGetNetPath";
			}
			else{
				$strSQL = "GetSP_spGetNetPath '" + $sName + "';";
			}
			$objTable = $null;
			$Error.Clear();
			$objTable = QueryDB $arrDBInfo[1] $arrDBInfo[2] $strSQL $True;

			if (!(($objTable.Rows[0].Message -eq "Error") -or ($Error) -or ($objTable -eq $null) -or ($objTable.Rows.Count -eq 0))){
				$objReturn.Message = "Success";
				$objReturn.Results = $objTable.Rows.Count;
			}
		}

		#Hard Coded
		#Check if File and DB failed.
		if (($objReturn.Results -eq $False) -or ($objReturn.Results -lt 1)){
			#Both the Config file and the DB process failed, so return the default hard coded values.
			#Write-Host "Hard coded values";
			#Create the DataTable to return
			$objTable = New-Object System.Data.DataTable;
			$col1 = New-Object System.Data.DataColumn Name,([String]);
			$col2 = New-Object System.Data.DataColumn Path,([String]);
			$col3 = New-Object System.Data.DataColumn Description,([String]);
			$objTable.columns.add($col1);
			$objTable.columns.add($col2);
			$objTable.columns.add($col3);

			if ($arrDefaults.ContainsKey($sName)){
				#Populate the DataTable, if we have the desired info.
				$row = $objTable.NewRow();
				$row.Name = $sName;
				#$row.Path = $strRawPath;
				$row.Path = $arrDefaults."$sName";
				$row.Description = "HardCoded Value.";
				$objTable.Rows.Add($row);

				$objReturn.Message = "Success";
				$objReturn.Results = $objTable.Rows.Count;
			}
			else{
				if ($sName -eq "all"){
					foreach ($strEntry in $arrDefaults.Keys){
						$row = $objTable.NewRow();
						$row.Name = $strEntry;
						$row.Path = $arrDefaults."$strEntry";
						$row.Description = "HardCoded Value.";
						$objTable.Rows.Add($row);
					}

					$objReturn.Message = "Success";
					$objReturn.Results = $objTable.Rows.Count;
				}
			}
		}

		$objReturn.Returns = $objTable;

		return $objReturn;

	}

	function isADInstalled{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bEnable = $False, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bDisable = $False
		)
		#Check if have AD Installed and Enabled.
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
		#Check if passed in value is a number.
		#IsNumeric() equivelant is -> [Boolean]([String]($x -as [int]))

		#http://rosettacode.org/wiki/Determine_if_a_string_is_numeric
		try {
			0 + $intX | Out-Null;
			return $True;
		}
		catch {
			return $False;
		}
	}

	function LoadRequired{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][Array]$RequiredFiles, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$RootDir, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$LogDir, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$LogFile
		)
		#Loads/includes ("dot" sources) all the files specified in $RequiredFiles.
			#This routine checks to see if the file to include exists in "..\PS-CFW\", if not assumes the files are in $RootDir.
			#Because this is a function the routines loaded are only available in this scope and NOT in the calling routines (the project).
			#So based on the following URL can either read in the files, and update all functions to "Global:Function" or we can update ALL the scripts to have the "Global:" value.
				#http://stackoverflow.com/questions/15187510/dot-sourcing-functions-from-file-to-global-scope-inside-of-function
			#Above method would not work.  But found the following too, and it works.
				#https://blairconrad.wordpress.com/2010/01/29/expand-your-scope-you-can-dot-source-more-than-just-files/
		#Returns $True or $False.  $True if no errors, else $False.
		#Updates $global:LoadedFiles.
		#$RequiredFiles = An array of the files to "dot" source / include.
		#$RootDir = The (Split-Path $MyInvocation.MyCommand.Path) of the running project.
		#$LogDir = The log Directory, that contains $LogFile, that any errors will be reported to.
		#$LogFile = The Log file that any errors will be reported to.
		#The following have some good ideas:
		#http://poshcode.org/668
		#http://www.gsx.com/blog/bid/81096/Enhance-your-PowerShell-experience-by-automatically-loading-scripts

		$bLoaded = $True;

		#Make sure $RootDir does NOT have a trailing slash.
		if ($RootDir.EndsWith("\")){
			$RootDir = $RootDir.SubString(0, $RootDir.Length - 1);
		}

		foreach ($strInclude in $RequiredFiles){
			$Error.Clear();
			if (Test-Path -Path ($RootDir + "\PS-CFW\" + $strInclude)){
				#if (Test-Path -Path ($RootDir + "\PS-CFW\" + $strInclude)){
					. ($RootDir + "\PS-CFW\" + $strInclude);
					$strFile = ($RootDir + "\PS-CFW\" + $strInclude);
				#}
				#else{
				#	. ($RootDir + "\" + $strInclude);
				#	$strFile = ($RootDir + "\" + $strInclude);
				#}
			}
			else{
				if (($RootDir.EndsWith("\PS-CFW")) -and ((Test-Path -Path ($RootDir + "\" + $strInclude)))){
					. ($RootDir + "\" + $strInclude);
					$strFile = ($RootDir + "\" + $strInclude);
				}
				else{
					. ($RootDir + "\..\PS-CFW\" + $strInclude);
					$strFile = ($RootDir + "\..\PS-CFW\" + $strInclude);
				}
			}

			if ($Error){
				#$strMessage = "------- Error 'loading' '$strInclude.ps1'." + "`r`n" + $Error;
				$strMessage = "------- Error 'loading' '$strInclude'." + "`r`n" + $Error;
				Write-Host $strMessage;
				$bLoaded = $False;

				if ((($LogDir -ne "") -and ($LogDir -ne $null)) -and (($LogFile -ne "") -and ($LogFile -ne $null))){
					WriteLogFile $strMessage $LogDir $LogFile;
				}
				$Error.Clear();
			}
			else{
				#Update $global:LoadedFiles
				$objFile = Get-Item -LiteralPath $strFile;
				$Date = $objFile.LastWriteTime;
				$strDateVer = (((Get-Date $Date).ToString("yyyyMMdd.hhmmss")));
				$Date = (((Get-Date $Date).ToUniversalTime()).ToString());

				#as a hash
				#$global:LoadedFiles."Common.ps1".Date
				$global:LoadedFiles.($objFile.Name) = (@{"Ver" = $strDateVer; "Date" = $Date; "Path" = ($objFile.FullName).Replace(($objFile.Name), "")});
				#Cound use the following, but then if have the entry already things error, the above just updates it if already exist.
					#$global:LoadedFiles.Add(($objFile.Name), (@{"Ver" = $strDateVer; "Date" = $Date}));
			}
		}

		return $bLoaded;
	}

	function LocalToUTC{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True, HelpMessage = "Local time to convert to UTC / GMT time.")][String]$strTime
		)
		#Converts passed in time, local time, to UTC.

		return ((Get-Date $strTime).ToUniversalTime()).ToString();
	}

	function SaveConfig{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strProject, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][Hashtable]$hSettings, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strWhatSet = "Personal"
		)
		#Save config/ini info/file.
			#Returns True or False.
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
					#Read file into Hash Array
					$arrFile = [System.IO.File]::ReadAllLines($strConfigFile);
					for ($intX = 0; $intX -lt $arrFile.Count; $intX++){
						$strKey = "";
						#$strVal = "";
						if (!([String]::IsNullOrEmpty($arrFile[$intX]))){
							if ($arrFile[$intX].Contains(" : ")){
								$strKey = $arrFile[$intX].SubString(0, $arrFile[$intX].IndexOf(" : ")).Trim();
								#$strVal = $arrFile[$intX].SubString($arrFile[$intX].IndexOf(" : ") + 2).Trim();
							}

							if (!([String]::IsNullOrEmpty($strKey))){
								#foreach ($strEntry in $hSettings.Keys){
									if ($hSettings.ContainsKey($strKey)){
										$arrFile[$intX] = $strKey + " : " + $hSettings.$strKey;
								#		break;
									}
								#}
							}
						}
					}

					$Error.Clear();
					$arrFile | Out-File -filepath $strConfigFile -Encoding Default;
					if (!($Error)){
						$bolComplete = $True;
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
			else{
				#Currently no existing config file, so need to create one.
				if ($hSettings.Count -gt 0){
					$arrFile = @();
					foreach ($strKey in $hSettings.Keys){
						$arrFile += $strKey + " : " + $hSettings.$strKey;
					}

					$Error.Clear();
					$arrFile | Out-File -filepath $strConfigFile -Encoding Default;
					if (!($Error)){
						$bolComplete = $True;
					}
				}
			}
		}

		return $bolComplete;
	}

	function SetLogPath{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strProjName, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strLogDirS = "", 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strLogDirL = "", 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strShareLoc = "", 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strLocalLoc = ""
		)
		#$strProjName = Project Name. The directory Logs should be put in.
		#$strLogDirS = The Path to logs on the Share.
		#$strLogDirL = The Path to logs Locally.
		#$strShareLoc = The "Name" of the Share log entry in the DB/config file.
		#$strLocalLoc = The "Name" of the Local log entry in the DB/config file.
			#$strLogDir = SetLogPath $strProjName $strLogDirS $strLogDirL $strShareLoc $strLocalLoc

		#Get pathing info.
		#Share
		if (([String]::IsNullOrEmpty($strLogDirS)) -and (!([String]::IsNullOrEmpty($strShareLoc)))){
			$strPathing = GetPathing $strShareLoc;
			if ($strPathing.Results -gt 0){
				$strLogDirS = $strPathing.Returns.Rows[0].Path + $strProjName + "\";
			}
			else{
				#Default
				$strLogDirS = "\\Nawespscfs101v.nadsuswe.nads.navy.mil\isf-ios$\ITSS-Tools\Logs\" + $strProjName + "\";
			}
		}
		#Local
		if (([String]::IsNullOrEmpty($strLogDirL)) -and (!([String]::IsNullOrEmpty($strLocalLoc)))){
			$strPathing = GetPathing $strLocalLoc;
			if ($strPathing.Results -gt 0){
				$strLogDirL = $strPathing.Returns.Rows[0].Path + "Logs\" + $strProjName + "\";
			}
			else{
				$strLogDirL = "C:\Users\Public\ITSS-Tools\Logs\" + $strProjName + "\";
			}
		}

		#Make sure the log directories exist.
		#Local
		if (!([String]::IsNullOrEmpty($strLogDirL))){
			if (!(Test-Path -Path $strLogDirL)){
				#Need to create the directory
				#PS mkdir, will create any parent folders needed.
				$strResults = mkdir $strLogDirL;
			}
		}
		#Share
		if (!([String]::IsNullOrEmpty($strLogDirS))){
			if (!(Test-Path -Path $strLogDirS)){
				#Need to create the directory, after make sure the root share exists.
				if ((Test-Path -Path ("\\" + $strLogDirS.Split("\")[2] + "\" + $strLogDirS.Split("\")[3]))){
					#PS mkdir, will create any parent folders needed.
					$strResults = mkdir $strLogDirS;
				}
			}
		}

		#Set logging path
		if ([String]::IsNullOrEmpty($strLogDirS)){
			#if (Test-Path -Path $strLogDirS){
				$strLogDir = $strLogDirL;
			#}
			#else{
			#	$strLogDir = $strLogDirS;
			#}
		}
		else{
			if (Test-Path -Path $strLogDirS){
				$strLogDir = $strLogDirS;
			}
			else{
				$strLogDir = $strLogDirL;
			}
		}

		return @($strLogDir, $strLogDirS, $strLogDirL);
	}

	function UpdateLocalFiles{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strLocalDir, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strProjName, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bolDoPrompts = $False, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strLogDir = "", 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strLogFile = ""
		)
		#Update Local files from Production.   From:(\\Server\"ITSS-Root"\ + $strProjName)  To:("C:\Users\Public\ITSS-Tools\" + $strProjName)
			#Returns a string.
		#$strLocalDir = The Project/file directory.  Typically $ScriptDir.
		#$strProjName = The Project/file name.
		#$bolDoPrompts = True or False.  Should this routine be interactive.  If files were updated, prompt user to restart from local.
		#$strLogDir = Directory path to put logs.
		#$strLogFile = The file to put log entries in.

		$strResults = "No files copied/updated.";

		#Check if running from ...Projects dir.
		$strDoUpdate = "Yes";
		if ($strLocalDir.StartsWith("C:\Projects\")){
			if ($bolDoPrompts -eq $True){
				$strMessage = "Do you want to copy down any Production files that are newer? `r`n[Yes], [No], [Y]es, [N]o `r`n";
				$strDoUpdate = Read-Host $strMessage;
				Write-Host "";
			}
			else{
				$strDoUpdate = "No";

				if ((!([String]::IsNullOrEmpty($strLogDir))) -and (!([String]::IsNullOrEmpty($strLogFile)))){
					WriteLogFile "Running $strProjName from 'C:\Projects\', so NOT updating." $strLogDir $strLogFile;
				}
			}
		}

		if (($strDoUpdate -eq "yes") -or ($strDoUpdate -eq "y")){
			#Get Share pathing info.
			$strProjRootDir = GetPathing "Root";
			if ($strProjRootDir.Results -gt 0){
				$strProjRootDir = $strProjRootDir.Returns.Rows[0].Path;
			}
			else{
				#Default
				$strProjRootDir = "\\nawesdnifs101v.nadsuswe.nads.navy.mil\NMCIISF02$\ITSS-Tools\";
			}
			$strSrcCFW = (GetPathing "CFW").Returns.Rows[0].Path;

			#Check if $strLocalDir is a local path, or a network path (assume Production).
			$strMessage = "updated";
			if (($strLocalDir.StartsWith($strProjRootDir)) -or ($strLocalDir.StartsWith("\\"))){
				#Running from network, lets "install" a local copy.
				$strMessage = "installed";
				$strLocalDir = (GetPathing "Local").Returns.Rows[0].Path;
				$strLocalDirCFW = $strLocalDir + "PS-CFW\";
				$strLocalDir = $strLocalDir + $strProjName + "\";

				#Make sure have local dir.
				if (!(Test-Path -Path $strLocalDir)){
					#PS mkdir, will create any parent folders needed.
					$strResults = mkdir $strLocalDir;

					if ((!([String]::IsNullOrEmpty($strLogDir))) -and (!([String]::IsNullOrEmpty($strLogFile)))){
						WriteLogFile "Created local directory for $strProjName at: $strLocalDir." $strLogDir $strLogFile;
					}
				}
			}
			else{
				#$strLocalDirCFW = $strLocalDir.Replace($strProjName, "PS-CFW");
				$strLocalDirCFW = $strLocalDir.SubString(0, $strLocalDir.IndexOf($strProjName)) + "PS-CFW\";
			}
			$strProjRootDir = $strProjRootDir + $strProjName + "\";
			$strBackUpDir = $strLocalDir.SubString(0, $strLocalDir.IndexOf($strProjName)) + "Backups\";

			#Make sure have local CFW dir.
			if (!(Test-Path -Path $strLocalDirCFW)){
				#PS mkdir, will create any parent folders needed.
				$strResults = mkdir $strLocalDirCFW;

				if ((!([String]::IsNullOrEmpty($strLogDir))) -and (!([String]::IsNullOrEmpty($strLogFile)))){
					WriteLogFile "Created local directory for CFW at: $strLocalDirCFW." $strLogDir $strLogFile;
				}
			}

			#Copy Project files
			$strResults = "Copying Project files to $strLocalDir ... `r`n";
			$strResults = $strResults + (BackUpDir $strProjRootDir $strLocalDir $True $False $True $strBackUpDir);

			#Copy CFW files
			$strResults = $strResults + "`r`n" + "Copying CFW files to $strLocalDirCFW ... `r`n";
			$strResults = $strResults + (BackUpDir $strSrcCFW $strLocalDirCFW $True $False $True $strBackUpDir);
			if ($bolDoPrompts -eq $True){
				Write-Host $strResults;
			}

			<# ---=== Add the following code block after this routine call if you want to do reboots on update. ===---
			#Check if should restart.
			#$strMessage = "installed/updated";
			$intCount = 0;
			$arrCopied = $strResults.Split("`r`n");
			#$arrCopied = $arrCopied | ? {$_};						#Removes blank entries.
			$arrCopied = $arrCopied | ? {$_ -Match("Copied ")};		#Gets just the "Copied #...." entries.
			$arrCopied = $arrCopied | ? {$_ -Match(" files.")};		#Gets just the "....# files." entries.
			foreach ($strLine in $arrCopied){
				if ($strLine.SubString(7, 2).Trim() -gt 0){
					$intCount = $intCount + $strLine.SubString(7, 2).Trim();
				}
			}
			$strMessage = "There have been " + $intCount + " files $strMessage." + "`r`n" + "Do you want to restart $strProjName?";
			if ($intCount -gt 0){
				if (Get-Command MsgBox -ErrorAction SilentlyContinue){
					$strResponse = MsgBox $strMessage "Installed a local copy" 4;
				}
				else{
					$strMessage = $strMessage + "`r`n" + "[Yes], [No], [Yes], [No] `r`n";
					$strResponse = Read-Host $strMessage;
				}

				if (($strResponse -eq "yes") -or ($strResponse -eq "y")){
					$strCommand = "& '" + $MyInvocation.MyCommand.Path + "'";

					$strMessage = "Restarting " + $strProjName + ", as admin, from local copy.";
					if ((!([String]::IsNullOrEmpty($strLogDir))) -and (!([String]::IsNullOrEmpty($strLogFile)))){
						WriteLogFile $strMessage $strLogDir $strLogFile;
					}

					#method 1.  Uses Windows UAC to get creds.
					#Start-Process $PSHOME\powershell.exe -verb runas -Wait -ArgumentList "-Command $strCommand";				#When I use " -Wait" kicks "Access Denied" error.
					Start-Process ($PSHOME + "\powershell.exe") -verb runas -ArgumentList "-ExecutionPolicy ByPass -Command $strCommand";
					#Start-Process ($PSHOME + "\powershell.exe") -verb runas -ArgumentList "-STA -ExecutionPolicy ByPass -Command $strCommand";
					exit;
					##http://powershell.com/cs/blogs/tobias/archive/2012/05/09/managing-child-processes.aspx
					#$objProcess = (Get-WmiObject -Class Win32_Process -Filter "ParentProcessID=$PID").ProcessID;
					#Stop-Process -Id $PID;
				}
			}
			#>
		}

		return $strResults;
	}

	function UpdateResults{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strText, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bolClear = $False, 
			[ValidateNotNull()][Parameter(Mandatory=$False)]$objControl, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strLogDir, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strLogFile
		)
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

		return [System.TimeZone]::CurrentTimeZone.ToLocalTime($strTime);
	}

	function VerifyPathing{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$sLocalPath, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$sSharePath

		)
		#Checks if both paths exist, and tries creating them if not, returns Share path unless it does not exist.
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

	function WriteLogFile{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$Message, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$LogDir, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$LogFile, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$Header = ""
		)
		#Uses Out-File to append $Message to the $LogFile, in the path $LogDir.
		#$Message = The message to add to $LogFile.  gets PrePended with a "Header":
		#$LogDir = The location of $LogFile.
		#$LogFile = The file to add $Message to.  get updated to a format of "yyyymmdd_"$LogFile.  (i.e. 20150513_AscII.log)
		#$Header = A custom header to prepend $Message with, rather than the default.  ("False" for no header at all. [NOT boolean])
		#Default Header is:
			#Date Time - Domain\User - MachineName (MAC) - IP - Ticket# -- $Message
			#i.e.:
			#5/13/2015 9:23:15 - NMCI-ISF\henry.schade - ADIDBO226572 (00:24:81:21:CA:CC) - 10.12.21.80 - 8989765 -- $Message

		$intRetry = 12;

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
						$strTicketNum = "NoTic#";
					}else{
						$strTicketNum = $txbTicketNum.Text;
					}
				}else{
					$strTicketNum = "NoTic#";
				}
				$strLogHeader = (([System.DateTime]::Now).ToString() + " - " + ([System.Security.Principal.WindowsIdentity]::GetCurrent()).Name + " - " + $env:computername + " (" + (Get-WmiObject Win32_NetworkAdapterConfiguration -Namespace "root\CIMV2" | WHERE{$_.IPEnabled -eq "True"}).MACAddress + ") - " + (Get-WmiObject Win32_NetworkAdapterConfiguration -Namespace "root\CIMV2" | WHERE{$_.IPEnabled -eq "True"}).IPAddress + " - " + $strTicketNum + " -- ");
				$Message = $strLogHeader + $Message;
			}else{
				if ($Header -ne "False"){
					$Message = $Header + $Message;
				}
			}

			#Write to log file
			$intTries = 0;
			do {
				$intTries++;
				$Error.Clear();
				try{
					$Message | Out-File -filepath ($LogDir + $LogFile) -Encoding Default -Append;
				}
				catch{
					Start-Sleep -Milliseconds 100; 
				}
			} while (($Error) -and ($intTries -lt $intRetry))
		}
	}

