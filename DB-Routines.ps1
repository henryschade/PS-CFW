###########################################
# Updated Date:	6 October 2016
# Purpose:		Provide a central location for all the PowerShell DataBase routines.
# Requirements: None
##########################################

<# ---=== Change Log ===---
	#Changes for 28 June 2016:
		#Added Change Log.
	#Changes for 4 Oct 2016
		#Add CDR SIPR info to GetDBInfo()
		#Add ECMD SIPR info to GetDBInfo()
		#Update MiscSettings.txt default path to be CFW instead of SupportFiles.

#>



	function SampleUsage{
		. C:\Projects\PS-CFW\DB-Routines.ps1

		#Get SRM Change Log / version info.  (Old process)
		$arrDBInfo = GetDBInfo "SRMDB";
		$strSQL = "SELECT UpdatedDate, ChangeLog, DisableOld FROM AppChanges WHERE AppInitials='CA'";
		$Results = QueryDB $arrDBInfo[1] $arrDBInfo[2] $strSQL $False $arrDBInfo[3] $arrDBInfo[4] 180 $True;
		$Results = QueryDB $arrDBInfo[1] $arrDBInfo[2] $strSQL $False $arrDBInfo[5] $arrDBInfo[6] 180 $True;
		if ($Results.Rows[0].Message -eq "Error"){
			Write-Host "Error running " $Results.Rows[0].Results;
		}else{
			if ($Results.Rows[0].Message -eq "Success"){
				Write-Host "Success running the following SQL command:`r`n" $Results.Rows[0].Results;
				Write-Host $Results.Rows[0].FieldName1;
				Write-Host $Results.Rows[0].FieldName2;
				Write-Host $Results.Rows[1].FieldName1;
				Write-Host $Results.Rows[1].FieldName2;
			}else{
				if ($Results.Rows[0].Message -ne "Success"){
					for ($intX=0; $intX -le $Results.Length; $intX++){
						#$OutPut = $OutPut + ($Results[$intX].ItemArray -join ",");
						Write-Host ($Results[$intX].ItemArray -join ", ");
					}
				}
			}
		}


		#Pull cert info for a machine
		if (($env:UserDomain.Contains("nmci-isf")) -or ($env:UserDomain.Contains("NMCI-ISF"))){
			$arrDBInfo = GetDBInfo "ECMD";
		}
		else{
			$arrDBInfo = GetDBInfo "ECMDSIPR";
		}
		$strSQL = "SELECT * FROM Certs_Collection WHERE (Subject like '%ADIDBO216191%')";
		$Results = QueryDB $arrDBInfo[1] $arrDBInfo[2] $strSQL $False $arrDBInfo[3] $arrDBInfo[4] 180 $True;
		#$Results = QueryDB $arrDBInfo[1] $arrDBInfo[2] $strSQL $False $arrDBInfo[5] $arrDBInfo[6] 180 $True;
		if ($Results.Rows[0].Message -eq "Error"){
			Write-Host "Error running " $Results.Rows[0].Results;
		}else{
			$Results;
		}


		#Get top 10 Score card entries.  The last 10 made.
		$arrDBInfo = GetDBInfo "Score";
		$strSQL = "SELECT TOP 10 ticket, Type FROM Transactions ORDER BY date_time DESC";
		$Results = QueryDB $arrDBInfo[1] $arrDBInfo[2] $strSQL $False $arrDBInfo[3] $arrDBInfo[4] 180 $True;
		$Results = QueryDB $arrDBInfo[1] $arrDBInfo[2] $strSQL $False $arrDBInfo[5] $arrDBInfo[6] 180 $True;
		if ($Results.Rows[0].Message -eq "Error"){
			Write-Host "Error running " $Results.Rows[0].Results;
		}else{
			if ($Results.Rows[0].Message -eq "Success"){
				Write-Host "Success running the following SQL command:`r`n" $Results.Rows[0].Results;
				Write-Host $Results.Rows[0].FieldName1;
				Write-Host $Results.Rows[0].FieldName2;
				Write-Host $Results.Rows[1].FieldName1;
				Write-Host $Results.Rows[1].FieldName2;
			}else{
				if ($Results.Rows[0].Message -ne "Success"){
					for ($intX=0; $intX -le $Results.Length; $intX++){
						#$OutPut = $OutPut + ($Results[$intX].ItemArray -join ",");
						Write-Host ($Results[$intX].ItemArray -join ", ");
					}
				}
			}
		}


		#http://sev17.com/2010/05/11/t-sql-tuesday-006-blobs-filestream-and-powershell/
		#http://www.techtalkz.com/microsoft-windows-powershell/511586-question-inserting-varbinary-sql-table-via-powershell.html
		#Upload a file to the DB
		$arrDBInfo = GetDBInfo "SRMDB";
		$strGUID = "49C0E1F5-E726-43C4-A435-11B6A603FD0E";		#ASCII GUID in DB
		$dteDateTime = ([System.DateTime]::Now).ToString();
		$hashFiles = @{"File_Binary" = "C:\Projects\PS-Scripts\PS-AScII.ps1"};
		$strSQL = "INSERT INTO SourceFiles(SourceDesc_GUID, Date_Up, File_Binary) VALUES ('$strGUID', '$dteDateTime', @File_Binary)"
		$Results = QueryDB $arrDBInfo[1] $arrDBInfo[2] $strSQL $True -Files $hashFiles;
		if ($Results.Rows[0].Message -eq "Error"){
			Write-Host "Error running " $Results.Rows[0].Results;
		}else{
			$Results;
		}
		#QueryDB $arrDBInfo[1] $arrDBInfo[2] "SELECT * FROM SourceFiles" $True;


		#Read a Binary file out of the DB, and write it to a file.
		$arrDBInfo = GetDBInfo "SRMDB";
		$strGUID = "DBE41647-7C0A-4F01-A23D-C3B17A3549B6";		#Change App [SourceFiles]
		$strSQL = "SELECT File_Binary FROM SourceFiles WHERE (GUID = '$strGUID')";
		$Results = QueryDB $arrDBInfo[1] $arrDBInfo[2] $strSQL $True;
		if ($Results.Rows[0].Message -ne "Error"){
			#$file = $Results.Rows[0].File_Binary;
			#[IO.File]::WriteAllBytes("C:\Projects\test.xls", $file)
			[IO.File]::WriteAllBytes("C:\Projects\test.xls", $Results.Rows[0].File_Binary)
		}

	}


	function GetDBInfo{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strWhatSystem
		)
		#Sets the DB variables (DBType, DBServer, DBName, LoginR, PassR, LoginW, PassW), and returns an array.
		#strWhatSystem = The DB info we are after.
			#Current values coded for are: "AgentActivity", "Score", "Sites" (Server Farm LookUp), "SRMDB", "CDR", "ECMD"

		#Set some defaults
		$strRawData = "";
		if (!(Get-Command "GetPathing" -ErrorAction SilentlyContinue)){
			#$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;
			#if (([String]::IsNullOrEmpty($ScriptDir)) -or ($Error)){
			if ([String]::IsNullOrEmpty($MyInvocation.MyCommand.Path)){
				$ScriptDir = (Get-Location).ToString();
			}
			else{
				$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;
			}
			if ((Test-Path ($ScriptDir + "\Common.ps1"))){
				. ($ScriptDir + "\Common.ps1")
			}
			else{
				. "\\nawesdnifs101v.nadsuswe.nads.navy.mil\NMCIISF02$\ITSS-Tools\PS-CFW\Common.ps1";
			}
		}

		$strConfigFile = ((GetPathing "CFW").Returns.Rows[0].Path);
		if ([String]::IsNullOrEmpty($strConfigFile)){
			$strConfigFile = "\\nawesdnifs101v.nadsuswe.nads.navy.mil\NMCIISF02$\ITSS-Tools\PS-CFW\MiscSettings.txt";
		}
		else{
			$strConfigFile = $strConfigFile + "MiscSettings.txt";
		}

		#Some defaults, incase the MiscSettings file can not be read.
		Switch ($strWhatSystem){
			"AgentActivity"{
				$strRawData = "c3RyREJUeXBlID0gbXNzcWwNCnN0ckRCU2VydmVyID0gTkFXRVNETklTUTcyVkJcU1E3MlZCSU5TVDAxDQpzdHJEQk5hbWUgPSBBZ2VudEFjdGl2aXR5DQpzdHJEQkxvZ2luUiA9IGFpb2RhdGFyZWFkZXINCnN0ckRCUGFzc1IgPSBDTVc2MTE2MWRhdGFyZWFkZXINCnN0ckRCTG9naW5XID0gYWlvZGF0YQ0Kc3RyREJQYXNzVyA9IENNVzYxMTYxZGF0YQ==";
			}
			"CDR"{
				$strRawData = "c3RyREJUeXBlID0gbXNzcWwNCnN0ckRCU2VydmVyID0gbmFlYW5yZmt0bTAyDQpzdHJEQk5hbWUgPSBkYnBob2VuaXg1NTENCnN0ckRCTG9naW5SID0gaXNmdXNlcg0Kc3RyREJQYXNzUiA9IG4vYQ0Kc3RyREJMb2dpblcgPSBpc2Z1c2VyDQpzdHJEQlBhc3NXID0gbi9h";
			}
			"CDRDev"{
				$strRawData = "c3RyREJUeXBlID0gbXNzcWwNCnN0ckRCU2VydmVyID0gREFXRVNETklTUTA2dlxzcTA2dmluc3QwMQ0Kc3RyREJOYW1lID0gZGJwaG9lbml4NTUxDQpzdHJEQkxvZ2luUiA9IGlzZnVzZXINCnN0ckRCUGFzc1IgPSBpc2Z1c2VyDQpzdHJEQkxvZ2luVyA9IGlzZnVzZXINCnN0ckRCUGFzc1cgPSBuL2E=";
			}
			"CDRDevSIPR"{
				$strRawData = "";
			}
			"CDRSIPR"{
				$strRawData = "c3RyREJUeXBlID0gbXNzcWwNCnN0ckRCU2VydmVyID0gbmFlYW5yZmt0bTAxDQpzdHJEQk5hbWUgPSBkYnBob2VuaXg1NTENCnN0ckRCTG9naW5SID0gaXNmdXNlcg0Kc3RyREJQYXNzUiA9IG4vYQ0Kc3RyREJMb2dpblcgPSBpc2Z1c2VyDQpzdHJEQlBhc3NXID0gbi9h";
			}
			"ECMD"{
				$strRawData = "c3RyREJUeXBlID0gbXNzcWwNCnN0ckRCU2VydmVyID0gTkFFQU5SRktTUTUzXFNRNTNJTlNUMDENCnN0ckRCTmFtZSA9IEVDTUQNCnN0ckRCTG9naW5SID0ga2JTaXRlQ29kZURCVXNlcg0Kc3RyREJQYXNzUiA9IEtCU2l0QENvZEBVc2VyMQ0Kc3RyREJMb2dpblcgPSBub25lDQpzdHJEQlBhc3NXID0gbm9uZQ==";
			}
			"ECMDSIPR"{
				$strRawData = "c3RyREJUeXBlID0gbXNzcWwNCnN0ckRCU2VydmVyID0gbm1jaW5yZmtzcTAyXHNxMDJpbnN0MDENCnN0ckRCTmFtZSA9IEVDTUQNCnN0ckRCTG9naW5SID0ga2JTaXRlQ29kZURCVXNlcg0Kc3RyREJQYXNzUiA9IEtCU2l0QENvZEBVc2VyMQ0Kc3RyREJMb2dpblcgPSBub25lDQpzdHJEQlBhc3NXID0gbm9uZQ==";
			}
			"Score"{
				#Same as "AgentActivity".
				$strRawData = "c3RyREJUeXBlID0gbXNzcWwNCnN0ckRCU2VydmVyID0gTkFXRVNETklTUTcyVkJcU1E3MlZCSU5TVDAxDQpzdHJEQk5hbWUgPSBBZ2VudEFjdGl2aXR5DQpzdHJEQkxvZ2luUiA9IGFpb2RhdGFyZWFkZXINCnN0ckRCUGFzc1IgPSBDTVc2MTE2MWRhdGFyZWFkZXINCnN0ckRCTG9naW5XID0gYWlvZGF0YQ0Kc3RyREJQYXNzVyA9IENNVzYxMTYxZGF0YQ==";
			}
			"Sites"{
				$strRawData = "c3RyREJUeXBlID0gbXNzcWwNCnN0ckRCU2VydmVyID0gTkFFQU5SRktTUTc1VkFcU1E3NVZBSU5TVDAxDQpzdHJEQk5hbWUgPSBTaXRlQ29kZXMNCnN0ckRCTG9naW5SID0gS0J1c2VyDQpzdHJEQlBhc3NSID0ga2M1JHNxMDI=";
			}
			"SRMDB"{
				$strRawData = "c3RyREJUeXBlID0gbXNzcWwNCnN0ckRCU2VydmVyID0gTkFXRVNETklTUTcyVkJcU1E3MlZCSU5TVDAxDQpzdHJEQk5hbWUgPSBTUk1fQXBwc19Ub29scw0Kc3RyREJMb2dpblIgPSBTUk1fQXBwc19Ub29sc19XRk0NCnN0ckRCUGFzc1IgPSAhU1JNX0FwcHNfVG9vbHNfV0ZNNjkNCnN0ckRCTG9naW5XID0gU1JNX0FwcHNfVG9vbHMNCnN0ckRCUGFzc1cgPSAhU1JNX0FwcHNfVG9vbHM2OQ==";
			}
		}

		$Error.Clear();
		if (Test-Path -Path $strConfigFile){
			foreach ($strLine in [System.IO.File]::ReadAllLines($strConfigFile)) {
				if ($strLine.StartsWith($strWhatSystem)){
					$strRawData = $strLine.SubString($strLine.IndexOf("=") + 1).Trim();
					break;
				}
			}
		}

		$strDecode = [System.Convert]::FromBase64String($strRawData);
		$strDecode = [System.Text.Encoding]::ASCII.GetString($strDecode);
		$arrDecode = $strDecode.Split("`r`n");

		foreach ($strEntry In $arrDecode){
			if (($strEntry -ne "") -and ($strEntry -ne $null)){
				#$strKey = Trim(Left($strEntry, ($strEntry.IndexOf("=") - 1)));
				$strKey = $strEntry.SubString(0, $strEntry.IndexOf("=") - 1).Trim();
				#$strVal = Trim(Mid($strEntry, ($strEntry.IndexOf("=") + 1)));
				$strVal = $strEntry.SubString($strEntry.IndexOf("=") + 1).Trim();

				Switch ($strKey){
					"strDBType"{
						$strDBType = $strVal
					}
					"strDBServer"{
						$strDBServer = $strVal
					}
					"strDBName"{
						$strDBName = $strVal
					}
					"strDBLoginR"{
						$strDBLoginR = $strVal
					}
					"strDBPassR"{
						if (($strWhatSystem -eq "Score") -or ($strWhatSystem -eq "AgentActivity")){
							$strDBPassR = "@!0" + $strVal
						}else{
							$strDBPassR = $strVal
						}
					}
					"strDBLoginW"{
						$strDBLoginW = $strVal
					}
					"strDBPassW"{
						if (($strWhatSystem -eq "Score") -or ($strWhatSystem -eq "AgentActivity")){
							$strDBPassW = "@!0" + $strVal
						}else{
							$strDBPassW = $strVal
						}
					}
				}
			}
		}

		$arrDBInfo = @($strDBType, $strDBServer, $strDBName, $strDBLoginR, $strDBPassR, $strDBLoginW, $strDBPassW);

		return $arrDBInfo;
	}

	function Prep4ScoreCard{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strCOI,  
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strSource, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strTeam, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strAssignment, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strTicketNum, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strAction, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strCTI, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strSummary, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Int]$intQuant, 
			[ValidateNotNull()][Parameter(Mandatory=$False)]$dteStart, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strToolName, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strQuoteNum, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strBONum
		)
		#Returns the SQL command that should be used.
			#Needs to be updated to use the SP Andrew created.
		#Should pass in the above with info.
		#strCOI = Domain/Network.
		#strSource = The source that is initiating this work.  i.e "Ticket", "Email"
		#strTeam = i.e. SRM
		#strAssignment = i.e. UA
		#strTicketNum = A Ticket #.
		#strAction = The Action being done. i.e. "Disable", "Create Account".   	[taskDesc]
		#strCTI = The CTI/CAS/CI of the ticket being worked.						[cas]
		#strSummary = The Ticket Summary/Short Desc.
		#intQuant = The Quantity on the Ticket.
		#dteStart = The time the work started.
		#strToolName = The name of the Tool doing the work.
		#$strQuoteNum = The Quote #.
		#$strBONum = BuildOut #.

		if (($strSource -eq "") -or ($strSource -eq $null)){
			$strSource = "Ticket";
		}
		if (($strTeam -eq "") -or ($strTeam -eq $null)){
			$strTeam = "SRM";
		}
		if (($strAssignment -eq "") -or ($strAssignment -eq $null)){
			$strAssignment = "UA";
		}
		if (($intQuant -eq "") -or ($intQuant -lt 0) -or ($intQuant -eq $null)){
			$intQuant = 0;
		}
		if (($strToolName -eq "") -or ($strToolName -eq $null)){
			#$strToolName = "PS GUI";
			$strToolName = $MyInvocation.MyCommand.Name;
			$strToolName = $strToolName.Replace("PS-", "");
			$strToolName = $strToolName.Replace(".ps1", "");
		}
		$strCI = $strCTI;
		if (($strCTI -ne "") -and ($strCTI -ne $null)){
			$strCTI = $strCTI.Replace("Service Request", "SR");
			$strCTI = $strCTI.Replace("Premier Support", "PS");
			$strCTI = $strCTI.Replace("User Account Services", "UAS");
		}else{
			$strCTI = "No CTI/CAS/CI";
			$strCI = $strCTI;
		}
		if (($strAction -eq "") -or ($strAction -eq $null)){
			$strAction = "SRM Work";
		}
		if ([String]::IsNullOrEmpty($strCOI)){
			$strCOI = $env:UserDomain;
			$strCOI = "USN";
		}
		if ([String]::IsNullOrEmpty($strTicketNum)){
			$strTicketNum = "No-Num";
		}
		if ([String]::IsNullOrEmpty($strBONum)){
			$strBONum = "No-Num";
		}
		if ([String]::IsNullOrEmpty($strQuoteNum)){
			$strQuoteNum = "No-Num";
		}
		if ([String]::IsNullOrEmpty($strSummary)){
			$strSummary = "No Title/Summary Provided";
		}

		#Make sure the values are not too long
		if ($strCOI.Length -gt 4){
			$strCOI = $strCOI.SubString(0, 4)
		}
		if ($strSource.Length -gt 20){
			$strSource = $strSource.SubString(0, 20)
		}
		if ($strTeam.Length -gt 12){
			$strTeam = $strTeam.SubString(0, 12)
		}
		if ($strAssignment.Length -gt 20){
			$strAssignment = $strAssignment.SubString(0, 20)
		}
		if ($strTicketNum.Length -gt 128){
			$strTicketNum = $strTicketNum.SubString(0, 128)
		}
		if ($strQuoteNum.Length -gt 128){
			$strQuoteNum = $strQuoteNum.SubString(0, 128)
		}
		if ($strBONum.Length -gt 128){
			$strBONum = $strBONum.SubString(0, 128)
		}
		if ($strToolName.Length -gt 20){
			$strToolName = $strToolName.SubString(0, 20)
		}
		if ($strCTI.Length -gt 70){
			$strCTI = $strCTI.SubString(0, 70)
		}
		if ($strAction.Length -gt 70){
			$strAction = $strAction.SubString(0, 70)
		}
		if ($strSummary.Length -gt 128){
			$strSummary = $strSummary.SubString(0, 128)
		}
		$intTimeOffset = [Int][System.DateTime]::Now.Hour - [Int]([System.DateTime]::Now).ToUniversalTime().Hour;
		if ([Int][System.DateTime]::Now.Day -lt [Int]([System.DateTime]::Now).ToUniversalTime().Day){
			$intTimeOffset = $intTimeOffset - 24;
		}
		if (($dteStart -eq "") -or ($dteStart -eq $null) -or ($dteStart -eq 0)){
			$intHandleTime = 0;
			#$strSQL = $strSQL + "0"																							#handle_time (minutes)
		}else{
			$intHandleTime = ([System.DateTime]::Now - $dteStart).TotalMinutes;
			$intHandleTime = [Math]::Round($intHandleTime, 2);
			#$strSQL = $strSQL + "'" + $intHandleTime + "'"																		#handle_time (minutes)
		}

		#https://blogs.technet.microsoft.com/heyscriptingguy/2014/08/03/weekend-scripter-a-hidden-gem-in-the-powershell-ocean-get-pscallstack/
		$objCallStack = Get-PSCallStack | Select-Object -Property *;
		$strLastCmd = $objCallStack[1].Command;			#0 is this routine.
		##To get the first command called do the following:
		##if ($objCallStack.Count -eq 1){
		#	$strLastCmd = $objCallStack[0].Command;
		##}
		##else{
		#	$strFirstCmd = $objCallStack[($objCallStack.Count - 1)].Command;
		##}

		#The SP SQL statement:
		if ($strLastCmd -eq "RecordTransaction"){
			$strSQL = "GetSP_spSetTransaction '" + ([Environment]::MachineName.SubString(2, 4)) + "'," + $intTimeOffset + ",'" + $strCOI + "','" + $strSource + "','" + $strTeam + "','" + $strAssignment + "','" + $strToolName + "','" + $strCI + "','" + $strAction + "'," + $intQuant + "," + $intHandleTime + ",'','" + $strTicketNum + "','" + $strQuoteNum + "','" + $strBONum + "','" + $strSummary + "'";
		}
		else{
			[String]$strSQL = "INSERT INTO Transactions (login_name, date_time, UTC_date_time, UTC_Offset, Zone, site, coi, Source, Team, Assignment, ticket, QuoteNumber, BuildOutNum, Type, cas, CI, taskDesc, taskRef, title, res, QTY, credit_time, handle_time) VALUES ("
			#[Environment]::UserDomainName
			$strSQL = $strSQL + "'" + [Environment]::UserName + "', "   														#login_name
			$strSQL = $strSQL + "'" + [System.DateTime]::Now + "', "      														#date_time
			$strSQL = $strSQL + "'" + (([System.DateTime]::Now).ToUniversalTime()).ToString() + "', "      						#UTC_date_time
			$strSQL = $strSQL + "'" + $intTimeOffset + "', "      																#UTC_Offset
			$strSQL = $strSQL + "'" + (-join ([System.TimeZoneInfo]::Local.Id.Split(" ") | Foreach-Object {$_[0]})) + "', "		#Zone
			$strSQL = $strSQL + "'" + [Environment]::MachineName.SubString(2, 4) + "', "   										#site
			$strSQL = $strSQL + "'" + $strCOI + "', " 																			#coi
			$strSQL = $strSQL + "'" + $strSource + "', "																		#Source
			$strSQL = $strSQL + "'" + $strTeam + "', "																			#Team
			$strSQL = $strSQL + "'" + $strAssignment + "', "																	#Assignment
			$strSQL = $strSQL + "'" + $strTicketNum + "', "																		#Ticket
			$strSQL = $strSQL + "'" + $strQuoteNum + "', "																		#QuoteNumber
			$strSQL = $strSQL + "'" + $strBONum + "', "																			#BuildOutNum
			$strSQL = $strSQL + "'" + $strToolName + "', "																		#type
			$strSQL = $strSQL + "'" + $strCTI + "', " 					   														#cas  (CTI or Category.Area.Sub-Area)
			$strSQL = $strSQL + "'" + $strCI + "', " 					   														#CI
			$strSQL = $strSQL + "'" + $strAction + "', "																		#taskDesc
			$strSQL = $strSQL + "'" + "0" + "', "                 																#taskRef
			$strSQL = $strSQL + "'" + $strSummary + "', "																		#title
			$strSQL = $strSQL + "'" + "" + "', "                        														#res
			$strSQL = $strSQL + "" + $intQuant + ", "              																#QTY
			$strSQL = $strSQL + "" + "0" + ", "               																	#credit_time
			$strSQL = $strSQL + "" + $intHandleTime + ""																		#handle_time (minutes)
			$strSQL = $strSQL + ")"
		}

		return [String]$strSQL;

	}

	function QueryDB{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$Server, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$DataBase, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$SQL, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][Boolean]$IntSec, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$User, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$Pass, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Int]$TimeOut = 180, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Boolean]$ForceTableRet = $True, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Hashtable]$Files
		)
		#http://irisclasson.com/2013/10/16/how-do-i-query-a-sql-server-db-using-powershell-and-how-do-i-filter-format-and-output-to-a-file-stupid-question-251-255/
		#Returns a System.Data.DataTable of results  (converted to an Array).
		#$objResults.Rows[0].Message = "Error", if error.    $objResults.Rows[0].Results = The error message.

		#$Server = The SQL Server name.
		#$DataBase = The DB name.
		#$SQL = The SQL command to run.  For Stored Procedures that return a result set, prepend "GetSP_" to the SP command, unless it starts "sp_".
		#$IntSec = Use Integrated Security (True or False).
		#$User = Username if NOT Integrated Security.
		#$Pass = Password if NOT Integrated Security.
		#$TimeOut = The TimeOut period to use.  180 sec is the default.
		#$ForceTableRet = Force the return to NOT convert the System.Data.DataTable to an Array.
			#http://www.vistax64.com/powershell/36882-how-return-specific-type-function.html
		#$Files = Hash array of the field and full path to any files to upload.  (i.e.  $Files = @{"File_Binary" = "C:\Projects\PS-Scripts\PS-AScII.ps1"}; ).
			#The SQL statement should look something like (notice the at sign "@"):
				#$strSQL = "INSERT INTO SourceFiles(SourceDesc_GUID, Date_Up, File_Binary) VALUES ('$strGUID', '$dteDateTime', @File_Binary)"
				#$strSQL = "UPDATE SourceFiles SET File_Binary = @File_Binary WHERE (GUID = '$strGUID')"
		#Should add $strSysType = "MSSQL", "Oracle", "Access".

		$strServer = $Server;
		$strDataBase = $DataBase;
		$strUser = $User;
		$strPass = $Pass;
		$strSQL = $SQL;

		#http://rahmanagoro.wordpress.com/2010/08/26/powershell-secret-timeout-running-sql-from-powershell-v1/
		#Indicates that the default time out, if not specified, is 30 sec.
		if (($TimeOut -lt 0) -or ($TimeOut -eq "") -or ($TimeOut -eq $null)){
			$TimeOut = 180;
		}

		$bolSP = $False;
		if (($strSQL.StartsWith("GetSP_")) -or ($strSQL.StartsWith("SP_")) -or ($strSQL.StartsWith("sp_"))){
			if ($strSQL.StartsWith("GetSP_")){
				$strSQL = $strSQL.SubString(6);
			}
			$bolSP = $True;
		}

		$objTable = New-Object System.Data.DataTable;

		if ($IntSec -eq $False){
			$strConStr = "Server=$strServer; Database=$strDataBase; uid=$strUser; pwd=$strPass; Integrated Security=False;";
		}
		else{
			$strConStr = "Server=$strServer; Database=$strDataBase; Integrated Security=True;";
		}

		$Error.Clear();
		$objCon = New-Object System.Data.SqlClient.SqlConnection;
		$objCon.ConnectionString = $strConStr;
		Try{$objCon.Open();}Catch{}

		if (($Error.Count -gt 0) -or ($Error)){
			#$objTable = New-Object System.Data.DataTable;
			$col1 = New-Object System.Data.DataColumn Message,([String]);
			$col2 = New-Object System.Data.DataColumn Results,([String]);
			$objTable.columns.add($col1);
			$objTable.columns.add($col2);
			$row = $objTable.NewRow();
			$row.Message = "Error";
			#$row.Results = $Error;
			#$row.Results = $Error[0].ToString();
			$row.Results = "Error";
			if (($Error.Count -gt 1)){
				foreach ($MsgError in $Errors){
					$row.Results = $row.Results + "`r`n" + "Error $($MsgError.Number), Line $($MsgError.Line): $($MsgError.Message)";
				}
			}
			else{
					$row.Results = $row.Results + "`r`n" + $Error;
			}
			$objTable.Rows.Add($row);
		}
		else{
			$objCommand = New-Object System.Data.SqlClient.SqlCommand;
			$objCommand = $objCon.CreateCommand();
			$objCommand.CommandTimeout = $TimeOut;		#Seconds
			$objCommand.CommandText = $strSQL;

			if (($strSQL.IndexOf("@") -gt 1) -and (($Files -ne "") -and ($Files -ne $null))){
				foreach ($sEntry in $Files.Keys){
					if (($sEntry -ne "") -and ($sEntry -ne $null) -and ($Files.$sEntry -ne "") -and ($Files.$sEntry -ne $null)){
						#Read in the file(s)
						$objFile = [System.IO.File]::OpenRead($Files.$sEntry);
						$strFileByteArr = New-Object System.Byte[] $objFile.Length;
						$objResult = $objFile.Read($strFileByteArr, 0, $objFile.Length);
						$objResult = $objFile.Close();

						if (($objFile.Length -gt 0) -and ($strFileByteArr -ne "") -and ($strFileByteArr -ne $null)){
							#$objCommand.Parameters.Add("@File_Binary", $strFileByteArr);
							#	#$objCommand.Parameters.Add("@File_Binary", [System.Data.SqlDbType]"VarBinary", $buffer.Length);
							#	#$objCommand.Parameters["@File_Binary"].Value = $buffer;
							$objCommand.Parameters.Add("@" + $sEntry, $strFileByteArr);
						}
					}
				}
			}

			$Error.Clear();
			$objResult = $null;
			#The Try() below causes a SELECT query to error when doing the Load [.Load($objResult)] of data into the DataTable.
			#$objResult = Try{$objCommand.ExecuteReader();}Catch{}
			$objResult = $objCommand.ExecuteReader();

			if (($Error.Count -gt 0) -or ($Error)){
				#$objTable = New-Object System.Data.DataTable;
				$col1 = New-Object System.Data.DataColumn Message,([String]);
				$col2 = New-Object System.Data.DataColumn Results,([String]);
				$objTable.columns.add($col1);
				$objTable.columns.add($col2);
				$row = $objTable.NewRow();
				$row.Message = "Error";
				#$row.Results = $Error;
				#$row.Results = $Error[0].ToString();
				$row.Results = "Error";
				if (($Error.Count -gt 1)){
					foreach ($MsgError in $Errors){
						$row.Results = $row.Results + "`r`n" + "Error $($MsgError.Number), Line $($MsgError.Line): $($MsgError.Message)";
					}
				}
				else{
					$row.Results = $row.Results + "`r`n" + $Error;
				}
				$objTable.Rows.Add($row);
			}
			else{
				if (($strSQL.StartsWith("SELECT")) -or ($bolSP -eq $True)){
					#$objTable = New-Object System.Data.DataTable;
					$objTable.Load($objResult);
					#Should check if $objResult has results.
					#if no error, and no results, add "message" and "results" with "Success". 
					if (($Error.Count -gt 0) -or ($Error)){
						#$objTable = New-Object System.Data.DataTable;
						$col1 = New-Object System.Data.DataColumn Message,([String]);
						$col2 = New-Object System.Data.DataColumn Results,([String]);
						$objTable.columns.add($col1);
						$objTable.columns.add($col2);
						$row = $objTable.NewRow();
						$row.Message = "Error";
						#$row.Results = $Error;
						#$row.Results = $Error[0].ToString();
						$row.Results = "Error";
						if (($Error.Count -gt 1)){
							foreach ($MsgError in $Errors){
								$row.Results = $row.Results + "`r`n" + "Error $($MsgError.Number), Line $($MsgError.Line): $($MsgError.Message)";
							}
						}
						else{
							$row.Results = $row.Results + "`r`n" + $Error;
						}
						$objTable.Rows.Add($row);
					}
				}
				else{
					#$objTable = New-Object System.Data.DataTable;
					$col1 = New-Object System.Data.DataColumn Message,([String]);
					$col2 = New-Object System.Data.DataColumn Results,([String]);
					$objTable.columns.add($col1);
					$objTable.columns.add($col2);
					$row = $objTable.NewRow();
					$row.Message = "Success";
					$row.Results = $strSQL;
					$objTable.Rows.Add($row);
				}
			}

			$objCon.Close();
		}

		if ($ForceTableRet -eq $True){
			return ,$objTable;
		}
		else{
			#Return a datatable in an array. (PS default, yuck.)
			return $objTable;
		}
	}

	function RecordTransaction{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strCOI,  
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strSource, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strTeam, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strAssignment, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strTicketNum, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strAction, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strCTI, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strSummary, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Int]$intQuant, 
			[ValidateNotNull()][Parameter(Mandatory=$False)]$dteStart = 0, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strToolName, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strQuoteNum, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strBONum
		)
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with the paramaters passed in.
			#$objReturn.Results		= $True or $False.
			#$objReturn.Message		= "Success" or the error message.
			#$objReturn.Returns		= The recordset, if any.
		#strCOI = Domain/Network.
		#strSource = The source that is initiating this work.  i.e "Ticket", "Email"
		#strTeam = i.e. SRM
		#strAssignment = i.e. UA
		#strTicketNum = A Ticket #.
		#strAction = The Action being done. i.e. "Disable", "Create Account".   	[taskDesc]
		#strCTI = The CTI/CAS of the ticket being worked.							[cas]
		#strSummary = The Ticket Summary/Short Desc.
		#intQuant = The Quantity on the Ticket.
		#dteStart = The time the work started.
		#strToolName = The name of the Tool doing the work.
		#$strQuoteNum = The Quote #.
		#$strBONum = BuildOut #.

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
			Results = $True
			Message = "Success";
			Returns = "rs"
		}
		#Assume success

		$arrDBInfo = GetDBInfo "Score";
		$strSQL = Prep4ScoreCard $strCOI $strSource $strTeam $strAssignment $strTicketNum $strAction $strCTI $strSummary $intQuant $dteStart $strToolName $strQuoteNum $strBONum;

		$Error.Clear();
		$objResults = QueryDB $arrDBInfo[1] $arrDBInfo[2] $strSQL $True;
		if (($objResults.Rows[0].Message -eq "Error") -or ($Error)){
			if ($Error){
				$strMessage = "Error writing to ScoreCard DB.`r`n" + $Error;
			}else{
				$strMessage = $objResults.Rows[0].Message + " writing to ScoreCard DB.`r`n" + $objResults.Rows[0].Results;
			}
			#MsgBox $strMessage "ScoreCard DB error";
			$strMessage = "`r`n" + ("-" * 100) + "`r`n" + $strMessage + "`r`n";
			$strMessage = $strMessage + $strSQL + "`r`n";

			$objReturn.Results = $False;
			$objReturn.Message = $strMessage;
		}
		$objReturn.Returns = $objResults;

		return $objReturn;
	}
