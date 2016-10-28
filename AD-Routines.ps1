###########################################
# Updated Date:	28 October 2016
# Purpose:		Provide a central location for all the PowerShell Active Directory routines.
# Requirements: For the PInvoked Code .NET 4+ is required.
#				CheckNameAvail() requires isNumeric() from Common.ps1, and optionally MsgBox() from Forms.ps1.
##########################################

<# ---=== Change Log ===---
	#Changes for 28 June 2016:
		#Added Change Log.
	#29 Sept 2016
		#Added code to be able to read/check TSProfile properties. (line 4463)
	#14 Oct 2016
		#Routine formatting updates.
	#28 Oct 2016
		#Routine documentation templates.
#>



	function TestRoutine{

		. C:\Projects\PS-CFW\AD-Routines.ps1;
		. C:\Projects\PS-CFW\Common.ps1;

		#$InitializeDefaultDrives=$False;
		#if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};



		#. C:\Projects\PS-CFW\Exchange.ps1;
		##if (!(Get-Command "Get-Recipient" -ErrorAction SilentlyContinue)){
		#if ((!(Get-Command "Get-Recipient" -ErrorAction SilentlyContinue)) -or (!(Get-Command "SetupConn" -ErrorAction SilentlyContinue))){
		#	$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;
		#	if (([String]::IsNullOrEmpty($ScriptDir)) -or ($Error)){
		#		$ScriptDir = (Get-Location).ToString();
		#	}
		#	if (Test-Path ($ScriptDir + "\Exchange.ps1")){
		#		. ($ScriptDir + "\Exchange.ps1")
		#	}
		#	else{
		#		if (Test-Path ($ScriptDir + "\..\PS-CFW\Exchange.ps1")){
		#			. ($ScriptDir + "\..\PS-CFW\Exchange.ps1")
		#		}
		#	}
		#}

		#$UserName = "margaret.matthews";
		#$bolInUse = $False;
		##Need to do each domain
		#foreach ($strDom in $arrDom){
		#	#SetupConn "e" "d";
		#	SetupConn $strDom "d";
		#	if (Get-Recipient -Identity ($UserName + "@navy.mil") -ErrorAction SilentlyContinue){
		#		#in use.
		#		$bolInUse = $True;
		#		break;
		#	}
		#}



		#---=== Test 1 ===---
		<#
		$UserName = "margaret.matthews";		# "margaret.toelken" has the following email address "margaret.matthews@navy.mil"
		#$UserName = "redirect.test";
		$strDomain = "nadsuswe";
		$strDomain = "nadsusea";
		$strFilter = "(&(objectCategory=user)(proxyAddresses=*" + $UserName + "*))";
		#$strFilter = "(&(objectCategory=user)(EDIPI=*" + "EDIPI#" + "*))";
		#$strFilter = "(&(objectCategory=user)(UserPrincipalName=*" + "redirect.test@mil" + "*))";
		#$strFilter = "(&(objectCategory=user)(UserPrincipalName=*" + "1158784609" + "*))";
		#$strFilter = "(&(objectCategory=user)(|(UserPrincipalName=*" + "1158784609" + "*)(EDIPI=*" + "1158784609" + "*)))";
		[Array]$arrDesiredProps = @("samAccountName", "name", "proxyAddresses", "mail", "EDIPI", "UserPrincipalName");
		$objResults = $null;

		(Get-Date).ToString("yyyy-MM-dd HH:mm:ss");
		$objResults = ADSearchADO $UserName $strDomain $strFilter $arrDesiredProps $True;
		(Get-Date).ToString("yyyy-MM-dd HH:mm:ss");
		#Run times (search proxyAddresses for username):
			#2 min 15 sec
			#1 min 51 sec
			#1 min 51 sec
			#1 min 44 sec
			#1 min 47 sec
			#1 min 55 sec
			#2 min 17 sec
		#Run times (search for EDIPI):
			#1 min 49 sec
		#Run times (search for UserPrincipalName):
			#2 min 30 sec
			#1 min 20 sec
			#1 min 42sec
			#0 min 15 sec		(nmci-isf)
			#0 min 15 sec		(nmci-isf)
		$objResults | FL;
		$objResults.Returns;
		$objResults.Returns | FL;

		$objResults.Returns[0].proxyaddresses;

		if ($objResults.Returns[0].proxyaddresses -Match $UserName){
			Write-Host "Found a match, so the Email in use already.";
		}
		#>
		#---=== Test 1 ===---


		#---=== Test 2 ===---
		$strName = "redirect.test";
		$strMid = "m";
		$bolInter = $False;
		$strEDIPI = "1158784609";

		$objRet = CheckNameAvail $strName $bolInter $strMid -strEDIPI $strEDIPI;
		$objRet | FL;

		$strNewName = $objRet.Returns;
		$objRet = CheckNameAvail $strNewName $bolInter $strMid;
		$objRet | FL;

		$strNewName = $objRet.Returns;
		$objRet = CheckNameAvail $strNewName $bolInter $strMid -bForceInc $True;
		$objRet | FL;

		#---=== Test 2 ===---
	}

	
	function AddUserToGroup{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$GroupName, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$UserName, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$DomainOrDC
		)
		#Adds a User/computer to a Group as a Member.
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters.
			#$objReturn.Results		= 0 or 1.  0 = Error, 1 = Success
			#$objReturn.Message		= A verbose message of the results (The error message).
		#$GroupName = The Group Name (SamAccountName) to add the user/computer to.
		#$UserName = The user/computer name to add (DistinguishedName / LDAP), or an AD Object.
		#$DomainOrDC = The domain name or DC name to use to do the work on.

		$InitializeDefaultDrives=$False;
		if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};

		#Setup the PSObject to return.
		#$strTemp = "AddUserToGroup(" + $GroupName + ", " + $UserName + ", " + $DomainOrDC + ")";
		#http://stackoverflow.com/questions/21559724/getting-all-named-parameters-from-powershell-including-empty-and-set-ones
		$CommandName = $PSCmdlet.MyInvocation.InvocationName;
		$ParameterList = (Get-Command -Name $CommandName).Parameters;
		$strTemp = "";
		foreach ($key in $ParameterList.keys){
			#$var = Get-Variable -Name $key -ErrorAction SilentlyContinue;
			$var = Get-Variable $key -ErrorAction SilentlyContinue;
			if ($var){$strTemp += "[$($var.name) = $($var.value)] ";}
		}
		$strTemp = $CommandName + "(" + $strTemp.Trim() + ")";
		$objReturn = New-Object PSObject -Property @{
			Name = $strTemp
			Results = 0
			Message = "Error"
		}

		#Check if the desired group(s) exists.
		$objGroup = $null;
		if (($DomainOrDC -eq "") -or ($DomainOrDC -eq $null)){
			#Need to get Domains.  GetDomains() requires "AD-Routines.ps1".
			if (!(Get-Command "GetDomains" -ErrorAction SilentlyContinue)){
				$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;
				if (([String]::IsNullOrEmpty($ScriptDir)) -or ($Error)){
					$ScriptDir = (Get-Location).ToString();
				}
				if ((Test-Path ($ScriptDir + "\AD-Routines.ps1"))){
					. ($ScriptDir + "\AD-Routines.ps1")
				}
			}
			$arrDomains = GetDomains $False $False;

			for ($intY = 0; $intY -lt $arrDomains.Count; $intY++){
				#Get-ADGroup finds Exchange Groups too
				$objGroup = $null;
				$objGroup =  $(Try {Get-ADGroup -Identity $GroupName -Server ($arrDomains[$intY].SubString($arrDomains[$intY].IndexOf("=") + 1).Trim()) -Properties *;} Catch {$null});
				#$objGroup =  $(Try {Get-ADGroup -Identity $GroupName -Server $strDomain;} Catch {$null});
				if (($objGroup -ne "") -and ($objGroup -ne $null)){
					$DomainOrDC = $arrDomains[$intY].SubString($arrDomains[$intY].IndexOf("=") + 1).Trim();
					break;
				}
			}
		}
		else{
			$objGroup =  $(Try {Get-ADGroup -Identity $GroupName -Server $DomainOrDC -Properties *;} Catch {$null});
		}

		#Check if found a group
		if (($objGroup -ne "") -and ($objGroup -ne $null)){
			#Have a group so add user/computer to it now.
				#SG
				#$GroupName = "W_USFF_GLFP_MIGRATION_GS01";
				#$DomainOrDC = "nadsusea";
				#SG-MailEnabled
				#$GroupName = "NMCI IT Service Support Tools Engineering";
				#$DomainOrDC = "nmci-isf";
			#if (($objGroup.GroupCategory -eq "Security") -and (($objGroup.mail -eq "") -or ($objGroup.mail -eq $null))){
			if ((($objGroup.mail -eq "") -or ($objGroup.mail -eq $null))){
				$strGroupDN = [String]($objGroup).DistinguishedName;
				$Error.Clear();
				Add-ADGroupMember -Identity $strGroupDN -Member $UserName -Server $DomainOrDC;
			}
			else{
				#DL
				#Import exchange commands for the DL actions.
				$Session = Get-PSSession | Select Name;
				if (($Session -ne "") -and ($Session -ne $null)){
					#Write-Host "have at least one session";
				}
				else{
					if (!(Get-Command "SetupConn" -ErrorAction SilentlyContinue)){
						$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;
						if (([String]::IsNullOrEmpty($ScriptDir)) -or ($Error)){
							$ScriptDir = (Get-Location).ToString();
						}
						if ((Test-Path ($ScriptDir + "\Exchange.ps1"))){
							. ($ScriptDir + "\Exchange.ps1")
						}
					}
					SetupConn "w" "Random";
				}

				if (([String]($objGroup).DistinguishedName -eq "") -or ([String]($objGroup).DistinguishedName -eq $null)){
					$strGroupDN = [String]$objGroup;
				}
				else{
					$strGroupDN = [String]($objGroup).DistinguishedName;
				}
				$Error.Clear();
				#Add-DistributionGroupMember $objGroup -Member $UserName -DomainController $DomainOrDC;
				Add-DistributionGroupMember -Identity $strGroupDN -Member $UserName -DomainController $DomainOrDC;
			}

			if ($Error){
				$objReturn.Results = 0;
				$strMessage = "Error, Could not add user/computer '" + $UserName + "' to Group '" + $GroupName + "'.`r`n";
				$strMessage = $strMessage + $Error + "`r`n";
				$strMessage = "`r`n" + ("-" * 100) + "`r`n" + $strMessage + "";
			}
			else{
				$objReturn.Results = 1;
				$strMessage = "Success";
			}
		}
		else{
			$objReturn.Results = 0;
			$strMessage = "Error, Could not find the Group '" + $GroupName + "'.`r`n"
		}
		$objReturn.Message = $strMessage;

		#return $strMessage;
		return ,$objReturn;
	}

	function CreateGroup{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$GroupName, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$Scope, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$OUPath, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$DomainOrDC, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$ManagedBy, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$Members, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$GroupType, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$GroupDisp, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$GroupAlias, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$GroupNotes = ""
		)
		#Description....
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters.
			#$objReturn.Results		= $True or $False.  $False = Error, $True = Success
			#$objReturn.Message		= A verbose message of the results (The error message).
			#$objReturn.Returns		= The SID of the newly created group.  But I am currently returning the Group Object.
		#$GroupName = The desired Group Name (SamAccountName) to create.
		#$Scope = "Universal", "Global", "DomainLocal".
		#$OUPath = The full OU path (LDAP) of where to create the Group.
		#$DomainOrDC = The Domain or DC to create the group on.
		#$ManagedBy = The user who Manages the Group. (Distinguished Name or SID)
		#$Members = The users to add to the Group while creating it.  Only works for Exchange Groups.
		#$GroupType = What Type of Group to create. "", "Distribution", "Mail-Security", "Security".   "Security" = Security Group (AD), "Mail-Security" = Mail Enabled Security Group, "Distribution" = Distribution List Group (Exchange 2010).
		#$GroupDisp = The Display Name to give the (DL) Group.
		#$GroupAlias = Email alias to use.
		#$GroupNotes = The notes to add to the Group.

		$InitializeDefaultDrives=$False;
		if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};

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
			Returns = ""
		}

		#Check if the Group exists.
		$objGroup = $null;
		if (($DomainOrDC -eq "") -or ($DomainOrDC -eq $null)){
			#Need to get Domains.  GetDomains() requires "AD-Routines.ps1".
			if (!(Get-Command "GetDomains" -ErrorAction SilentlyContinue)){
				$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;
				if (([String]::IsNullOrEmpty($ScriptDir)) -or ($Error)){
					$ScriptDir = (Get-Location).ToString();
				}
				if ((Test-Path ($ScriptDir + "\AD-Routines.ps1"))){
					. ($ScriptDir + "\AD-Routines.ps1")
				}
			}
			$arrDomains = GetDomains $False $False;

			for ($intY = 0; $intY -lt $arrDomains.Count; $intY++){
				#Get-ADGroup finds Exchange Groups too
				$objGroup = $null;
				$objGroup =  $(Try {Get-ADGroup -Identity $GroupName -Server ($arrDomains[$intY].SubString($arrDomains[$intY].IndexOf("=") + 1).Trim());} Catch {$null});
				#$objGroup =  $(Try {Get-ADGroup -Identity $GroupName -Server $strDomain;} Catch {$null});
				if (($objGroup -ne "") -and ($objGroup -ne $null)){
					break;
				}
			}
			#$objGroup;
		}
		else{
			$objGroup =  $(Try {Get-ADGroup -Identity $GroupName -Server $DomainOrDC;} Catch {$null});
		}

		#Check if found an existing Group
		if (($objGroup -ne "") -and ($objGroup -ne $null)){
			#Found an existing Group
			$objReturn.Results = $False;
			$strResults = "Error Found a Group named '" + $GroupName + "' already exists.`r`n" + ($objGroup.DistinguishedName)
		}
		else{
			#Check that the OU exists ($OUPath), and if $DomainOrDC is not set, set it.
			$objOUReturn = Check4OU $OUPath $DomainOrDC;
			if (($objOUReturn.Results -gt 1) -and ((($DomainOrDC -ne "") -and ($DomainOrDC -ne $null)))){
				if ($objOUReturn -Match $DomainOrDC){
					$objOUReturn.Results = 1;
					$objOUReturn.Returns = $DomainOrDC;
					$objOUReturn.Message = "";
				}
			}

			#if ($objOUReturn.Results -eq $True){
			if ($objOUReturn.Results -eq 1){
				if (($DomainOrDC -eq "") -or ($DomainOrDC -eq $null)){
					#If $DomainOrDC is not set, set it.
					$DomainOrDC = $objOUReturn.Returns[0];
				}
				#$DomainOrDC SHOULD be a DC.
				if ($DomainOrDC.Contains(".") -eq $False){
					$DomainOrDC = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $DomainOrDC))).RidRoleOwner.Name;
				}

				if ((($GroupDisp -eq "") -or ($GroupDisp -eq $null)) -and (($GroupType -eq "Mail-Security") -or ($GroupType -eq "Distribution") -or ($GroupName.StartsWith("M") -or ($GroupName.StartsWith("m"))))){
					$GroupDisp = $GroupName;
				}

				#Make sure the Display name requested follows the naming standards.
				if ($GroupDisp -eq $GroupName){
					$arrBeginnings = @("M_", "M-", "m_", "m-", "W_", "W-", "w_", "w-", "A_", "a_", "A-", "a-");
					$arrEndings = @("_US", "_us", "_UD", "_ud", "_GS", "_gs", "_GD", "_gd", "_LS", "_ls", "_LD", "_ld");

					#Remove #'s off the end.
					#IsNumeric() equivelant is -> [Boolean]([String]($x -as [int]))
					Do{
						$strLastChar = $GroupDisp.SubString($GroupDisp.Length - 1, 1);
						if ([Boolean]([String]($strLastChar -as [int])) -eq $True){
							#MsgBox "Was a number, so remove it and check again."
							$GroupDisp = $GroupDisp.SubString(0, $GroupDisp.Length - 1);
						}
					}Until ([Boolean]([String]($strLastChar -as [int])) -eq $False)

					#Remove beginning
					for ($intY = 0; $intY -lt $arrBeginnings.Length; $intY++){
						if ($GroupDisp.StartsWith($arrBeginnings[$intY]) -eq $True){
							$GroupDisp = $GroupDisp.SubString($arrBeginnings[$intY].Length);
							break;
						}
					}

					#Remove ending.  Also verify the Scope matches the name ending.
					for ($intY = 0; $intY -lt $arrEndings.Length; $intY++){
						if ($GroupDisp.ToLower().EndsWith($arrEndings[$intY]) -eq $True){
							$GroupDisp = $GroupDisp.SubString(0, ($GroupDisp.Length - ($arrBeginnings[$intY].Length + 1)));
							if (($Scope -ne "Universal") -and ($Scope -ne "Global") -and ($Scope -ne "DomainLocal")){
								if ($Scope.ToUpper() -ne $arrEndings[$intY].SubString(1,1).ToUpper()){
									#Should we really be changing the group scope, to match the name provided instead of what was passed in?
									#If it is NOT one of the knows, then yes.
									$Scope = $arrEndings[$intY].SubString(1,1).ToUpper();
								}
							}
							break;
						}
					}
				}

				#Make sure a valid Scope was provided.
				if (($Scope -ne "Universal") -and ($Scope -ne "Global") -and ($Scope -ne "DomainLocal")){
					switch ($Scope){
						"U"{
							$Scope = "Universal";
						}
						"L"{
							$Scope = "DomainLocal";
						}
						"D"{
							$Scope = "DomainLocal";
						}
						default{
							$Scope = "Global";
						}
					}
				}

				if (($GroupType -ne "Security") -and ((($GroupName.StartsWith("M")) -or ($GroupName.StartsWith("m"))) -or (($GroupName.EndsWith("D")) -or ($GroupName.EndsWith("d"))) -or ($GroupType -eq "Distribution") -or ($GroupType -eq "Mail-Security"))){
					#Need to import exchange commands for the DL actions.
					$Session = Get-PSSession | Select Name, State;
					if (($Session -eq "") -or ($Session -eq $null) -or ($Session.State -ne "Opened")){
						if (!(Get-Command "SetupConn" -ErrorAction SilentlyContinue)){
							$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;
							if (([String]::IsNullOrEmpty($ScriptDir)) -or ($Error)){
								$ScriptDir = (Get-Location).ToString();
							}
							if ((Test-Path ($ScriptDir + "\Exchange.ps1"))){
								. ($ScriptDir + "\Exchange.ps1")
							}
						}
						if ($Session.State -ne "Opened"){
							#CleanUpConn;
						}
						SetupConn "w" "Random";
					}
					else{
						#Write-Host "have at least one session";
						#if ($Session -is [array]){
						#	For ($i=0; $i -lt $Session.length; $i++){
						#		Write-Host $Session[$i].Name;
						#	}
						#}
						#else{
						#	$Session = (Get-PSSession).Name;
						#	Write-Host "Session is: " $Session;
						#}
					}

					if (($GroupAlias -eq "") -or ($GroupAlias -eq $null)){
						if (($GroupDisp -eq "") -or ($GroupDisp -eq $null)){
							$GroupAlias = $GroupName;
							$GroupDisp = $GroupName;
						}
						else{
							$GroupAlias = $GroupDisp;
						}
					}

					if (($GroupType -eq "Distribution") -or (($GroupName.EndsWith("D")) -or ($GroupName.EndsWith("d")))){
						#Create Exch Distrobution Group
						$strMessage = " - Exch Distrobution Group";
						$Error.Clear();
						#For the New-DistributionGroup cmdlet. The Type parameter specifies the group type created in Active Directory. The group's scope is always Universal. Valid values are Distribution or Security.
						if ($ManagedBy){
							if ($Members){
								#Has Manager & Has Users
									$strResults = New-DistributionGroup -Name $GroupName -Type "Distribution" -DomainController $DomainOrDC -DisplayName $GroupDisp -SamAccountName $GroupName -OrganizationalUnit $OUPath -Alias $GroupAlias -Notes $GroupNotes -ManagedBy $ManagedBy -Members $Members;
									#$objJobCode = [scriptblock]::create({param($strGroupName, $strOpsMaster, $strGroupDisplayName, $strPath, $strGroupAlias, $strGroupEmail, $strGroupNotes, $strManagedBy, $strGrpMembers); New-DistributionGroup -Name $strGroupName -Type "Distribution" -DomainController $strOpsMaster -DisplayName $strGroupDisplayName -SamAccountName $strGroupName -OrganizationalUnit $strPath -Alias $strGroupAlias -PrimarySmtpAddress $strGroupEmail -Notes $strGroupNotes -ManagedBy $strManagedBy -Members $strGrpMembers;});
									#$arrArgs = @($GroupName, $strGroupDomain, $GroupDisp, $OUPath, $GroupAlias, ($GroupAlias + "@navy.mil"), $GroupNotes, $ManagedBy, $Members);
							}ellse{
								#Has Manager & NO Users
									$strResults = New-DistributionGroup -Name $GroupName -Type "Distribution" -DomainController $DomainOrDC -DisplayName $GroupDisp -SamAccountName $GroupName -OrganizationalUnit $OUPath -Alias $GroupAlias -Notes $GroupNotes -ManagedBy $ManagedBy;
									#$objJobCode = [scriptblock]::create({param($strGroupName, $strOpsMaster, $strGroupDisplayName, $strPath, $strGroupAlias, $strGroupEmail, $strGroupNotes, $strManagedBy); New-DistributionGroup -Name $strGroupName -Type "Distribution" -DomainController $strOpsMaster -DisplayName $strGroupDisplayName -SamAccountName $strGroupName -OrganizationalUnit $strPath -Alias $strGroupAlias -PrimarySmtpAddress $strGroupEmail -Notes $strGroupNotes -ManagedBy $strManagedBy;});
									#$arrArgs = @($GroupName, $strGroupDomain, $GroupDisp, $OUPath, $GroupAlias, ($GroupAlias + "@navy.mil"), $GroupNotes, $ManagedBy);
							}
						}
						else{
							if ($Members){
								#NO Manager & Has Users
									$strResults = New-DistributionGroup -Name $GroupName -Type "Distribution" -DomainController $DomainOrDC -DisplayName $GroupDisp -SamAccountName $GroupName -OrganizationalUnit $OUPath -Alias $GroupAlias -Notes $GroupNotes -Members $Members;
									#$objJobCode = [scriptblock]::create({param($strGroupName, $strOpsMaster, $strGroupDisplayName, $strPath, $strGroupAlias, $strGroupEmail, $strGroupNotes, $strGrpMembers); New-DistributionGroup -Name $strGroupName -Type "Distribution" -DomainController $strOpsMaster -DisplayName $strGroupDisplayName -SamAccountName $strGroupName -OrganizationalUnit $strPath -Alias $strGroupAlias -PrimarySmtpAddress $strGroupEmail -Notes $strGroupNotes -Members $strGrpMembers;});
									#$arrArgs = @($GroupName, $strGroupDomain, $GroupDisp, $OUPath, $GroupAlias, ($GroupAlias + "@navy.mil"), $GroupNotes, $Members);
							}
							else{
								#NO Manager & NO Users
									$strResults = New-DistributionGroup -Name $GroupName -Type "Distribution" -DomainController $DomainOrDC -DisplayName $GroupDisp -SamAccountName $GroupName -OrganizationalUnit $OUPath -Alias $GroupAlias -Notes $GroupNotes;
									#$objJobCode = [scriptblock]::create({param($strGroupName, $strOpsMaster, $strGroupDisplayName, $strPath, $strGroupAlias, $strGroupEmail, $strGroupNotes); New-DistributionGroup -Name $strGroupName -Type "Distribution" -DomainController $strOpsMaster -DisplayName $strGroupDisplayName -SamAccountName $strGroupName -OrganizationalUnit $strPath -Alias $strGroupAlias -PrimarySmtpAddress $strGroupEmail -Notes $strGroupNotes;});
									#$arrArgs = @($GroupName, $strGroupDomain, $GroupDisp, $OUPath, $GroupAlias, $strGroupFullEmail, $GroupNotes);
							}
						}
					}
					else{
						#Create Exch Mail Enabled Security Group
						$strMessage = " - Exch Mail Enabled Security Group";
						$Error.Clear();
						#For the New-DistributionGroup cmdlet. The Type parameter specifies the group type created in Active Directory. The group's scope is always Universal. Valid values are Distribution or Security.
						if ($ManagedBy){
							if ($Members){
								#Has Manager & Has Users
									$strResults = New-DistributionGroup -Name $GroupName -Type "Security" -DomainController $DomainOrDC -DisplayName $GroupDisp -SamAccountName $GroupName -OrganizationalUnit $OUPath -Alias $GroupAlias -Notes $GroupNotes -ManagedBy $ManagedBy -Members $Members;
									#$objJobCode = [scriptblock]::create({param($strGroupName, $strOpsMaster, $strGroupDisplayName, $strPath, $strGroupAlias, $strGroupEmail, $strGroupNotes, $strManagedBy, $strGrpMembers); New-DistributionGroup -Name $strGroupName -Type "Security" -DomainController $strOpsMaster -DisplayName $strGroupDisplayName -SamAccountName $strGroupName -OrganizationalUnit $strPath -Alias $strGroupAlias -PrimarySmtpAddress $strGroupEmail -Notes $strGroupNotes -ManagedBy $strManagedBy -Members $strGrpMembers;});
									#$arrArgs = @($GroupName, $strGroupDomain, $GroupDisp, $OUPath, $GroupAlias, ($GroupAlias + "@navy.mil"), $GroupNotes, $ManagedBy, $Members);
							}
							else{
								#Has Manager & NO Users
									$strResults = New-DistributionGroup -Name $GroupName -Type "Security" -DomainController $DomainOrDC -DisplayName $GroupDisp -SamAccountName $GroupName -OrganizationalUnit $OUPath -Alias $GroupAlias -Notes $GroupNotes -ManagedBy $ManagedBy;
									#$objJobCode = [scriptblock]::create({param($strGroupName, $strOpsMaster, $strGroupDisplayName, $strPath, $strGroupAlias, $strGroupEmail, $strGroupNotes, $strManagedBy); New-DistributionGroup -Name $strGroupName -Type "Security" -DomainController $strOpsMaster -DisplayName $strGroupDisplayName -SamAccountName $strGroupName -OrganizationalUnit $strPath -Alias $strGroupAlias -PrimarySmtpAddress $strGroupEmail -Notes $strGroupNotes -ManagedBy $strManagedBy;});
									#$arrArgs = @($GroupName, $strGroupDomain, $GroupDisp, $OUPath, $GroupAlias, ($GroupAlias + "@navy.mil"), $GroupNotes, $ManagedBy);
							}
						}
						else{
							if ($Members){
								#NO Manager & Has Users
									$strResults = New-DistributionGroup -Name $GroupName -Type "Security" -DomainController $DomainOrDC -DisplayName $GroupDisp -SamAccountName $GroupName -OrganizationalUnit $OUPath -Alias $GroupAlias -Notes $GroupNotes -Members $Members;
									#$objJobCode = [scriptblock]::create({param($strGroupName, $strOpsMaster, $strGroupDisplayName, $strPath, $strGroupAlias, $strGroupEmail, $strGroupNotes, $strGrpMembers); New-DistributionGroup -Name $strGroupName -Type "Security" -DomainController $strOpsMaster -DisplayName $strGroupDisplayName -SamAccountName $strGroupName -OrganizationalUnit $strPath -Alias $strGroupAlias -PrimarySmtpAddress $strGroupEmail -Notes $strGroupNotes -Members $strGrpMembers;});
									#$arrArgs = @($GroupName, $strGroupDomain, $GroupDisp, $OUPath, $GroupAlias, ($GroupAlias + "@navy.mil"), $GroupNotes, $Members);
							}
							else{
								#NO Manager & NO Users
									$strResults = New-DistributionGroup -Name $GroupName -Type "Security" -DomainController $DomainOrDC -DisplayName $GroupDisp -SamAccountName $GroupName -OrganizationalUnit $OUPath -Alias $GroupAlias -Notes $GroupNotes;
									#$objJobCode = [scriptblock]::create({param($strGroupName, $strOpsMaster, $strGroupDisplayName, $strPath, $strGroupAlias, $strGroupEmail, $strGroupNotes); New-DistributionGroup -Name $strGroupName -Type "Security" -DomainController $strOpsMaster -DisplayName $strGroupDisplayName -SamAccountName $strGroupName -OrganizationalUnit $strPath -Alias $strGroupAlias -PrimarySmtpAddress $strGroupEmail -Notes $strGroupNotes;});
									#$arrArgs = @($GroupName, $strGroupDomain, $GroupDisp, $OUPath, $GroupAlias, ($GroupAlias + "@navy.mil"), $GroupNotes);
							}
						}
					}
				}
				else{
					#if (($GroupName.StartsWith("W")) -or ($GroupName.StartsWith("w"))){
						#Create AD Security Group
						$strMessage = " - AD Security Group";
						$InitializeDefaultDrives=$False;
						if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};

						$Error.Clear();
						if ($ManagedBy){
							#Has Manager
								#$strResults = New-ADGroup -Name $GroupName -GroupScope $Scope -Server $DomainOrDC -SamAccountName $GroupName -Path $OUPath -ManagedBy $ManagedBy;
								if ($GroupType -eq "Distribution"){
									$strResults = New-ADGroup -Name $GroupName -GroupScope $Scope -GroupCategory "Distribution" -Server $DomainOrDC -SamAccountName $GroupName -Path $OUPath -ManagedBy $ManagedBy -OtherAttributes @{'info'=$GroupNotes};
								}
								else{
									$strResults = New-ADGroup -Name $GroupName -GroupScope $Scope -Server $DomainOrDC -SamAccountName $GroupName -Path $OUPath -ManagedBy $ManagedBy -OtherAttributes @{'info'=$GroupNotes};
								}
								#$objJobCode = [scriptblock]::create({param($strGroupName, $strGrpType, $strOpsMaster, $strPath, $strManagedBy); New-ADGroup -Name $strGroupName -GroupScope $strGrpType -Server $strOpsMaster -SamAccountName $strGroupName -Path $strPath -ManagedBy $strManagedBy;});
								#$arrArgs = @($GroupName, $strGroupScope, $strGroupDomain, $OUPath, $ManagedBy);
						}
						else{
							#NO Manager
								#$strResults = New-ADGroup -Name $GroupName -GroupScope $Scope -Server $DomainOrDC -SamAccountName $GroupName -Path $OUPath;
								if ($GroupType -eq "Distribution"){
									$strResults = New-ADGroup -Name $GroupName -GroupScope $Scope -GroupCategory "Distribution" -Server $DomainOrDC -SamAccountName $GroupName -Path $OUPath -OtherAttributes @{'info'=$GroupNotes};
								}
								else{
									$strResults = New-ADGroup -Name $GroupName -GroupScope $Scope -Server $DomainOrDC -SamAccountName $GroupName -Path $OUPath -OtherAttributes @{'info'=$GroupNotes};
									#$strResults = New-ADGroup -Name $GroupName -GroupScope $Scope -GroupCategory "Security" -Server $DomainOrDC -SamAccountName $GroupName -Path $OUPath -OtherAttributes @{'info'=$GroupNotes};
								}
								#$objJobCode = [scriptblock]::create({param($strGroupName, $strGrpType, $strOpsMaster, $strPath); New-ADGroup -Name $strGroupName -GroupScope $strGrpType -Server $strOpsMaster -SamAccountName $strGroupName -Path $strPath;});
								#$arrArgs = @($GroupName, $strGroupScope, $strGroupDomain, $OUPath);
						}
					#}
				}

				#New-ADGroup() has NO results if SUCCESS.
				#New-DistributionGroup() does return the Group Object.

				if ($Error){
					$objReturn.Results = $False;
					$objReturn.Message = "Error";
					$objReturn.Message = $objReturn.Message + $strMessage + "`r`n";
					$objReturn.Message = $objReturn.Message + $Error;
				}
				else{
					$objReturn.Results = $True;
					$objReturn.Message = "Success";
					$objReturn.Message = $objReturn.Message + $strMessage;
					if (($strResults -ne "") -and ($strResults -ne $null)){
						$objReturn.Returns = $strResults;
					}
				}
			}
			else{
				$objReturn.Results = $False;

				if ($objOUReturn.Results -gt 1){
					$strTemp = $objOUReturn.Returns[0];
					for ($intX = 1; $intX -lt $objOUReturn.Results; $intX++){
						$strTemp = $strTemp + ", " + $objOUReturn.Returns[$intX];
					}

					$strResults = "The OU provided, to create the group in, was found on multiple Domains.`r`n $strTemp";
				}
				else{
					#OU path provided does not exist.
					$strResults = "The OU path provided, to create the group in, could not be found found on any available Domains.";
				}

				$objReturn.Message = $strResults;
			}
		}

		return $objReturn;
	}

	function GetGroups{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)]$ADObject, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Array]$arrList, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bolRecurse = $False
		)
		#Description....
		#Return.....
		#$ADObject = An AD object, or the sAMAccountName (String) of the AD object to get.
		#$arrList = The Array, of strings, that will be updated/returned, that will have the list of Memberships $ADObject has.
		#$bolRecurse = Get the Groups any Groups are Members Of as well.

		#Based heavily on code from:
		#http://www.reich-consulting.net/2013/12/05/retrieving-recursive-group-memberships-powershell/
		#Which we got to from:
		#http://stackoverflow.com/questions/5072996/how-to-get-all-groups-that-a-user-is-a-member-of

		$InitializeDefaultDrives=$False;
		if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};

		if (($arrList -eq "") -or ($arrList -eq $null) -or ($arrList.GetType().IsArray -ne $True)){
			$arrList = @();
		}

		#Check if $ADObject is a String, or already an AD Object.
		if ($ADObject.GetType().FullName -eq "System.String"){
			if ($ADObject.Contains("\")){
				#$objADObject = (Get-ADUser -Identity $arrSplit[1] -Server $arrSplit[0] -Properties MemberOf, DistinguishedName, sAMAccountName | Select-Object MemberOf, DistinguishedName, sAMAccountName);
				$objADObject = $(Try {(Get-ADUser -Identity $ADObject.Split("\")[-1] -Server $ADObject.Split("\")[0] -Properties MemberOf, DistinguishedName, sAMAccountName | Select-Object MemberOf, DistinguishedName, sAMAccountName)} Catch {$null});
			}
			else{
				#$objADObject = (Get-ADUser -Identity $ADObject -Properties MemberOf, DistinguishedName, sAMAccountName | Select-Object MemberOf, DistinguishedName, sAMAccountName);
				$objADObject = $(Try {(Get-ADUser -Identity $ADObject -Properties MemberOf, DistinguishedName, sAMAccountName | Select-Object MemberOf, DistinguishedName, sAMAccountName)} Catch {$null});
			}

			if (($objADObject -eq $null) -or ($objADObject -eq "")){
				#Could not find an AD Object, check if it is a Group.
				if ($ADObject.Contains("\")){
					$objADObject = $(Try {Get-ADObject $ADObject.Split("\")[-1] -Server $ADObject.Split("\")[0] -Properties MemberOf, sAMAccountName, DistinguishedName} Catch {$null});
				}
				else{
					$objADObject = $(Try {Get-ADObject $ADObject -Properties MemberOf, sAMAccountName, DistinguishedName} Catch {$null});
				}
			}

			if (($objADObject -eq $null) -or ($objADObject -eq "")){
				#Could not find an AD Object, check if it is a Machine.
				if ($ADObject.Contains("\")){
					$objADObject = $(Try {Get-ADComputer -Identity $ADObject.Split("\")[-1] -Server $ADObject.Split("\")[0] -Properties MemberOf, sAMAccountName, DistinguishedName} Catch {$null});
				}
				else{
					$objADObject = $(Try {Get-ADComputer -Identity $ADObject -Properties MemberOf, sAMAccountName, DistinguishedName} Catch {$null});
				}
			}
		}
		#$arrList += [string] "-" + $objADObject.sAMAccountName;

		if (($objADObject -ne $null) -and ($objADObject -ne "")){
			foreach ($strGroup in $objADObject.MemberOf){
				$objGroup = "";
				#$objGroup = Get-ADObject $strGroup -Properties MemberOf, sAMAccountName, DistinguishedName;
				$objGroup = $(Try {Get-ADObject $strGroup -Properties MemberOf, sAMAccountName, DistinguishedName} Catch {$null});
				if (($objGroup -ne $null) -and ($objGroup -ne "")){
					if ($arrList -NotContains $objGroup.sAMAccountName){
						$arrList += $objGroup.sAMAccountName;
						if ($bolRecurse -eq $True){
							$arrList = GetGroups $objGroup $arrList $bolRecurse;
						}
					}
				}
				else{
					#Failed to find an AD object for these ones, just from the DistinguishedName.
					#So now we need to specify a DC to search against.
					$arrSplit = $strGroup -Split ',';
					for ($intX = 0; $intX -lt $arrSplit.Count; $intX++){
						if ($arrSplit[$intX].Contains("DC=")){
							$strDomain = $arrSplit[$intX].SubString(3);
							if (($strDomain -ne $null) -and ($strDomain -ne "")){
								break;
							}
						}
					}
					if ($strDomain -eq $null){
						$strDomain = "";
					}

					#$strRIDMaster = (Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
					Switch ($strDomain){
						"nadsuswe"{
							if ($txbRidW -ne $null){
								if ($txbRidW.Text -ne ""){
									$strRIDMasterW = $txbRidW.Text.Trim();
								}
								else{
									if (!(($strRIDMasterW -ne $null) -and ($strRIDMasterW -ne ""))){
										$strStatus = CheckRunSpaceJob "GetRIDMasterW" $global:objJobs;
										if ($strStatus -ne "Failed"){
											$strRIDMasterW = WaitForRunSpaceJob "GetRIDMasterW" $global:objJobs $txbRidW;
										}
										else{
											$strRIDMasterW = (Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
											$txbRidW.Text = $strRIDMasterW;
										}
									#}
									#else{
										#Have $strRIDMasterW already.
									}
								}
							}
							else{
								#This Function is probably running in the background.
								if (!(($strRIDMasterW -ne $null) -and ($strRIDMasterW -ne ""))){
									$strRIDMasterW = (Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
								#}
								#else{
									#Have $strRIDMasterW already.
								}
							}
							$strRIDMaster = $strRIDMasterW;
						}
						"nadsusea"{
							if ($txbRidE -ne $null){
								if ($txbRidE.Text -ne ""){
									$strRIDMasterE = $txbRidE.Text.Trim();
								}
								else{
									if (!(($strRIDMasterE -ne $null) -and ($strRIDMasterE -ne ""))){
										$strStatus = CheckRunSpaceJob "GetRIDMasterE" $global:objJobs;
										if ($strStatus -ne "Failed"){
											$strRIDMasterE = WaitForRunSpaceJob "GetRIDMasterE" $global:objJobs $txbRidE;
										}
										else{
											$strRIDMasterE = (Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
											$txbRidE.Text = $strRIDMasterE;
										}
									#}
									#else{
										#Have $strRIDMasterE already.
									}
								}
							}
							else{
								#This Function is probably running in the background.
								if (!(($strRIDMasterE -ne $null) -and ($strRIDMasterE -ne ""))){
									$strRIDMasterE = (Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
								#}
								#else{
									#Have $strRIDMasterE already.
								}
							}
							$strRIDMaster = $strRIDMasterE;
						}
						"pads"{
							if ($txbRidP -ne $null){
								if ($txbRidP.Text -ne ""){
									$strRIDMasterP = $txbRidP.Text.Trim();
								}
								else{
									if (!(($strRIDMasterP -ne $null) -and ($strRIDMasterP -ne ""))){
										$strStatus = CheckRunSpaceJob "GetRIDMasterP" $global:objJobs;
										if ($strStatus -ne "Failed"){
											$strRIDMasterP = WaitForRunSpaceJob "GetRIDMasterP" $global:objJobs $txbRidP;
										}
										else{
											$strRIDMasterP = (Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
											$txbRidP.Text = $strRIDMasterP;
										}
									#}
									#else{
										#Have $strRIDMasterP already.
									}
								}
							}
							else{
								#This Function is probably running in the background.
								if (!(($strRIDMasterP -ne $null) -and ($strRIDMasterP -ne ""))){
									$strRIDMasterP = (Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
								#}
								#else{
									#Have $strRIDMasterP already.
								}
							}
							$strRIDMaster = $strRIDMasterP;
						}
						"nmci-isf"{
							if ($txbRidN -ne $null){
								if ($txbRidN.Text -ne ""){
									$strRIDMasterN = $txbRidN.Text.Trim();
								}
								else{
									if (!(($strRIDMasterN -ne $null) -and ($strRIDMasterN -ne ""))){
										$strStatus = CheckRunSpaceJob "GetRIDMasterN" $global:objJobs;
										if ($strStatus -ne "Failed"){
											$strRIDMasterN = WaitForRunSpaceJob "GetRIDMasterN" $global:objJobs $txbRidN;
										}
										else{
											$strRIDMasterN = (Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
											$txbRidN.Text = $strRIDMasterN;
										}
									#}
									#else{
										#Have $strRIDMasterN already.
									}
								}
							}
							else{
								#This Function is probably running in the background.
								if (!(($strRIDMasterN -ne $null) -and ($strRIDMasterN -ne ""))){
									$strRIDMasterN = (Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
								#}
								#else{
									#Have $strRIDMasterN already.
								}
							}
							$strRIDMaster = $strRIDMasterN;
						}
						""{
							$strRIDMaster = (Get-ADDomain -ErrorAction SilentlyContinue).RIDMaster;
						}
						default{
							$strRIDMaster = (Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
						}
					}

					#$objGroup = Get-ADObject $strGroup -Server $strRIDMaster -Properties MemberOf, sAMAccountName, DistinguishedName;
					$objGroup = $(Try {Get-ADObject $strGroup -Server $strRIDMaster -Properties  MemberOf, sAMAccountName, DistinguishedName} Catch {$null});
					if (($objGroup -ne $null) -and ($objGroup -ne "")){
						#Now we found it.
						if ($arrList -NotContains $objGroup.sAMAccountName){
							$arrList += $objGroup.sAMAccountName;
							if ($bolRecurse -eq $True){
								$arrList = GetGroups $objGroup $arrList $bolRecurse;
							}
						}
					}
					else{
						#Still could not find it, so just return the DistinguishedName.
						for ($intX = 0; $intX -lt $arrSplit.Count; $intX++){
							if ($arrSplit[$intX].Contains("CN=")){
								$strGroup = $arrSplit[$intX].SubString(3);
								break;
							}
						}
						$arrList += $strGroup;
					}
				}
			}
		}
		else{
			$arrList += "Error Could not find $ADObject in AD.";
		}
		#$strPSCmds = $strPSCmds.Replace(", ", ",`r`n");

		#$arrList = $arrList | Sort-Object;

		return $arrList;
	}


	function ADSearchADO{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$Username, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strDomain, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strFilter, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Array]$arrDesiredProps = @("name"), 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bPSobj = $False
		)
		#Description....
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters.
			#$objReturn.Results		= # of objects found.
			#$objReturn.Message		= A verbose message of the results (The error message).
			#$objReturn.Returns		= The object(s) found.   (System.DirectoryServices.SearchResult)  or  (SearchResultCollection)  or  PowerShell Object(s).
		#$Username = The user name to search for, if $strFilter is NOT provided.
		#$strDomain = The domain to search for $Username on/in. (Defaults to GC).
			#i.e. "nadsuswe", or "DC=nadsusea,DC=nads,DC=navy,DC=mil", or "GC" (super fast) if the attribute gets replicated.
		#$strFilter = A custom LDAP search filter, instead of the default.   Default is  -->  "(&(objectCategory=user)(name=*" + $Username + "*))"
			#$strFilter = "(&(objectCategory=user))";
			#$strFilter = "(&(objectCategory=user)(mail=*" + $Username + "*))";
			#$strFilter = "(&(objectCategory=user)(proxyAddresses=*))";
		#$arrDesiredProps = A list of Properties you want returned instead of just "name" and "adsPath".  (adsPath is default w/ all options.)
		#$bPSobj = $True to return a [collection of] PowerShell Object(s) instead of a System.DirectoryServices.SearchResult.

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
			Message = "Success"
			Returns = ""
		}

		#https://technet.microsoft.com/en-us/library/ff730967.aspx

		$bolGC = $False;
		$Error.Clear();
		if (($strDomain -eq "") -or ($strDomain -eq $null)){
			#Can we use "rootDSE" with PS?
				#Nope.   :-(
			##$strDomain = "nadsuswe";
			##$objDomain = New-Object System.DirectoryServices.DirectoryEntry;
			#$strDomain = "LDAP://rootDSE";											#Looking like this does NOT work.
			##Maybe this will work.
			#$strDomain = ([ADSI]"LDAP://RootDse").configurationNamingContext;		#Looking like this does NOT work either.

			$bolGC = $True;
			#Jason's code....
			$strForestName = ([System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()).Name;
			$strDomain = [ADSI]"GC://$strForestName"; 
		}
		else{
			if ($strDomain -eq "GC"){
				$bolGC = $True;
				$strForestName = ([System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()).Name;
				$strDomain = [ADSI]"GC://$strForestName"; 
				#$Search = New-Object System.DirectoryServices.DirectorySearcher($strDomain);
			}
			else{
				#$objDomain = New-Object System.DirectoryServices.DirectoryEntry("LDAP://DC=nads,DC=navy,DC=mil");
				if ($strDomain.IndexOf("DC=") -eq 0){
					#$strDomain = "DC=" + $strDomain + ",DC=nads,DC=navy,DC=mil";
					$strDomain = ([System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $strDomain)))).name;
					$strDomain = "DC=" + $strDomain.Replace(".", ",DC=");
				}
				$strDomain = "LDAP://" + $strDomain;

				#$objDomain = New-Object System.DirectoryServices.DirectoryEntry("LDAP://DC=" + $strDomain + ",DC=nads,DC=navy,DC=mil");
				$objDomain = New-Object System.DirectoryServices.DirectoryEntry($strDomain);
			}
		}

		if (($strFilter -eq "") -or ($strFilter -eq $null)){
			#$strFilter = "(&(objectCategory=person))";
			# same results as:
				#$strFilter = "(&(objectCategory=user))";
			#$strFilter = "(&(objectCategory=user)(proxyAddresses=*))";
			#$strFilter = "(&(objectCategory=user)(mail=*" + $Username + "*))";
			$strFilter = "(&(objectCategory=user)(name=*" + $Username + "*))";
			#$strFilter = "(&(name=*" + $Username + "*))";
		}

		if ($bolGC -eq $True){
			$objSearcher = New-Object System.DirectoryServices.DirectorySearcher($strDomain);
		}
		else{
			$objSearcher = New-Object System.DirectoryServices.DirectorySearcher;
			$objSearcher.SearchRoot = $objDomain;
		}
		$objSearcher.PageSize = 1000;
		$objSearcher.Filter = $strFilter;
		$objSearcher.SearchScope = "Subtree";

		#$arrDesiredProps = "name", "proxyAddresses";
		foreach ($i in $arrDesiredProps){$strResults = $objSearcher.PropertiesToLoad.Add($i)};

		$colResults = $objSearcher.FindAll();

		if ($Error){
			$objReturn.Message = "Error" + "`r`n" + $Error;
		}
		else{
			$objReturn.Message = "Success";
			$objReturn.Results = [Int]$colResults.Count;
			if ($colResults.Count -gt 0){
				$objReturn.Returns = $colResults;

				if ($bPSobj -eq $True){
					$objItems = @();
					foreach ($objResult in $objReturn.Returns){
						$objItm = $objResult.Properties;
						#$objItm;
						#Write-Host $objItm.name;
						#Write-Host $objItm.proxyAddresses;
						#Write-Host $objItm.mail;

						$objItem = New-Object PSObject;
						foreach ($prop in $objItm.PropertyNames){
							$objItem | Add-Member -MemberType NoteProperty -Name $prop -Value $objItm.$prop;
						}

						$objItems += $objItem;
					}
					$objReturn.Returns = $objItems;
				}
			}
		}

		return $objReturn;
	}

	function AssignDevPerms{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strCompDN, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strDelegateSID, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strDomainOrDC
		)
		#Description....
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters.
			#$objReturn.Results		= True or False (Were there Errors).
			#$objReturn.Message		= A verbose message of the results (The error message).
			#$objReturn.Returns		= The Computer object that was updated, or $null.
		#$strCompDN = The AD objects DistinguishedName to grant permissions on for $strDelegateSID.  (i.e. CN=WLNRFK390tst,OU=COMPUTERS,OU=NRFK,OU=NAVRESFOR,DC=nadsusea,DC=nads,DC=navy,DC=mil  or  CN=DDALNT000032,OU=COMPUTERS,OU=ALNT,OU=ONR,DC=nadsusea,DC=nads,DC=navy,DC=mil)
		#$strDelegateSID = The SID of the AD object to grant permissions over $strCompDN.
		#$strDomainOrDC = Domain or DC to do the work on.
			#Sample call:
			#$strUserName = "redirect.test";
			#$strUserName = "michael.j.rogers.dev";
			#$objRet = AssignDevPerms ((FindComputer "WLNRFK390TST").DistinguishedName) ((FindUser $strUserName).SID.Value);
			#$objRet = AssignDevPerms ((FindComputer "DLCHLK085463").DistinguishedName) ((FindUser $strUserName).SID.Value);
			#$objRet.Returns
			#$objRet.Returns.ObjectSecurity.Access | Where-Object {$_.IdentityReference -Match $strUserName};
		
		#Helpful URL's:
			#https://social.technet.microsoft.com/Forums/windowsserver/en-US/df3bfd33-c070-4a9c-be98-c4da6e591a0a/forum-faq-using-powershell-to-assign-permissions-on-active-directory-objects?forum=winserverpowershell
			#http://blogs.technet.com/b/joec/archive/2013/04/25/active-directory-delegation-via-powershell.aspx
			#http://blogs.msdn.com/b/adpowershell/archive/2009/10/13/add-object-specific-aces-using-active-directory-powershell.aspx

			#https://social.technet.microsoft.com/Forums/windowsserver/en-US/f7855fb7-99e9-43fe-9852-93e97011df5f/adsicomitchanges-a-constraint-violation-occurred?forum=winserverpowershell
			#https://msdn.microsoft.com/en-us/library/system.security.accesscontrol.objectaccessrule(v=vs.110).aspx

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
			Returns = $null
		}

		if (($strDomainOrDC -eq "") -or ($strDomainOrDC -eq $null)){
			$strDomain = $strCompDN.SubString($strCompDN.IndexOf("DC=") + 3);
			$strDomain = $strDomain.SubString(0, $strDomain.IndexOf(","));

			$InitializeDefaultDrives=$False;
			if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};

			#$strDomainOrDC = "DC=" + [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $strDomain))).Name.Replace(".", ",DC=")
			$strDomainOrDC = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $strDomain))).RidRoleOwner.Name;
		}

		#Chris's idea:
		#Read in the DACL, and then append the following to the DACL, then save/comit changes.
		#(OA;;RPWP;e48d0154-bcf8-11d1-8702-00c04fb96050;;$strDelegateSID)(OA;;WP;4c164200-20c0-11d0-a768-00aa006e0529;;$strDelegateSID)(OA;;SW;f3a64788-5306-11d1-a9c5-0000f80367c1;;$strDelegateSID)(OA;;SW;72e39547-7b18-11d1-adef-00c04fd8d5cd;;$strDelegateSID)(OA;;CR;00299570-246d-11d0-a768-00aa006e0529;;$strDelegateSID)

		if (($strDomain -eq "") -or ($strDomain -eq $null)){
			$objReturn.Message = "Error, need a full DN.  Could not parse a domain out of the supplied DN.";
		}
		else{
			$objTarget = $null;
			$Error.Clear();
			#$objAcl = Get-Acl $objTarget;
			if (($strDomainOrDC -eq "") -or ($strDomainOrDC -eq $null)){
				$objTarget = [ADSI]("LDAP://" + $strCompDN);
			}
			else{
				$objTarget = [ADSI]("LDAP://" + $strDomainOrDC + "/" + $strCompDN);
			}

			if (($objTarget -eq "") -or ($objTarget -eq $null) -or ($Error) -or ($objTarget.Path -eq "") -or ($objTarget.Path -eq $null)){
				$objReturn.Message = "Error, Could not find a Target AD object that matched the DN provided.";
			}
			else{
				#To get the GUIDs from the Network:
				#$rootdse = Get-ADRootDSE;
				#$guidmap = @{};
				#Get-ADObject -SearchBase ($rootdse.SchemaNamingContext) -LDAPFilter "(schemaidguid=*)" -Properties lDAPDisplayName,schemaIDGUID | % {$guidmap[$_.lDAPDisplayName]=[System.GUID]$_.schemaIDGUID}
				#$extendedrightsmap = @{};
				#Get-ADObject -SearchBase ($rootdse.ConfigurationNamingContext) -LDAPFilter "(&(objectclass=controlAccessRight)(rightsguid=*))" -Properties displayName,rightsGuid | % {$extendedrightsmap[$_.displayName]=[System.GUID]$_.rightsGuid};

				#Read current ACLs
				#https://social.technet.microsoft.com/Forums/windowsserver/en-US/df3bfd33-c070-4a9c-be98-c4da6e591a0a/forum-faq-using-powershell-to-assign-permissions-on-active-directory-objects?forum=winserverpowershell
				<#
					#DLCHLK085463 = $strCompDN = "CN=DLCHLK085463,OU=COMPUTERS,OU=CHLK,OU=NAVAIR,DC=nadsuswe,DC=nads,DC=navy,DC=mil";
					GetACLs $strCompDN;
					$objTarget = [ADSI]("LDAP://CN=......");
					$objTarget = [ADSI]("LDAP://" + $strCompDN);
					$objACL = $objTarget.psbase.ObjectSecurity;
					$objACLList = $objACL.GetAccessRules($True, $True, [System.Security.Principal.SecurityIdentifier]);
					foreach($acl in $objACLList){
						$acl;
						Write-Host "`r`n";
					}
				#>

				#Variables common to all
				#$objIdentityReference = New-Object System.Security.Principal.SecurityIdentifier($strDelegateSID);
				$objSID = [System.Security.Principal.SecurityIdentifier] $strDelegateSID;
				$objIdentityReference = [System.Security.Principal.IdentityReference] $objSID;
				$objAccessControlType = [System.Security.AccessControl.AccessControlType] "Allow";
				$objInheritanceType = [System.DirectoryServices.ActiveDirectorySecurityInheritance] "None";
				$objInheritedObjectType = New-Object Guid "bf967aba-0de6-11d0-a285-00aa003049e2";			#the schemaIDGuid for user;
				<#
					#the syntax of "New-ObjectSystem.DirectoryServices.ActiveDirectoryAccessRule":
					#[sid of the object who is either getting or losing permissions],
					#[the permission i'm allowing/denying],
					#[whether i'm allowing or denying],
					#[rightsguid of the property i'm allowing/denying],
					#[type of inheritance],
					#[guid of the class of the object i'm allowing/denying permissions ON]

					#michael.j.rogers.dev  =  $strDelegateSID = "S-1-5-21-283434708-1855628083-519896044-1430629";
					#redirect.test  =  $strDelegateSID = "S-1-5-21-1801674531-2146617017-725345543-4178497";
					#DLCHLK085463 = $strCompDN = "CN=DLCHLK085463,OU=COMPUTERS,OU=CHLK,OU=NAVAIR,DC=nadsuswe,DC=nads,DC=navy,DC=mil";
				#>

				$Error.Clear();
				#VALIDATED_SPN = "{F3A64788-5306-11D1-A9C5-0000F80367C1}"  --  Validated Write Service Principle Name
				<#
					ActiveDirectoryRights : Self
					InheritanceType       : None
					ObjectType            : f3a64788-5306-11d1-a9c5-0000f80367c1
					InheritedObjectType   : 00000000-0000-0000-0000-000000000000
					ObjectFlags           : ObjectAceTypePresent
					AccessControlType     : Allow
					IdentityReference     : S-1-5-21-1801674531-2146617017-725345543-4178497
					IsInherited           : False
					InheritanceFlags      : None
					PropagationFlags      : None
				#>
				$objActiveDirectoryRights = [System.DirectoryServices.ActiveDirectoryRights] "Self";
				$objObjectType = New-Object Guid "F3A64788-5306-11D1-A9C5-0000F80367C1";
				$objAce1 = New-Object System.DirectoryServices.ActiveDirectoryAccessRule $objIdentityReference, $objActiveDirectoryRights, $objAccessControlType, $objObjectType, $objInheritanceType, $objInheritedObjectType;
				#$objAce1 = New-Object System.DirectoryServices.ActiveDirectoryAccessRule $objIdentityReference, "Self", "Allow", "F3A64788-5306-11D1-A9C5-0000F80367C1", "None", "F3A64788-5306-11D1-A9C5-0000F80367C1";

				$Error.Clear();
				#VALIDATED_DNS_HOST_NAME = "{72E39547-7B18-11D1-ADEF-00C04FD8D5CD}"  --  Validated Write dNSHostName
				<#
					ActiveDirectoryRights : Self
					InheritanceType       : None
					ObjectType            : 72e39547-7b18-11d1-adef-00c04fd8d5cd
					InheritedObjectType   : 00000000-0000-0000-0000-000000000000
					ObjectFlags           : ObjectAceTypePresent
					AccessControlType     : Allow
					IdentityReference     : S-1-5-21-1801674531-2146617017-725345543-4178497
					IsInherited           : False
					InheritanceFlags      : None
					PropagationFlags      : None
				#>
				$objActiveDirectoryRights = [System.DirectoryServices.ActiveDirectoryRights] "Self";
				$objObjectType = New-Object Guid "72E39547-7B18-11D1-ADEF-00C04FD8D5CD";
				$objAce2 = New-Object System.DirectoryServices.ActiveDirectoryAccessRule $objIdentityReference, $objActiveDirectoryRights, $objAccessControlType, $objObjectType, $objInheritanceType, $objInheritedObjectType;
				#$objAce2 = New-Object System.DirectoryServices.ActiveDirectoryAccessRule $objIdentityReference, "Self", "Allow", "72E39547-7B18-11D1-ADEF-00C04FD8D5CD", "None", "72E39547-7B18-11D1-ADEF-00C04FD8D5CD";

				$Error.Clear();
				#USER_ACCOUNT_RESTRICTIONS = "{4C164200-20C0-11D0-A768-00AA006E0529}"  --  Write Account Restrictions
				<#
					ActiveDirectoryRights : WriteProperty
					InheritanceType       : None
					ObjectType            : 4c164200-20c0-11d0-a768-00aa006e0529
					InheritedObjectType   : 00000000-0000-0000-0000-000000000000
					ObjectFlags           : ObjectAceTypePresent
					AccessControlType     : Allow
					IdentityReference     : S-1-5-21-1801674531-2146617017-725345543-4178497
					IsInherited           : False
					InheritanceFlags      : None
					PropagationFlags      : None
				#>
				$objActiveDirectoryRights = [System.DirectoryServices.ActiveDirectoryRights] "WriteProperty";
				$objObjectType = New-Object Guid "4C164200-20C0-11D0-A768-00AA006E0529";
				$objAce3 = New-Object System.DirectoryServices.ActiveDirectoryAccessRule $objIdentityReference, $objActiveDirectoryRights, $objAccessControlType, $objObjectType, $objInheritanceType, $objInheritedObjectType;
				#$objAce3 = New-Object System.DirectoryServices.ActiveDirectoryAccessRule $objIdentityReference, "WriteProperty", "Allow", "4C164200-20C0-11D0-A768-00AA006E0529", "None", "4C164200-20C0-11D0-A768-00AA006E0529";

				$Error.Clear();
				#RESET_PASSWORD_GUID = "{00299570-246D-11D0-A768-00AA006E0529}"  --  Reset Password
				<#
					ActiveDirectoryRights : ExtendedRight
					InheritanceType       : None
					ObjectType            : 00299570-246d-11d0-a768-00aa006e0529
					InheritedObjectType   : 00000000-0000-0000-0000-000000000000
					ObjectFlags           : ObjectAceTypePresent
					AccessControlType     : Allow
					IdentityReference     : S-1-5-21-1801674531-2146617017-725345543-4178497
					IsInherited           : False
					InheritanceFlags      : None
					PropagationFlags      : None
				#>
				$objActiveDirectoryRights = [System.DirectoryServices.ActiveDirectoryRights] "ExtendedRight";
				$objObjectType = New-Object Guid "00299570-246D-11D0-A768-00AA006E0529";
				$objAce4 = New-Object System.DirectoryServices.ActiveDirectoryAccessRule $objIdentityReference, $objActiveDirectoryRights, $objAccessControlType, $objObjectType, $objInheritanceType, $objInheritedObjectType;
				#$objAce4 = New-Object System.DirectoryServices.ActiveDirectoryAccessRule $objIdentityReference, "ExtendedRight", "Allow", "00299570-246D-11D0-A768-00AA006E0529", "None", "bf967aba-0de6-11d0-a285-00aa003049e2";

				#Still not sure why the next 2 commands make things work, but Jason had them in his code sample, and they help here.
				$SecOptions = [System.DirectoryServices.DirectoryEntryConfiguration]$objTarget.get_Options();
				#$objTarget.get_Options(); returns -->  Owner, Group, Dacl
				$SecOptions.SecurityMasks = [System.DirectoryServices.SecurityMasks]'Dacl';
				#Doint the above fixes the "Exception calling "CommitChanges" with "0" argument(s): "A constraint violation occurred."" error, 

				$objRet = $objTarget.get_ObjectSecurity().AddAccessRule($objAce1);
				$objRet = $objTarget.get_ObjectSecurity().AddAccessRule($objAce2);
				$objRet = $objTarget.get_ObjectSecurity().AddAccessRule($objAce3);
				$objRet = $objTarget.get_ObjectSecurity().AddAccessRule($objAce4);

				$Error.Clear();
				$objRet = $objTarget.CommitChanges();
				if ($Error){
					$objReturn.Message = $objReturn.Message + "`r`n" + "setInfo: " + $Error;
				}
			}
		}

		$objReturn.Returns = $objTarget;
		if ($objReturn.Message -eq "Error"){
			#No Errors happened/found/reported
			$objReturn.Message = "Success";
			$objReturn.Results = $True;
		}
		#To verify permissions are set (?on local ver only?)
		#$objTarget.ObjectSecurity.Access;
		#$objRet.Returns.ObjectSecurity.Access | Where-Object {$_.IdentityReference -Match "michael.j.rogers.dev"};

		return $objReturn;
	}

	function BuildDisplayName{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$LastName, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$FirstName, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$MI = "", 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$Rank = "", 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$Dep = "", 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$Office = "", 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$Company = "USN", 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$KnownBy = "", 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$Gen = "", 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$FNcc = "US", 
			[ValidateNotNull()][Parameter(Mandatory=$False)]$bNNPI = $False
			#[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bNNPI = $False
		)
		#Description....
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters.
			#$objReturn.Results		= $True or $False.  Were there errors?
			#$objReturn.Message		= A verbose message of the results (The error message).
			#$objReturn.Returns		= The Display Name.
		#$LastName = Last Name.
		#$FirstName = First Name.
		#$MI = Middle Initial.
		#$Rank = Rank.  NOT E/O Grade.
		#$Dep = Department, or GalCMD.
		#$Office = Office, or GalOff.
		#$Company = The company (USN, PACOM, USMC, etc).  Used to determine the exact format of things.
		#$KnownBy = KnownBy Name.  i.e. Tony for Anthony.
		#$Gen = Generation.  i.e. Jr, Sr, etc.
		#$FNcc = Foreign National Country Code.  i.e. FR, GE.
		#$bNNPI = $True or $False.  Is an NNPI Display Name being built?
			#20160325 - Naming standards are being updated. - The NNPI display name for NIPR should be like the display name for SIPR:
			#LastName<COMMA><SPACE>FirstName<SPACE>Initials<SPACE>GenerationQualifier<SPACE>Rank<SPACE>NNPI<HYPHEN>Department<COMMA><SPACE>Office

		#Sample Usage:
			#$strDisplayName = (BuildDisplayName $LastName $FirstName $MI $Rank $Dep $Office $Company $KnownBy $Gen $FNcc).Returns;

		#Display Names - per NMCI Naming Standards (D400 11939.01 section 3.9.4.1)
		#Navy --> Last, First[or KnownBy] MI [Generation] [FORNATL-cc] Rank Department [or GalCMD], Office [or GalOff]
		#NNPI --> Last, First[or KnownBy] MI [Generation] Rank NNPI-Department [or GalCMD], Office [or GalOff]
			#The Standards say to use "http://www.nima.mil/gns/html/fips_10_digraphs.html" for FORNATL-cc values, but it is dead.
			#SRM uses "http://www.state.gov/s/inr/rls/4250.htm".

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
			Returns = ""
		}

		if ($bNNPI -eq $True){
		}

		$Error.Clear();
		$strDisplayName = "";
		#Display Names done per NMCI Naming Standards (D400 11939.01 section 3.9.4.1)
		if (($Company -eq "USN") -or ($Company -eq "usn") -or ($Company -eq "PACOM") -or ($Company -eq "pacom")){
			#USN/Pacom Display Name
				#Last, First[or KnownBy] MI Gen FORNATL-cc Rank GalCMD [or Department], GalOff [or Office]
			#Last, First[or KnownBy]
			if (($KnownBy -ne "") -and ($KnownBy -ne $null)){
				$strDisplayName = $LastName + ", " + $KnownBy + " ";
			}
			else{
				$strDisplayName = $LastName + ", " + $FirstName + " ";
			}
			#Middle
			if (($MI -ne "") -and ($MI -ne $null)){
				if ($MI.Trim().Length -gt 1){
					$MI = $MI.Trim();
					$MI = $MI.SubString(0, 1);
				}
				$strDisplayName = $strDisplayName + $MI + " ";
			}
			#Gen
			if (($Gen -ne "") -and ($Gen -ne $null)){
				$strDisplayName = $strDisplayName + $Gen + " ";
			}
			#Rank
			if (($Rank -ne "") -and ($Rank -ne $null)){
				$strDisplayName = $strDisplayName + $Rank + " ";
			}
			if ($bNNPI -eq $True){
				if ([String]::IsNullOrEmpty($Dep)){
					$strDisplayName = $strDisplayName + "NNPI";
				}
				else{
					$strDisplayName = $strDisplayName + "NNPI-";
				}
			}
			else{
				#CC / FORNATL
				if (($FNcc.Trim() -ne "US") -and ($FNcc -ne "") -and ($FNcc -ne $null)){
					if ($FNcc.Trim().Length -gt 2){
						$FNcc = $FNcc.Trim();
						$FNcc = $FNcc.SubString(0, 2);
					}
					$strDisplayName = $strDisplayName + "FORNATL-" + $FNcc.ToLower() + " ";
				}
			}
			#GALCmd / Department
			if (!([String]::IsNullOrEmpty($Dep))){
				$strDisplayName = $strDisplayName + $Dep;
			}
			#GALOffice / Office
			if (($Office -ne "") -and ($Office -ne $null)){
				$strDisplayName = $strDisplayName.Trim() + ", " + $Office;
			}
		}

		if (($Error) -or ([String]::IsNullOrEmpty($strDisplayName))){
			$objReturn.Results = $False;
			if ($Error){
				$objReturn.Message = "Error: `r`n" + $Error;
			}
			else{
				$objReturn.Message = "Error: Display Name is blank, make sure you provided a valid Company. `r`n";
			}
		}
		else{
			$objReturn.Results = $True;
			$objReturn.Message = "Success";
		}
		$objReturn.Returns = $strDisplayName;

		return $objReturn;
	}

	function CheckNameAvail{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strSamName, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bInteractive = $False, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strMI = "", 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Int]$intMaxLen = 20, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bCheckEmail = $False, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strEDIPI = "", 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bForceInc = $False, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strEmailEnding = "@navy.mil"
		)
		#Description....
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters.
			#$objReturn.Results		= $True or $False.  Found a new "good" name.
			#$objReturn.Message		= A verbose message of the results (The error message).
			#$objReturn.Returns		= The Verified/New Name.
		#$strSamName = sAMAccountName (String) to look for.
		#$bInteractive = $True or $False.  If supplied name is found should we prompt for new name.  If $False this routine will autoincrement the name and return the new name.
		#$strMI = The Middle Initial of the user account being created (only needed if the provided SamName does not have it).
		#$intMaxLen = The Max length the name can be.
		#$bCheckEmail = $True or $False.  Use ADO Search to check ProxyAddresses too.  Can add up to about 2 minutes (per domain) to the search.
		#$strEDIPI = The EDIPI to look for in AD.  If blank (default), the check is skipped.
		#$bForceInc = $True or $False.  Force the name to be incremented before searching AD.  (Mainly for use when CDR has "reserved" the next available name).
		#$strEmailEnding = The email ending to use with the $bCheckEmail option.

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
			Returns = ""
		}

		$strSamName = $strSamName.Trim();
		$strOrigName = $strSamName;
		$strNewName = "";
		$bolNameOK = $False;
		$strWorkLog = "`r`n";

		#For Check for email in use, and EDIPI (if flagged).
		[Array]$arrDesiredProps = @("samAccountName", "name", "proxyAddresses", "mail", "EDIPI", "UserPrincipalName");
		$arrDomains = GetDomains;

		#Get any "custom" endings.
		$strCustEnd = "";
		if ($strSamName.EndsWith(".nnpi")){
			$strCustEnd = ".nnpi";
		}
		if ($strSamName.EndsWith(".dev")){
			$strCustEnd = ".dev";
		}
		if ($strSamName.EndsWith(".cel")){
			$strCustEnd = ".cel";
		}
		if ($strSamName.EndsWith(".fct")){
			$strCustEnd = ".fct";
		}
		if ($strSamName.EndsWith(".adm")){
			$strCustEnd = ".adm";
		}
		$strSamName = $strSamName.SubString(0, ($strSamName.Length - $strCustEnd.Length));

		#Get any ending #'s already provided.
		#Make sure isNumeric() is available.
		if (!(Get-Command "isNumeric" -ErrorAction SilentlyContinue)){
			$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;
			if (([String]::IsNullOrEmpty($ScriptDir)) -or ($Error)){
				$ScriptDir = (Get-Location).ToString();
			}
			if (Test-Path ($ScriptDir + "\Common.ps1")){
				. ($ScriptDir + "\Common.ps1")
			}
			else{
				if (Test-Path ($ScriptDir + "\..\PS-CFW\Common.ps1")){
					. ($ScriptDir + "\..\PS-CFW\Common.ps1")
				}
			}
		}

		$intNameCount = 0;
		for ($intY = 1; $intY -le $strSamName.Length; $intY++){
			#$Error.Clear();
			if (isNumeric ($strSamName.SubString(($strSamName.Length - $intY), $intY))){
				$intNameCount = $strSamName.SubString(($strSamName.Length - $intY), $intY);
			}
			$Error.Clear();		#The if statements errors when ever a number is NOT found.
		}
		[Int]$intNameCount = [Int]$intNameCount;
		if ($bForceInc){
			$intNameCount++;
			if ($intNameCount -eq 1){
				$strOrigName = $strOrigName.Replace($strSamName, $strSamName + [String]$intNameCount);
				$strSamName = $strSamName + [String]$intNameCount
			}
			if ($intNameCount -gt 1){
				$strOrigName = $strOrigName.Replace($strSamName, $strSamName.SubString(0, (($strSamName.Length - $intNameCount.ToString().Length))) + [String]$intNameCount);
				$strSamName = $strSamName.SubString(0, (($strSamName.Length - $intNameCount.ToString().Length))) + [String]$intNameCount;
			}
		}
		if ($intNameCount -gt 0){
			#$strSamName = $strSamName.SubString(0, ($strSamName.Length - $intNameCount));
			#The above does characters equal to the #, not characters equal to the # "width".
				#i.e. a # of 3 would remove 3 characters rather than just the last 1 char.  The next also accomodates 2 dig #'s.
			$strSamName = $strSamName.SubString(0, ($strSamName.Length - $intNameCount.ToString().Length));
		}

		#Break the name down to the last parts.
		$bHadMI = $False;
		$MidName = "";
		$FirstName = $strSamName.SubString(0, ($strSamName.IndexOf(".")));
		$intCount = ($strSamName.IndexOf(".") + 1);
		$LastName = $strSamName.SubString($intCount);
		if ($LastName.IndexOf(".") -gt 0){
			$bHadMI = $True;
			$MidName = $LastName.SubString(0, ($LastName.IndexOf(".")));
			$LastName = $LastName.SubString(($LastName.IndexOf(".") + 1));
		}
		#Compare MiddleNames
		if ([String]::IsNullOrEmpty($strMI)){
			if ([String]::IsNullOrEmpty($MidName)){
				$strMI = "";
				$MidName = "";
			}
			else{
				$strMI = $MidName;
			}
		}
		else{
			$strMI = $strMI.Trim();
			if ([String]::IsNullOrEmpty($MidName)){
				#Use the $strMI provided, make sure only 1 char long.
				if ($strMI.Length -gt 1){
					$strMI = $strMI.SubString(0, 1);
				}
			}
			else{
				$MidName = $MidName.Trim();
				if ($MidName.Length -gt 1){
					$MidName = $MidName.SubString(0, 1);
				}
				
				#Compare $MidName and $strMI.
				if ($strMI -ne $MidName){
					$strMessage = "  The Middle Initial in the SamAccountName ($MidName), and the Middle Initial provided ($strMI) do not match." + "`r`n" + "    We will be using the one out of the SamAccountName (as needed).`r`n";
					$strWorkLog = $strWorkLog + "`r`n" + $strMessage;
					$strMI = $MidName;
				}
			}
		}
		#$strMI

		#Check EDIPI.  If the EDIPI is in use, the name does not matter.
		if (!([String]::IsNullOrEmpty($strEDIPI))){
			$strProgress = "  Verifying EDIPI '*" + $strEDIPI + "*' is NOT in use.`r`n";
			if (([String]::IsNullOrEmpty($txbResults))){
				$strWorkLog = $strWorkLog + "`r`n" + $strProgress;
			}
			else{
				UpdateResults $strProgress $False;
			}
			$strFilter = "(&(objectCategory=user)(|(UserPrincipalName=*" + $strEDIPI + "*)(EDIPI=*" + $strEDIPI + "*)))";
			foreach ($strDomain in $arrDomains){
				$objResults = $null;
				$objResults = ADSearchADO $strOrigName $strDomain $strFilter $arrDesiredProps $True;
				if ($objResults.Results -gt 0){
					#Found EDIPI in use
					#$strMessage = "EDIPI in use: " + ([String]($objResults.Returns)[0].samAccountName).Trim() + " (UPN: " + ([String]($objResults.Returns)[0].UserPrincipalName).Trim() + ") is using EDIPI ";
					#if ([String]::IsNullOrEmpty(([String]($objResults.Returns)[0].EDIPI).Trim())){
					#	$strMessage = $strMessage + "'" + $strEDIPI + "'.";
					#}
					#else{
					#	$strMessage = $strMessage + "'" + ([String]($objResults.Returns)[0].EDIPI).Trim() + "'.";
					#}
					$strMessage = "EDIPI in use: " + ([String]($objResults.Returns)[0].samAccountName).Trim() + " (UPN: " + ([String]($objResults.Returns)[0].UserPrincipalName).Trim() + ") is using EDIPI " + "'" + $strEDIPI + "'.";
					$strWorkLog = $strWorkLog + "`r`n" + $strMessage;
					#break;
					#If the EDIPI is in use, the name does not matter.
					$objReturn.Results = $False;
					$objReturn.Returns = $strOrigName;
					$objReturn.Message = $strWorkLog.Trim();
					return $objReturn;
				}
			}
		}

		$UserADInfo = $null;
		if ($strOrigName.Length -gt $intMaxLen){
			$UserADInfo = "Over $intMaxLen chars.";
			$strWorkLog = $strWorkLog + "`r`n" + "Over $intMaxLen chars.";
		}
		else{
			$UserADInfo = FindUser $strOrigName;

			if ([String]::IsNullOrEmpty($UserADInfo)){
				#$strOrigName not found.
				#Check for email in use
				if ($bCheckEmail -eq $True){
					$strProgress = "  Verifying email '*" + $strOrigName + $strEmailEnding + "*' is NOT in use.`r`n";
					if (([String]::IsNullOrEmpty($txbResults))){
						$strWorkLog = $strWorkLog + "`r`n" + $strProgress;
					}
					else{
						UpdateResults $strProgress $False;
					}
					#if ((!(Get-Command "Get-Recipient" -ErrorAction SilentlyContinue)) -and (!(Get-Command "SetupConn" -ErrorAction SilentlyContinue))){
						#$bolInUse = $null;
						#foreach ($strDomain in $arrDomains){
						#	SetupConn $strDomain "d";
						#	if (Get-Recipient -Identity ($strOrigName + $strEmailEnding) -ErrorAction SilentlyContinue){
						#		#in use.
						#		$UserADInfo = "Email in use.";
						#		$strMessage = ([String]($objResults.Returns)[0].name).Trim() + " is using the email address *" + $strOrigName + "*" + "`r`n" + ([String]($objResults.Returns)[0].proxyAddresses).Trim();
						#		$strWorkLog = $strWorkLog + "`r`n" + $strMessage;
						#		break;
						#	}
						#}
					#}
					#else{
						$strFilter = "(&(objectCategory=user)(proxyAddresses=*" + $strOrigName + $strEmailEnding + "*))";
						foreach ($strDomain in $arrDomains){
							$objResults = $null;
							$objResults = ADSearchADO $strOrigName $strDomain $strFilter $arrDesiredProps $True;
							if ($objResults.Results -gt 0){
								#Found email in use
								$UserADInfo = "Email in use.";
								$strMessage = ([String]($objResults.Returns)[0].name).Trim() + " is using the email address *" + $strOrigName + "*" + "`r`n" + ([String]($objResults.Returns)[0].proxyAddresses).Trim();
								$strWorkLog = $strWorkLog + "`r`n" + $strMessage;
								break;
							}
						}
					#}
				}
			}
		}

		if ([String]::IsNullOrEmpty($UserADInfo)){
			#$strOrigName not found.
			$strNewName = $strOrigName;
			$bolNameOK = $True;
		}
		else{
			#$strOrigName found in use, or too long, or email in use.
			if ($bInteractive){
				#Prompt/confirm the proposed name looks good. 
				#Make sure MsgBox() is available.
				if (!(Get-Command "MsgBox" -ErrorAction SilentlyContinue)){
					$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;
					if (([String]::IsNullOrEmpty($ScriptDir)) -or ($Error)){
						$ScriptDir = (Get-Location).ToString();
					}
					if ((Test-Path ($ScriptDir + "\Forms.ps1"))){
						. ($ScriptDir + "\Forms.ps1")
					}
					else{
						if ((Test-Path ($ScriptDir + "\..\PS-CFW\Forms.ps1"))){
							. ($ScriptDir + "\..\PS-CFW\Forms.ps1")
						}
					}
				}
			}

			do{
				#if ($strNewName -eq ""){
				if ([String]::IsNullOrEmpty($strNewName)){
					$strMessage = "Found an existing AD account with a SamAccountName of '" + $strOrigName + "'.`r`n";
				}
				else{
					$strMessage = "Found an existing AD account with a SamAccountName of '" + $strNewName + "'.`r`n";
				}
				#$strWorkLog = $strWorkLog + "`r`n" + $strMessage;

				#Piece together a new account name suggestion.
				$bolNameOK = $True;
				if ([String]::IsNullOrEmpty($strMI)){
					$strNewName = ($FirstName + "." + $LastName).ToLower();
					$bHadMI = $True;
				}
				else{
					$strNewName = ($FirstName + "." + $strMI + "." + $LastName).ToLower();
				}

				if ($bHadMI){
					$intNameCount++;
				}

				if ($intNameCount -gt 0){
					$strNewName = $strNewName + [String]$intNameCount;
				}

				#Add the "custom" ending.  $strCustEnd
				if (!([String]::IsNullOrEmpty($strCustEnd))){
					#$strNewName = CheckNameEnding $strNewName $strCustEnd;
					$strNewName = $strNewName + $strCustEnd;
				}

				if ($bInteractive){
					$strWorkLog = $strWorkLog + "`r`n" + $strMessage;
					$strWorkLog = $strWorkLog + "Prompted for / Confirmed new name.";
					#Prompt/confirm the proposed name looks good. 
					$strMessage = $strMessage + "Provide a new Login Name (SamAccountName) to use.`r`n`r`n" + "Type 'exit' to abort the process.";
					$strNewName = MsgBox $strMessage "Name already in use." 6 $strNewName;
					$strNewName = $strNewName.Trim();
					if ($strNewName -eq "exit"){
						$bolNameOK = $False;
						$bolDoWork = $False;
						break;
					}
				}
				else{
					#Automatically increment name w/out prompting/confirming.
					#OCM wanted this, and it can be useful for "Automated" processes.
					$strNewName = $strNewName.Trim();
					$strWorkLog = $strWorkLog + "`r`n" + $strMessage;
					$strWorkLog = $strWorkLog + "Assumed the proposed name '$strNewName' is good.";
				}

				if (([String]::IsNullOrEmpty($strNewName)) -or ($strNewName -eq "exit")){
					$bolNameOK = $False;
				}

				#Make sure NewName meets the length requirements.
				if (($strNewName.Length -gt $intMaxLen) -and ($bolNameOK -eq $True)){
					#$strNewName = $strNewName.SubString(0, 20);
					$strMessage = $strMessage + "`r`n" + "The name provided is over $intMaxLen characters long, please shorten it.";
					$strWorkLog = $strWorkLog + "`r`n" + $strMessage;
					$strTempName = $strNewName;
					do{
						if ($bInteractive){
							$strNewName = MsgBox $strMessage "Login Names can only be $intMaxLen Characters" 6 $strTempName;
							$strWorkLog = $strWorkLog + "`r`n" + "User shortened name.";
						}
						else{
							$intShortF = $strNewName.Length - $intMaxLen;
							if ($intShortF -lt $FirstName.Length){
								$strNewName = $strNewName.Replace(($FirstName + "."), (($FirstName.Substring(0, ($FirstName.Length - $intShortF))) + "."));
							}
							else{
								#Need to shorten First and Last.
								$intShortF = $FirstName.Length - 1;
								$strNewName = $strNewName.Replace(($FirstName + "."), (($FirstName.Substring(0, ($FirstName.Length - $intShortF))) + "."));
								$intShortL = $strNewName.Length - $intMaxLen;
								$strNewName = $strNewName.Replace(("." + $LastName), ("." + ($LastName.Substring(0, ($LastName.Length - $intShortL)))));
							}
							$strWorkLog = $strWorkLog + "`r`n" + "Shortened name automatically.";
						}
					} while(($strNewName.Length -gt $intMaxLen));
					$strTempName = "";
				}

				#Check if CDR is OK with the new Name.
				#The CDR check needs to be done outside this routine, so this routine remains generic enough to be used "everywhere".

				#Check AD again
				if ($bolNameOK -eq $True){
					#UpdateResults "  Checking the proposed name in AD, again...`r`n" $False;
					$UserADInfo = $null;
					$UserADInfo = FindUser $strNewName;						#FindUser() searches all domains for the provided username.
					#UpdateResults "`r`n" $False;
					if ((($UserADInfo -ne "") -and ($UserADInfo -ne $null))){
						#Found an existing AD account.
						$bolNameOK = $False;
					}
				}

				#Check for email in use
				if (($bCheckEmail -eq $True) -and ($bolNameOK -eq $True)){
					if ((!(Get-Command "Get-Recipient" -ErrorAction SilentlyContinue)) -and (!(Get-Command "SetupConn" -ErrorAction SilentlyContinue))){
						#$bolInUse = $null;
						foreach ($strDomain in $arrDomains){
							SetupConn $strDomain "d";
							if (Get-Recipient -Identity ($strNewName + "@navy.mil") -ErrorAction SilentlyContinue){
								#in use.
								$UserADInfo = "Email in use.";
								$strMessage = ([String]($objResults.Returns)[0].name).Trim() + " is using the email address *" + $strNewName + "*" + "`r`n" + ([String]($objResults.Returns)[0].proxyAddresses).Trim();
								$strWorkLog = $strWorkLog + "`r`n" + $strMessage;
								break;
							}
						}
					}
					else{
						$strFilter = "(&(objectCategory=user)(proxyAddresses=*" + $strNewName + "*))";
						foreach ($strDomain in $arrDomains){
							$objResults = $null;
							$objResults = ADSearchADO $strNewName $strDomain $strFilter $arrDesiredProps $True;
							if ($objResults.Results -gt 0){
								#Found email in use
								$bolNameOK = $False;
								$strMessage = ([String]($objResults.Returns)[0].name).Trim() + " is using the email address *" + $strNewName + "*" + "`r`n" + ([String]($objResults.Returns)[0].proxyAddresses).Trim();
								$strWorkLog = $strWorkLog + "`r`n" + $strMessage;
								break;
							}
						}
					}
				}

				if ($strNewName -eq "exit"){
					break;
				}

				#If we did not have a MI, we have now used it now (if available).
				$bHadMI = $True;
			} while (($bolNameOK -eq $False) -or ([String]::IsNullOrEmpty($strNewName)))
			$strNewName = $strNewName.ToLower().Trim();
		}

		if ($strWorkLog.Trim() -eq ""){
			$strWorkLog = "Success";
		}
		#return $strNewName;
		$objReturn.Results = $bolNameOK;
		$objReturn.Returns = $strNewName;
		$objReturn.Message = $strWorkLog.Trim();
		return $objReturn;

	}

	function CreateADComputer{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strCompName, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strOU, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strDomain, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strDC = "", 
			[ValidateNotNull()][Parameter(Mandatory=$False)]$objADInfo = $null
		)
		#Description....
		#Note: If the SAMAccountName string provided, does not end with a '$', one will be appended (by powershell) if needed.
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters.
			#$objReturn.Results		= $True or $False.  Did the AD Computer get created?
			#$objReturn.Message		= A verbose message of the results (The error message).
			#$objReturn.Returns		= The Computer object.
		#$strCompName = The SAMAccountName (also Name if not provided) of the computer account to create.
		#$strOU = The LDAP OU path. (i.e. "OU=COMPUTERS,OU=BASE,OU=CMD" or "ou=mfg,dc=noam,dc=corp,dc=contoso,dc=com")
		#$strDomain = The Domain to create the computer account on.  i.e. "sysadmingeek", or "sysadmingeek.com".
		#$strDC = The Domain Controller to create the computer account at.  FQDN or just the server name.
		#$objADInfo = Cutom PowerShell Object that has all the "extra" AD fields to be set.
			#See the following URL for all possiable options:    https://technet.microsoft.com/en-us/library/ee617245.aspx
			#$objADInfo = New-Object PSObject;
			#Add-Member -InputObject $objADInfo -MemberType NoteProperty -Name "Enabled" -Value $True;
			#Add-Member -InputObject $objADInfo -MemberType NoteProperty -Name "OperatingSystem" -Value "OS";
			#Add-Member -InputObject $objADInfo -MemberType NoteProperty -Name "OperatingSystemVersion" -Value "4.7";
			#Add-Member -InputObject $objADInfo -MemberType NoteProperty -Name "Description" -Value "None";

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

		#http://blogs.technet.com/b/heyscriptingguy/archive/2013/12/23/powertip-create-computer-account-in-active-directory-with-powershell.aspx
		#New-ADcomputer -Name $strCompName -SamAccountName $strCompName -Enabled $True -Path $strOU;

		if (($strDomain -eq "") -or ($strDomain -eq $null)){
			#No domain provided.
			$objReturn.Message = "No domain provided.";
		}
		else{
			if (($strOU.IndexOf("DC=") -lt 1) -and ($strOU.IndexOf("dc=") -lt 1)){
				#OU provided does not have the domain ending.
				if (!$strOU.EndsWith(",")){
					$strOU = $strOU + ",";
				}
				$strOU = $strOU + "DC=" + [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $strDomain))).Name.Replace(".", ",DC=")
			}
			#Make sure the OU exists.
			$objResults = Check4OU $strOU $strDomain;
			if ($objResults.Results -gt 0){
			#	#$strDomain = $objResults.Returns[0];
			#	if ($objResults.Results -gt 1){
			#		#"The OU was found on multiple Domains."
			#		$bMatch = $False;
			#		if (($objResults.Returns -Contains $strDomain) -eq $False){
			#			#$strDomain = $strDomain;
			#			$bMatch = $True;
			#		}
			#		else{
			#			for ($intX = 1; $intX -lt $objResults.Results; $intX++){
			#				#$strTemp = $strTemp + ", " + $objResults.Returns[$intX];

			#				if ((($strDomain -eq $objResults.Returns[$intX])) -or (($objResults.Returns[$intX] -Contains $strDomain))){
			#					#One of the domains found matches the requested domain.
			#					$bMatch = $True;
			#					break;
			#				}
			#			}
			#		}

			#		if ($bMatch -eq $False){
			#			$objReturn.Message = "OU found on multiple domains, but none of them match the domain supplied.";
			#			$strDomain = "";
			#		}
			#	}
			#	else{
			#		if (!($strDomain -eq $objResults.Returns[0])){
			#			$objReturn.Message = "OU found, but the domain it was found on does NOT match the supplied domain.";
			#			$strDomain = "";
			#		}
			#	}

				if (!(($objReturn.Message -ne "") -and ($objReturn.Message -ne $null))){
					$InitializeDefaultDrives=$False;
					if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};

					#Make sure we have a DC to do the work on.
					if (($strDC -eq "") -or ($strDC -eq $null)){
						$strDC = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $strDomain))).RidRoleOwner.Name;
					}
					if (($strDC -eq "") -or ($strDC -eq $null)){
						$strDC = $strDomain;
					}

					#Now we can do the work.
					#Create a variable w/ the command, and then run it.
					$strPSCmd = "New-ADcomputer -SamAccountName '" + $strCompName + "' -Path '" + $strOU + "' -Server '" + $strDC + "'";
					foreach ($strProp in $objADInfo.PSObject.Properties){
						#Check if the property exists:
						#if ($objADInfo.PSObject.Properties.Match('Test1').Count) {Write-Host "True"} else {Write-Host "False"};
						if (($strProp.Name -ne "") -and ($strProp.Name -ne $null)){
							if (($strProp.Name -eq "SamAccountName") -or ($strProp.Name -eq "Path") -or ($strProp.Name -eq "Server")){
								#Skip these ones.
									#SamAccountName
									#Path
									#Server
							}
							else{
								if (($strProp.Value -eq $True) -or ($strProp.Value -eq "True") -or ($strProp.Value -eq $False) -or ($strProp.Value -eq "False")){
									if (($strProp.Value -eq $True) -or ($strProp.Value -eq "True")){
										$strPSCmd = $strPSCmd + " -" + $strProp.Name + " $" + $True;
									}
									else{
										$strPSCmd = $strPSCmd + " -" + $strProp.Name + " $" + $False;
									}
								}
								else{
									$strPSCmd = $strPSCmd + " -" + $strProp.Name + " '" + $strProp.Value + "'";
								}
							}
						}
					}

					#Make sure "Name" is Provided.
					if ($strPSCmd.Contains("-Name") -eq $False){
						$strPSCmd = $strPSCmd + " -Name '" + $strCompName + "'";
					}

					#Run the command, that is in the variable.
					$Error.Clear();
					#$strPSCmd = "New-ADcomputer -SamAccountName 'AA-Testing' -Path 'OU=Computers,OU=EDSB,DC=nmci-isf,DC=com' -Server 'NMCISDNIDC02.nmci-isf.com' -Enabled $True -OperatingSystem 'OS' -OperatingSystemVersion '4.7' -Description 'None' -Name 'AA-Testing'";
					Invoke-Expression $strPSCmd;
					if ($Error){
						$objReturn.Message = $Error;
					}
					else{
						$objReturn.Results = $True;
						$objReturn.Message = "Success";

						#Now get the new object and return it.
						$objComp = $(Try {Get-ADComputer -Identity $strCompName -Server $strDC -Properties *} Catch {$null});
						if (($objComp.DistinguishedName -ne "") -and ($objComp.DistinguishedName -ne $null)){
							$objReturn.Returns = $objComp;
						}
					}
				}
			}
			else{
				$objReturn.Message = "OU was not found on any available domains.";
			}
		}

		return $objReturn;
	}

	function CreateADUser{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)]$oADInfo, 
			[ValidateNotNull()][Parameter(Mandatory=$True)]$strOU, 
			[ValidateNotNull()][Parameter(Mandatory=$True)]$strDomain, 
			[ValidateNotNull()][Parameter(Mandatory=$False)]$strDC
		)
		#Description....
		#Taking about 6 seconds in my initial testing.
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters.
			#$objReturn.Results		= $True or $False.  Did the AD User get created?
			#$objReturn.Message		= A verbose message of the results (Or the error message).
			#$objReturn.Returns		= The User object.
		#$oADInfo = Cutom PowerShell Object that has all the AD fields to be set. (Minimum of "CN", "givenName" (First), "sAMAccountName", "SN" (Last), "userPrincipalName")
			#$oADInfo = New-Object PSObject;
			#Add-Member -InputObject $oADInfo -MemberType NoteProperty -Name "CN" -Value "";
			#Add-Member -InputObject $oADInfo -MemberType NoteProperty -Name "sAMAccountName" -Value "";
			#Add-Member -InputObject $oADInfo -MemberType NoteProperty -Name "SN" -Value "";
			#Add-Member -InputObject $oADInfo -MemberType NoteProperty -Name "userPrincipalName" -Value "";
		#$strOU = The LDAP OU path. i.e. "OU=USERS,OU=BASE,OU=CMD".
		#$strDomain = The Domain to create the account on.  i.e. "sysadmingeek", or "sysadmingeek.com" or "DC=nadsusea,DC=nads,DC=navy,DC=mil".
		#$strDC = The Domain Controller to create the account at.  FQDN or just the server name.

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

		#MUST have the following fields, no matter what, to create a User Object.
		if (([String]::IsNullOrEmpty($oADInfo.CN)) -or ([String]::IsNullOrEmpty($oADInfo.givenName)) -or ([String]::IsNullOrEmpty($oADInfo.sAMAccountName)) -or ([String]::IsNullOrEmpty($oADInfo.SN)) -or ([String]::IsNullOrEmpty($oADInfo.userPrincipalName))){
			#CN, sAMAccountName, userPrincipalName, givenName (First), SN (Last)
			$strMessage = "Required AD fields are missing.`r`n(CN, sAMAccountName, userPrincipalName, givenName, SN)";

			#Add-Member -InputObject $oADInfo -MemberType NoteProperty -Name "TheResults" -Value $strMessage -Force;
			#return;

			$objReturn.Message = "Error" + "`r`n" + $strMessage;
		}

		if ($objReturn.Message -eq ""){
			$strMessage = "";
			$strRunningWorkLog = "";

			$InitializeDefaultDrives=$False;
			if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};

			#Get the Domain DistinguishedName from $strDomain.
			if (!($strDomain.StartsWith("DC="))){
				$strDomain = (Get-ADDomain $strDomain).DistinguishedName;
			}

			#$objOU = [ADSI]"LDAP://DomainController/OU=USERS,OU=NRFK,OU=NAVRESFOR,DC=nadsusea,DC=nads,DC=navy,DC=mil"
			#if ([String]::IsNullOrEmpty($strDC)){
			#	#$objOU = [ADSI]"LDAP://OU=USERS,OU=NRFK,OU=NAVRESFOR,DC=nadsusea,DC=nads,DC=navy,DC=mil";
			#}
			#else{
			#	#$objOU = [ADSI]"LDAP://$strDC/OU=USERS,OU=NRFK,OU=NAVRESFOR,DC=nadsusea,DC=nads,DC=navy,DC=mil";
			#}

			$objOU = [ADSI]"LDAP://";
			if (!([String]::IsNullOrEmpty($strDC))){
				$objOU = $objOU + $strDC + "/";
			}
			$objOU = $objOU + $strOU;
			$objOU = $objOU + "," + $strDomain;

			#$cn = $oADInfo.CN;
			#$objUser = $objOU.Create("user", "CN=" + $cn);
			$objUser = $objOU.Create("user", "CN=" + $oADInfo.CN);

			#http://stackoverflow.com/questions/17927525/accessing-values-of-object-properties-in-powershell
			#$oADInfo.psobject.properties | % {$_.Name + " = " + $_.Value};
			foreach ($oProp in $oADInfo.PsObject.Properties){
				if (!(($oProp.Name -eq "cn") -or ($oProp.Name -eq "password") -or ($oProp.Name -eq "AccountDisabled"))){
					$Error.Clear();
					#$oProp.Name + " = " + $oProp.Value;
					$objUser.Put($oProp.Name, $oProp.Value);
					if ($Error){
						if ([String]::IsNullOrEmpty($strMessage)){
							$strMessage = "Error populating the following AD fields: `r`n";
						}
						$strMessage = $strMessage + $oProp.Name + " with '" + $oProp.Value + "'" + "`r`n";
					}
				}
			}
			if (!([String]::IsNullOrEmpty($strMessage))){
				$strMessage = $strMessage + "`r`n";
			}
			#$objUser.Put("sAMAccountName", $sAMAccountName);
			#$objUser.Put("userPrincipalName", $userPrincipalName);
			#$objUser.Put("displayName", $displayName);
			#$objUser.Put("givenName", $FirstName);
			#$objUser.Put("sn", $LastName);
			$Error.Clear();
			$objUser.SetInfo();
			if ($Error){
				$strMessage = $strMessage + "Error saving AD User properties.`r`n" + $Error;
				$strMessage = $strMessage + "`r`n";
			}
			else{
				$strMessage = "Created the following account: `r`n";
				if (!([String]::IsNullOrEmpty($strDC))){
					$strMessage = $strMessage + $strDC + "   ";
				}
				$strMessage = $strMessage + "CN=" + $objUser.CN + "`r`n";
				$objReturn.Results = $True;
				$strSID = $objUser.Sid;
			}
			$strRunningWorkLog = $strRunningWorkLog + $strMessage;
			$strMessage = "";

			$Error.Clear();
			if (!([String]::IsNullOrEmpty($strSID))){
				if (([String]::IsNullOrEmpty($oADInfo.password))){
					#$objUser.SetPassword('S0me.P@$$w0rd4Y0u');

					#Set a random password on the account.
					$strMessage = "Set a random password for the account.`r`n";
					$strRunningWorkLog = $strRunningWorkLog + $strMessage;
					#UpdateResults $strMessage $False;
					$Error.Clear();
					#Generate a random password (19 to 24 chars).
					$strSomePass = [String](Get-Random -Maximum 999 -Minimum 10) + '_NMCI.P@$$w0rd.' + (Get-Random -Maximum 100000 -Minimum 0);
					if ($Error){
						$strSomePass = 'Re@][Y.!S0.R@nd0m.' + (Get-Date).ToString("yyyyMMdd");
					}
				}
				else{
					#$objUser.SetPassword($oADInfo.password);
					$strSomePass = $oADInfo.password;
				}
				$Error.Clear();
				if (([String]::IsNullOrEmpty($strDC))){
					$strResult = Set-ADAccountPassword -Identity $strSID -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $strSomePass -Force);
				}
				else{
					#Set-ADAccountPassword -Identity $strSID -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $strSomePass -Force) -Server $strDC -PassThru | Enable-ADAccount;
					$strResult = Set-ADAccountPassword -Identity $strSID -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $strSomePass -Force) -Server $strDC;
				}
				#$objUser.SetPassword($strSomePass);
				if ($Error){
					$strMessage = "  Error setting a password on the account.`r`n" + $Error;
					$strMessage = $strMessage + "`r`n";
					$strMessage = "" + ("-" * 100) + "`r`n" + $strMessage + "`r`n";
					#UpdateResults $strMessage $False;
					$strRunningWorkLog = $strRunningWorkLog + $strMessage;
				}

				$Error.Clear();
				#$objUser.psbase.InvokeSet("AccountDisabled", $False);
				if (([String]::IsNullOrEmpty($strDC))){
					$strResult = Enable-ADAccount -Identity $strSID;
				}
				else{
					$strResult = Enable-ADAccount -Identity $strSID -Server $strDC;
				}
				if ($Error){
					$strMessage = "Error setting the AccountDisabled property.`r`n" + $Error;
					$strMessage = $strMessage + "`r`n";
					$strRunningWorkLog = $strRunningWorkLog + $strMessage;
				}

				#Do the "User must change password at next logon" checkbox.
				$bolMustChange = $oADInfo.ChangePasswordAtLogon;
				#if ($strWhatToDo.StartsWith("Create-Fct-Mail")){
				#	#$bolMustChange = $False;
				#	#Set LogonHours to Deny ALL.
				#	#Set Mailbox permissions.
				#		#http://support.microsoft.com/kb/304935/
				#	$txbPhoneNotes.Text = $txbPhoneNotes.Text + "`r`nPOCs/Owners: " + $dgvBulk.SelectedRows[0].Cells["OwnerPOC"].Value;

				#	if ($txbCompany.Text -eq "USMC"){
				#		#"User Cannot Change Password"  =  $True
				#			#Should use ACEs to do it, otherwise can fail. http://msdn.microsoft.com/en-us/library/aa746398(VS.85).aspx
				#		#"Password Never Expires"  =  $True
				#	}
				#}
				$strMessage = "Set 'User must change password at next logon' to $bolMustChange.`r`n";
				$strRunningWorkLog = $strRunningWorkLog + $strMessage;
				#UpdateResults $strMessage $False;
				$Error.Clear();
				if (([String]::IsNullOrEmpty($strDC))){
					$strResult = Set-ADUser -Identity $strSID -ChangePasswordAtLogon $bolMustChange;
				}
				else{
					$strResult = Set-ADUser -Identity $strSID -Server $strDC -ChangePasswordAtLogon $bolMustChange;
				}
				if ($Error){
					$strMessage = "  Error setting 'User must change password at next logon'.`r`n" + $Error + "`r`n";
					$strMessage = "" + ("-" * 100) + "`r`n" + $strMessage + "`r`n";
					#UpdateResults $strMessage $False;
					$strRunningWorkLog = $strRunningWorkLog + $strMessage;
				}

				#$Error.Clear();
				#$objUser.SetInfo();
				#if ($Error){
				#	$strMessage = "Error saving AD User object.`r`n" + $Error + "`r`n";
				#	$strRunningWorkLog = $strRunningWorkLog + $strMessage;
				#}
			}
		}

		$objReturn.Returns = $objUser;
		$objReturn.Message = $strRunningWorkLog;
		return $objReturn;
	}

	function Check4OU{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$OUPath, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$RequiredDomain = ""
		)
		#Description....
		#Return.....
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters.
			#$objReturn.Results		= # (typically 0 or 1, but could be more) of domains the OU was found on/in.
			#$objReturn.Message		= A verbose message of the results (The error message).
			#$objReturn.Returns		= Array of Domains the OU was found on/in.
		#$OUPath = The full OU path (LDAP) to check for.  (i.e.  $OUPath="OU=USERS,OU=NRFK,OU=NAVRESFOR,")
		#$RequiredDomain = Any Domains that MUST be checked, and that may not be found by GetDomains().

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
			Message = ""
			Returns = ""
		}

		$InitializeDefaultDrives=$False;
		if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};

		#Get network Domain(s), and trusted domains, from the Network.
		#$arrDomains = GetDomains;
		##Need to get Domains.  GetDomains() requires "AD-Routines.ps1".
		#if (!(Get-Command "GetDomains" -ErrorAction SilentlyContinue)){
		#	$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;
		#	if ((Test-Path ($ScriptDir + "\AD-Routines.ps1"))){
		#		. ($ScriptDir + "\AD-Routines.ps1")
		#	}
		#}

		#if (($RequiredDomain -Contains ".") -or ($RequiredDomain -Match ".")){
		#	$RequiredDomain = $RequiredDomain.SubString(0, $RequiredDomain.IndexOf("."));
		#}
		[System.Collections.ArrayList]$arrDomains = GetDomains $False $False;
		if (($RequiredDomain -ne "") -and ($RequiredDomain -ne $null) -and (!($arrDomains -Contains $RequiredDomain))){
			#if the data has a domain not in the list, add it
			[System.Collections.ArrayList]$arrDomains = GetDomains $False $True;
			$arrDomains += $RequiredDomain;
		}

		#Clean OUPath of any Domain info.
		if ($OUPath.Contains(",DC=")){
			$OUPath = $OUPath.SubString(0, ($OUPath.IndexOf(",DC=") + 1));
		}
		if (!$OUPath.EndsWith(",")){
			$OUPath = $OUPath + ",";
		}

		#Check each domain for the OU.
		#http://techibee.com/active-directory/pstip-quick-way-to-check-if-a-ou-exists-in-active-directory/1975
		#[ADSI]::Exists("LDAP://OU=test,DC=domain,DC=com")
		#[ADSI]::Exists("LDAP://" + $OUPath + "DC=domain,DC=com")
		#[ADSI]::Exists("LDAP://" + $OUPath + (Get-ADDomain $strDomain).DistinguishedName)
		for ($intX = 0; $intX -lt $arrDomains.Count; $intX++){
			#$strDisName = (Get-ADDomain $arrDomains[$intX]).DistinguishedName;
			#Write-Host $arrDomains[$intX];
			$strDisName = "DC=" + [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $arrDomains[$intX]))).Name.Replace(".", ",DC=")
			$arrDomains[$intX] = [String][ADSI]::Exists("LDAP://" + $OUPath + $strDisName) + " = " + $arrDomains[$intX];
		}
		#now check the results, and determine the domain to use.
		$strDomain = "";
		for ($intX = ($arrDomains.Count - 1); $intX -ge 0; $intX--){
			if ($arrDomains[$intX].Contains("=")){
				if ($arrDomains[$intX].Contains("True =")){
					$arrDomains[$intX] = $arrDomains[$intX].SubString($arrDomains[$intX].IndexOf("=") + 1).Trim();
					if ($strDomain -eq ""){
						$strDomain = $arrDomains[$intX].SubString($arrDomains[$intX].IndexOf("=") + 1).Trim();
					}
					else{
						$objReturn.Message = "The OU was found on multiple Domains.";
					}
				}
				else{
					#Remove the entry
					$arrDomains.Remove($arrDomains[$intX]);
				}
			}
			else{
				#Remove the entry
				$arrDomains.Remove($arrDomains[$intX]);
			}
		}
		$objReturn.Results = $arrDomains.Count;
		$objReturn.Returns = $arrDomains

		return $objReturn;
	}

	function DoCLIN16s{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$False)]$objADUser, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$Action, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Int]$NumCLINs, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Long]$lngDefaultMailSize
		)
		#Description....
		#Return.....
		#$objADUser = AD User Object.
		#$Action = The Action to perform.  "Read" (Default), "Set"
		#$NumCLINs = The Number of CLIN16's to assign, if Action = "Set".
		#$lngDefaultMailSize = The Default Mailbox Size (in KB).  Defaults to 250 MB (256000 KB) if not provided (and no interface).

		#USN/USMC 1 CLIN16 = 25MB space  (based on "Prohibit Send")
			#ReadCurrentNumCLINs = (CurrentProhibitSend - lng50MB) / lng25MB
			#SetCurrentNumCLINs = lng50MB + (intNumCLINs * lng25MB)
		#Default Mailbox ("Prohibit Send") is 50MB (USMC).
		#As of 10/30/2012 USN is 100MB.
		#6/1/2014 Now the USN default is 250MB.
		#AD uses KB values
		#The Warning amount is -10MB of the Prohibit Send amount.
		#The Prohibit Send & Receive amount is +250MB of the Prohibit Send amount.

		[Long]$lng1MB = 1024;
		[Long]$lng5MB = 5120;
		[Long]$lng10MB = 10240;
		[Long]$lng25MB = 25600;
		[Long]$lng50MB = 51200;
		[Long]$lng100MB = 102400;
		[Long]$lng250MB = 256000;
		[Long]$lng1GB = 1048576;

		if (($Action -eq "") -or ($Action -eq $null)){
			$Action = "Read";
		}

		if (($lngDefaultMailSize -eq "") -or ($lngDefaultMailSize -eq $null)){
			if (($txbDefaultMailSize.Text -ne "") -and ($txbDefaultMailSize.Text -ne $null)){
				[Long]$lngDefaultMailSize = $txbDefaultMailSize.Text;
				if ($lblSize.Text -eq "MB"){
					[Long]$lngDefaultMailSize = ($lng1MB * $lngDefaultMailSize);
				}
				if ($lblSize.Text -eq "GB"){
					[Long]$lngDefaultMailSize = ($lng1GB * $lngDefaultMailSize);
				}
			}
			else{
				#$txbDefaultMailSize does not exist, or is blank.
				[Long]$lngDefaultMailSize = $lng250MB;
			}
		}

		if ($Action -eq "Set"){
			if (($NumCLINs -ne "") -and ($NumCLINs -ne 0) -and ($NumCLINs -ne $null)){
				$lngProSend = [Long]($lngDefaultMailSize + ($NumCLINs * $lng25MB));
				$lngWarn = [Long]($lngDefaultMailSize + ($NumCLINs * $lng25MB) - $lng10MB);
				$lngProSendRec = [Long]($lngDefaultMailSize + ($NumCLINs * $lng25MB) + $lng250MB);

				#Update the Interface if it exists (not running in the background).
				if (($txbDefaultMailSize.Text -ne "") -and ($txbDefaultMailSize.Text -ne $null)){
					$chkMBDefaults.Checked = $False;
					$txbMBProhibitSend.Text = $lngProSend;
					$txbMBWarning.Text = $lngWarn;
					$txbMBProhibitSendReceive.Text = $lngProSendRec;
				}

				#Update the AD object.
				if (($objADUser -ne "") -and ($objADUser -ne $null)){
					$strRIDMaster = ($objADUser.DistinguishedName).SubString($objADUser.DistinguishedName.IndexOf(",DC=") + 4);
					$strRIDMaster = $strRIDMaster.SubString(0, $strRIDMaster.IndexOf(",DC="));
					#$strRIDMaster = (Get-ADDomain $strRIDMaster).RIDMaster;
					$strRIDMaster = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $strRIDMaster))).RidRoleOwner.Name;

					#$objADUser.mDBUseDefaults = $False;
					UpdateADField $objADUser.DistinguishedName "mDBUseDefaults" $False $False $strRIDMaster;
					#$objADUser.mDBOverQuotaLimit = $lngProSend;
					UpdateADField $objADUser.DistinguishedName "mDBOverQuotaLimit" $lngProSend $False $strRIDMaster;
					#$objADUser.mDBStorageQuota = $lngWarn;
					UpdateADField $objADUser.DistinguishedName "mDBStorageQuota" $lngWarn $False $strRIDMaster;
					#$objADUser.mDBOverHardQuotaLimit = $lngProSendRec;
					UpdateADField $objADUser.DistinguishedName "mDBOverHardQuotaLimit" $lngProSendRec $False $strRIDMaster;
				}
			}
			else{
				if ($NumCLINs -eq 0){
					#Update the Interface if it exists (not running in the background).
					if (($txbDefaultMailSize.Text -ne "") -and ($txbDefaultMailSize.Text -ne $null)){
						$chkMBDefaults.Checked = $True;
						$txbMBProhibitSend.Text = 0;
						$txbMBWarning.Text = 0;
						$txbMBProhibitSendReceive.Text = 0;
					}

					#Update the AD object.
					if (($objADUser -ne "") -and ($objADUser -ne $null)){
						$strRIDMaster = ($objADUser.DistinguishedName).SubString($objADUser.DistinguishedName.IndexOf(",DC=") + 4);
						$strRIDMaster = $strRIDMaster.SubString(0, $strRIDMaster.IndexOf(",DC="));
						#$strRIDMaster = (Get-ADDomain $strRIDMaster).RIDMaster;
						$strRIDMaster = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $strRIDMaster))).RidRoleOwner.Name;

						UpdateADField $objADUser.DistinguishedName "mDBUseDefaults" $True $False $strRIDMaster;
						UpdateADField $objADUser.DistinguishedName "mDBOverQuotaLimit" "" $False $strRIDMaster;
						UpdateADField $objADUser.DistinguishedName "mDBStorageQuota" "" $False $strRIDMaster;
						UpdateADField $objADUser.DistinguishedName "mDBOverHardQuotaLimit" "" $False $strRIDMaster;
					}
				}
			}
		}

		#Read the # of CLINS.  $lng25MB is the size of one CLIN.
		$intCLINs = 0;
		if (($txbDefaultMailSize.Text -ne "") -and ($txbDefaultMailSize.Text -ne $null)){
			if ($chkMBDefaults.Checked -eq $False){
				[Int]$intCLINs = [Int]($txbMBProhibitSend.Text - $lngDefaultMailSize) / $lng25MB;
			}
		}
		else{
			#$txbDefaultMailSize does not exist, or is blank.
			if (($objADUser -ne "") -and ($objADUser -ne $null)){
				if ($objADUser.mDBUseDefaults -eq $False){
					[Int]$intCLINs = [Int]($objADUser.mDBOverQuotaLimit - $lngDefaultMailSize) / $lng25MB;
				}
			}
		}

		return $intCLINs;
	}

	function FindComputer{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$ComputerName, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strDomain
		)
		#Checks All domains (gotten from the Network) for $ComputerName, or just the ones provided.
		#Return.....
		#Paramater Explanation.....

		$InitializeDefaultDrives=$False;
		if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};

		if ((($strDomain -ne "") -and ($strDomain -ne $null)) -or ($ComputerName.Contains("\"))){
			$arrDomains = @($strDomain);
			if ($ComputerName.Contains("\")){
				$arrDomains += $ComputerName.Split("\")[0];
				$ComputerName = $ComputerName.Split("\")[-1];
			}
		}
		else{
			##Need to get Domains.  GetDomains() requires "AD-Routines.ps1".
			#if (!(Get-Command "GetDomains" -ErrorAction SilentlyContinue)){
			#	$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;
			#	if ((Test-Path ($ScriptDir + "\AD-Routines.ps1"))){
			#		. ($ScriptDir + "\AD-Routines.ps1")
			#	}
			#}
			$arrDomains = GetDomains $False $False;
			##----------------------------------------GetDomains----------------------------------------
			#$objForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest();
			#$DomainList = @($objForest.Domains | Select-Object Name);
			#$arrDomains = @($DomainList | foreach {($_.Name).split(".")[0]});
			##Does NOT accomodate FQDN Domain names.
			#[array] $ADDomainTrusts = Get-ADObject -Filter "ObjectClass -eq 'TrustedDomain'" -Properties CN, flatName, Name;
			#foreach ($strDomain in $ADDomainTrusts){
			#	if (($strDomain -ne $null) -and ($strDomain -ne "")){
			#		if ($arrDomains -NotContains $strDomain.flatName){
			#			$arrDomains += $strDomain.flatName;
			#		}
			#	}
			#}
			#foreach ($strDomain in @("pads", "nadsusea", "nadsuswe", "nmci-isf")){
			#	if ($arrDomains -NotContains $strDomain){
			#		$arrDomains += $strDomain;
			#	}
			#}
			##----------------------------------------GetDomains----------------------------------------
		}

		$strDomain = "";
		foreach ($strDomain in $arrDomains){
			if (($strDomain -eq $null) -or ($strDomain -eq "")){
				#break;
			}
			else{
				if ($strDomain -ne "nads"){
					$strProgress = "  Looking in " + $strDomain + " domain for " + $ComputerName + ".`r`n";

					if (($txbResults -ne "") -and ($txbResults -ne $null)){
						UpdateResults $strProgress $False;

						$strRIDMaster = GetOpsMaster2WorkOn $strDomain;
					}
					else{
						#$strProgress;		#Outputs info for when running as background job.

						#$strRIDMaster = (Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
						$strRIDMaster = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $strDomain))).RidRoleOwner.Name;
					}

					$objComp = $(Try {Get-ADComputer -Identity $ComputerName -Server $strRIDMaster -Properties *} Catch {$null});
					if (($objComp.DistinguishedName -ne "") -and ($objComp.DistinguishedName -ne $null)){
						#$strSrcDomain = $strDomain;
						break;
					}
				}
			}
		}

		return $objComp;
	}

	function FindGroup{
 		Param(
 			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$GrpName, 
 			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strDomain, 
 			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strDC
 		)
 		#Checks All domains (gotten from the Network) for $GrpName, or just the ones provided.
 		#Returns the AD Group object.
 		#$GrpName = The group, samaccountname, to look for.
 		#$strDomain = The domain to look for $GrpName in.
 		#$strDC = The DC to use to do the search, over rides $strDomain.  (Short or FQDN.)
 
 		$InitializeDefaultDrives=$False;
 		if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};
 
 		if ([String]::IsNullOrEmpty($strDC)){
 			if (!([String]::IsNullOrEmpty($strDomain))){
 				$arrDomains = @($strDomain);
 			}
 			else{
 				#Get a list of available Domains, from the network.
 				$arrDomains = GetDomains $False $False;
 			}
 
 			if ($GrpName.Contains("\")){
 				#$arrDomains += $Username.Split("\")[0];
 				$arrDomains = @($GrpName.Split("\")[0]) + $arrDomains;
 				$GrpName = $GrpName.Split("\")[-1];
 			}
 
 			$strDomain = "";
 			foreach ($strDomain in $arrDomains){
 				if (!([String]::IsNullOrEmpty($strDomain))){
 					if ($strDomain -ne "nads"){
 						$strProgress = "  Looking in " + $strDomain + " domain for " + $GrpName + ".`r`n";
 
 						if (!([String]::IsNullOrEmpty($txbResults))){
 							UpdateResults $strProgress $False;
 						}
 
 						if (Get-Command "GetOpsMaster2WorkOn" -ErrorAction SilentlyContinue){
 							$strRIDMaster = GetOpsMaster2WorkOn $strDomain;
 						}
 						else{
 							#$strRIDMaster = (Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
 							$ErrorActionPreference = 'SilentlyContinue';
 							$strRIDMaster = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $strDomain))).RidRoleOwner.Name;
 							$ErrorActionPreference = 'Continue';
 						}
 
 						if ([String]::IsNullOrEmpty($strRIDMaster)){
 							$strRIDMaster = $strDomain;
 						}
 
 						#Get-ADGroup finds Exchange Groups too
 						$objGroup = $(Try {Get-ADGroup -Identity $GrpName -Server $strRIDMaster -Properties *} Catch {$null});
 						#if (($objGrp.DistinguishedName -ne "") -and ($objGrp.DistinguishedName -ne $null)){
 						if (!([String]::IsNullOrEmpty($objGroup.DistinguishedName))){
 							#$strSrcDomain = $strDomain;
 							break;
 						}
 					}
 				}
 			}
 		}
 		else{
 			if ($GrpName.Contains("\")){
 				#$arrDomains = @($Username.Split("\")[0]);
 				$GrpName = $GrpName.Split("\")[-1];
 			}
 			#Get-ADGroup finds Exchange Groups too
 			$objGroup = $(Try {Get-ADGroup -Identity $GrpName -Server $strDC -Properties *} Catch {$null});
 		}
 
 		return $objGroup;
 	}
 
	function FindUser{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$Username, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strDomain, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strDC
		)
		#Checks All domains (gotten from the Network) for $Username, or just the ones provided.
		#Return.....
		#$Username = The user, samaccountname, to look for.
		#$strDomain = The domain to look for $Username in.
		#$strDC = The DC to use to do the search, over rides $strDomain.  (Short or FQDN.)

		$InitializeDefaultDrives=$False;
		if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};

		if ([String]::IsNullOrEmpty($strDC)){
			if (!([String]::IsNullOrEmpty($strDomain))){
				$arrDomains = @($strDomain);
			}
			else{
				#Get a list of available Domains, from the network.
				$arrDomains = GetDomains $False $False;
			}

			if ($Username.Contains("\")){
				#$arrDomains += $Username.Split("\")[0];
				$arrDomains = @($Username.Split("\")[0]) + $arrDomains;
				$Username = $Username.Split("\")[-1];
			}

			$strDomain = "";
			foreach ($strDomain in $arrDomains){
				if (!([String]::IsNullOrEmpty($strDomain))){
					if ($strDomain -ne "nads"){
						$strProgress = "  Looking in " + $strDomain + " domain for " + $Username + ".`r`n";

						if (!([String]::IsNullOrEmpty($txbResults))){
							UpdateResults $strProgress $False;
						}

						if (Get-Command "GetOpsMaster2WorkOn" -ErrorAction SilentlyContinue){
							$strRIDMaster = GetOpsMaster2WorkOn $strDomain;
						}
						else{
							#$strRIDMaster = (Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
							$ErrorActionPreference = 'SilentlyContinue';
							$strRIDMaster = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $strDomain))).RidRoleOwner.Name;
							$ErrorActionPreference = 'Continue';
						}

						if ([String]::IsNullOrEmpty($strRIDMaster)){
							$strRIDMaster = $strDomain;
						}

						$objUser = $(Try {Get-ADUser -Identity $UserName -Server $strRIDMaster -Properties *} Catch {$null});
						if (($objUser.DistinguishedName -ne "") -and ($objUser.DistinguishedName -ne $null)){
							#$strSrcDomain = $strDomain;
							break;
						}
					}
				}
			}
		}
		else{
			if ($Username.Contains("\")){
				#$arrDomains = @($Username.Split("\")[0]);
				$Username = $Username.Split("\")[-1];
			}
			$objUser = $(Try {Get-ADUser -Identity $UserName -Server $strDC -Properties *} Catch {$null});
		}

		return $objUser;
	}

	function GetACLs{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strDistName, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bolTranslate = $True
		)
		#Description....
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters.
			#$objReturn.Results		= True or False (Were there Errors).
			#$objReturn.Message		= A verbose message of the results (The error message).
			#$objReturn.Returns		= $null, or an array/list of the ACLs.
		#$strDistName = The AD objects DistinguishedName to get ACL's of.  (i.e. CN=WLNRFK390tst,OU=COMPUTERS,OU=NRFK,OU=NAVRESFOR,DC=nadsusea,DC=nads,DC=navy,DC=mil  or  CN=DDALNT000032,OU=COMPUTERS,OU=ALNT,OU=ONR,DC=nadsusea,DC=nads,DC=navy,DC=mil)
		#$bolTranslate = $True or $False. Translate the GUID's into meaningful names.  (i.e. F3A64788-5306-11D1-A9C5-0000F80367C1 = "Validated Write Service Principle Name")

		#Helpful URL's:
			#https://social.technet.microsoft.com/Forums/windowsserver/en-US/df3bfd33-c070-4a9c-be98-c4da6e591a0a/forum-faq-using-powershell-to-assign-permissions-on-active-directory-objects?forum=winserverpowershell

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
			Message = "Success"
			Returns = $null
		}

		#Sample usage:
		<#
			$objRet = GetACLs "CN=john.alusik,OU=USERS,OU=NRFK,DC=nmci-isf,DC=com";
			foreach ($strEntry in $objRet.Returns){
				#if ($strEntry.NTAccount -eq "NT AUTHORITY\SELF"){
				if ($strEntry.NTAccount -eq "UnKnown"){
					$strEntry;
				}
			}
		#>

		$InitializeDefaultDrives=$False;
		if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};

		$objRootDSE = Get-ADRootDSE;
		$guidmap = @{};				#Hashtable
		Get-ADObject -SearchBase ($objRootDSE.SchemaNamingContext) -LDAPFilter "(schemaidguid=*)" -Properties lDAPDisplayName,schemaIDGUID | % {$guidmap[$_.lDAPDisplayName]=[System.GUID]$_.schemaIDGUID};
		$extendedrightsmap = @{};	#Hashtable
		Get-ADObject -SearchBase ($objRootDSE.ConfigurationNamingContext) -LDAPFilter "(&(objectclass=controlAccessRight)(rightsguid=*))" -Properties displayName,rightsGuid | % {$extendedrightsmap[$_.displayName]=[System.GUID]$_.rightsGuid};

		if (($strDC -eq "") -or ($strDC -eq $null)){
			$strDomain = $strDistName.SubString($strDistName.IndexOf("DC=") + 3);
			$strDomain = $strDomain.SubString(0, $strDomain.IndexOf(","));

			$strDC = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $strDomain))).RidRoleOwner.Name;
		}

		#Read current ACLs
		if (($strDC -eq "") -or ($strDC -eq $null)){
			$objTarget = [ADSI]("LDAP://" + $strDistName);
		}
		else{
			$objTarget = [ADSI]("LDAP://" + $strDC + "/" + $strDistName);
		}
		$objACL = $objTarget.psbase.ObjectSecurity;
		$objACLList = $objACL.GetAccessRules($True, $True, [System.Security.Principal.SecurityIdentifier]);
		$arrReturn=@();
		foreach($acl in $objACLList){
			#$acl;
			#Write-Host "`r`n";

			$Error.Clear();
			$strAccount = $null;
			$objSID = New-Object System.Security.Principal.SecurityIdentifier($acl.IdentityReference);
			$strAccount = $objSID.Translate([System.Security.Principal.NTAccount]);
			if ($Error){
				#The following SIDs are not Translating.
					#S-1-5-32-548 = Account Operators
					#S-1-5-32-560 = BUILTIN\Windows Authorization Access Group
					#S-1-5-32-561 = BUILTIN\Terminal Server License Servers
					#S-1-5-32-554 = BUILTIN\Pre-Windows 2000 Compatible Access
				#https://support.microsoft.com/en-us/kb/243330
				switch ($acl.IdentityReference){
					"S-1-5-32-548"{
						$strAccount = "Account Operators";
						break;
					}
					"S-1-5-32-554"{
						$strAccount = "BUILTIN\Pre-Windows 2000 Compatible Access";
						break;
					}
					"S-1-5-32-560"{
						$strAccount = "BUILTIN\Windows Authorization Access Group";
						break;
					}
					"S-1-5-32-561"{
						$strAccount = "BUILTIN\Terminal Server License Servers";
						break;
					}
					default{
						$objReturn.Message = $objReturn.Message + "`r`n" + $objSID + ": " + $Error;
						#$strAccount = $acl.IdentityReference;
						$strAccount = "UnKnown";
						break;
					}
				}
			}
			$objGUID1 = $acl.ObjectType;
			$objGUID2 = $acl.InheritedObjectType;
			if (($bolTranslate -eq $True) -and ($objGUID1 -ne "00000000-0000-0000-0000-000000000000") -and ($objGUID2 -ne "00000000-0000-0000-0000-000000000000")){
				$bolFound1 = $False;
				$bolFound2 = $False;
				foreach ($strKey in $guidmap.Keys){
					if (($guidmap[$strKey] -eq $objGUID1) -and ($bolFound1 -eq $False)){
						$objGUID1 = $strKey;
						$bolFound1 = $True;
					}
					if (($guidmap[$strKey] -eq $objGUID2) -and ($bolFound2 -eq $False)){
						$objGUID2 = $strKey;
						$bolFound2 = $True;
					}
					if (($bolFound2 -eq $True) -and ($bolFound1 -eq $True)){
						break;
					}
				}
				if (($bolFound2 -eq $False) -or ($bolFound1 -eq $False)){
					foreach ($strKey in $extendedrightsmap.Keys){
						if (($extendedrightsmap[$strKey] -eq $objGUID1) -and ($bolFound1 -eq $False)){
							$objGUID1 = $strKey;
							$bolFound1 = $True;
						}
						if (($extendedrightsmap[$strKey] -eq $objGUID2) -and ($bolFound2 -eq $False)){
							$objGUID2 = $strKey;
							$bolFound2 = $True;
						}
						if (($bolFound2 -eq $True) -and ($bolFound1 -eq $True)){
							break;
						}
					}
				}
			}

			$info = @{
				'ActiveDirectoryRights' = $acl.ActiveDirectoryRights;
				'InheritanceType' = $acl.InheritanceType;
				'ObjectType' = $objGUID1;
				'InheritedObjectType' = $objGUID2;
				'ObjectFlags' = $acl.ObjectFlags;
				'AccessControlType' = $acl.AccessControlType;
				'IdentityReference' = $acl.IdentityReference;
				'NTAccount' = $strAccount;
				'IsInherited' = $acl.IsInherited;
				'InheritanceFlags' = $acl.InheritanceFlags;
				'PropagationFlags' = $acl.PropagationFlags;
			}
			$obj = New-Object -TypeName PSObject -Property $info;
			$arrReturn += $obj;
		}

		$objReturn.Returns = $arrReturn;

		return $objReturn;
	}

	function GetDomains{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bolFQDN = $False, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bolTrusted = $False
		)
		#Description....
		#Return.....
		#$bolFQDN = $True, $False.  Return results in FQDN format.
		#$bolTrusted = $True, $False.  Get Trusted Domains too.

		#Changed to the .NET methods to get AD info, so don't need next 2 lines.
		#$InitializeDefaultDrives=$False;
		#if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};

		#$arrDomains = @("nadsusea", "nadsuswe", "pads", "nmci-isf");

		#Get Domain List, from AD.
		$objForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest();
		$DomainList = @($objForest.Domains | Select-Object Name);
		if ($bolFQDN -eq $False){
			$arrDomains = @($DomainList | foreach {($_.Name).Split(".")[0]});
		}
		else{
			$arrDomains = @($DomainList);
			for ($intX = 0; $intX -lt $arrDomains.Count; $intX++){
				[String]$strTemp = $arrDomains[$intX];
				$strTemp = $strTemp.Replace("@{Name=", "");
				$strTemp = $strTemp.Replace("}", "");
				$arrDomains[$intX] = $strTemp.Trim();
			}
		}

		#Get Trusted Domains:
		if ($bolTrusted -eq $True){
			##http://blog.tyang.org/2011/08/05/powershell-function-get-alldomains-in-a-forest/
			##http://blogs.metcorpconsulting.com/tech/?p=313
			##[array] $ADDomainTrusts = Get-ADObject -Filter "ObjectClass -eq 'TrustedDomain'" -Properties *;
			#[array] $ADDomainTrusts = Get-ADObject -Filter "ObjectClass -eq 'TrustedDomain'" -Properties CN, flatName, Name;
			#foreach ($strDomain in $ADDomainTrusts){
			#	if (($strDomain -ne $null) -and ($strDomain -ne "")){
			#		#$strProgress = "We trust " + $strDomain.Name + " domain.";
			#		#$strProgress;		#Outputs info for when running as background job.
			#		if ($bolFQDN -eq $False){
			#			if ($arrDomains -NotContains $strDomain.flatName){
			#				$arrDomains += $strDomain.flatName;
			#			}
			#		}
			#		else{
			#			if ($arrDomains -NotContains $strDomain.Name){
			#				$arrDomains += $strDomain.Name;
			#			}
			#		}
			#	}
			#}

			#.NET method is more reliable than the PS commandlets.
			#http://blogs.technet.com/b/ashleymcglone/archive/2011/10/12/powershell-sid-walker-texas-ranger-part-3-getting-domain-sids-and-trusts.aspx
			[Array] $ADDomainTrusts = ($objForest.GetAllTrustRelationships())[0].TrustedDomainInformation;
			foreach ($strDomain in $ADDomainTrusts){
				if (($strDomain -ne $null) -and ($strDomain -ne "")){
					#Write-Host "We trust " $strDomain.DnsName
					if ($bolFQDN -eq $False){
						if ($arrDomains -NotContains $strDomain.NetBiosName){
							$arrDomains += $strDomain.NetBiosName;
						}
					}
					else{
						if ($arrDomains -NotContains $strDomain.DnsName){
							$arrDomains += $strDomain.DnsName;
						}
					}
				}
			}
		}

		#List of Domains that are a "must have" in the list.
		foreach ($strDomain in @("pads", "nadsusea", "nadsuswe", "nmci-isf")){
			#if ($arrDomains -NotContains $strDomain){			#does not work w/ partial matches in an array.
			if (!($arrDomains -Match $strDomain)){
				$arrDomains += $strDomain;
			}
		}

		return $arrDomains;
	}

	function GetDCs{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strDomain 
		)
		#Description....
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters.
			#$objReturn.Results		= True or False (Were there Errors).
			#$objReturn.Message		= A verbose message of the results (The error message).
			#$objReturn.Returns		= An array/list of the DCs.
		#$strDomain = The domain to get DC's from.

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
			Message = "Success"
			Returns = $null
		}

		$Error.Clear();
		if (($strDomain -eq "") -or ($strDomain -eq $null)){
			$objDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain();
		}
		else{
			$objDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $strDomain)));
		}
		$arrDCs = @($objDomain.FindAllDomainControllers());
		for ($intX = 0; $intX -lt $arrDCs.Count; $intX++){
			[String]$strTemp = $arrDCs[$intX];
			$strTemp = $strTemp.Replace("@{Name=", "");
			$strTemp = $strTemp.Replace("}", "");
			$arrDCs[$intX] = $strTemp.Trim();
		}

		$objReturn.Returns = $arrDCs;
		if ($Error){
			$objReturn.Results = $False;
			$objReturn.Message = "Error" + "`r`n" + $Error;
		}

		return $objReturn;
	}

	function MoveUser{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$ADUser, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$DestOU
		)
		#Description....
		#Returns a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters.
			#$objReturn.Results		= $True or $False.  Moved or not.
			#$objReturn.Message		= A verbose message of the results. (The error message, or the DN.)
		#$ADUser = The name of the user to move, or an AD Object of the user to move.
		#$DestOU = The LDAP path of the OU to move the user to, with or w/out the domain.  (i.e. OU=USERS,OU=ToBase,OU=ToCmd,DC=domain,DC=com  or  OU=USERS,OU=ToBase,OU=ToCmd)

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
			Message = ""
			Returns = ""
		}

		$InitializeDefaultDrives=$False;
		if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};

		$strDomain = "";
		$objResults = Check4OU $DestOU;
		if ($objResults.Results -gt 0){
			if ($objResults.Results -gt 1){
				#OU found on more than one Domain.
				if ($DestOU.Contains(",DC=")){
					#OU Path has a Domain, will check it and use it if match
					$strDomain = $DestOU.SubString($DestOU.IndexOf(",DC=") + 4);
					$strDomain = $strDomain.SubString(0, $strDomain.IndexOf(",DC="));

					#if ($DestOU.Contains(",DC=$strDomain")){
					if ($objResults.Returns -Contains $strDomain){
						#We can move the User.
						#$strDomain = $strDomain;
					}
					else{
						#Don't know what Domain to use.
						$objReturn.Message = "The OU Path provided was found on more than one Domain," + "`r`n" + "and did not match the Domain specified in the provided OU Path.";
						$objReturn.Results = $False;
						$strDomain = "";
					}
				}
				else{
					#Don't know what Domain to use.
					$objReturn.Message = "The OU Path provided was found on more than one Domain.";
					$objReturn.Results = $False;
				}
			}
			else{
				#We can move the User.
				$strDomain = $objResults.Returns;
			}
		}
		else{
			#OU Path provided could not be found on any of the Domain.
			$objReturn.Message = "The OU Path provided could not be found on any available Domains.";
			$objReturn.Results = $False;
		}

		$objUser = $null;
		if (($strDomain -ne "") -and ($strDomain -ne $null)){
			#We know what domain to move to.
			if (!$DestOU.Contains(",DC=")){
				if (!$DestOU.EndsWith(",")){
					$DestOU = $DestOU + ",";
				}
				$DestOU = $DestOU + (Get-ADDomain $strDomain).DistinguishedName;
			}

			#Check if $ADUser is a String, or already an AD User Object.
			if (($ADUser.GetType().FullName -eq "System.String") -or ($ADUser.GetType().FullName -eq "String")){
				#Get an AD User Object
				#$objUser = $(Try {Get-ADUser -Server $strRIDMaster -Identity $ADUser} Catch {$null});
				$objUser = FindUser $ADUser;
			}
			else{
				#Already have an AD Object
				$objUser = $ADUser;
			}

			if (($objUser -ne "") -and ($objUser -ne $null)){
				#Can do the actual move, once we pull all the parts together
				#Include following Script/File.
				if ((!(Get-Command "GetPathing" -ErrorAction SilentlyContinue))){
					$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;
					if (([String]::IsNullOrEmpty($ScriptDir)) -or ($Error)){
						$ScriptDir = (Get-Location).ToString();
					}
					$arrIncludes = @("Common.ps1");
					foreach ($strInclude in $arrIncludes){
						$Error.Clear();
						if (Test-Path -Path ($ScriptDir + "\..\PS-CFW\" + $strInclude)){
							. ($ScriptDir + "\..\PS-CFW\" + $strInclude)
						}
						else{
							. ($ScriptDir + "\" + $strInclude)
						}
					}
				}
				#$strConfigFile = "\\...\ITSS-Tools\SupportFiles\MiscSettings.txt";
				$strConfigFile = (GetPathing "SupportFiles").Returns.Rows[0]['Path'];
				if ([String]::IsNullOrEmpty($strConfigFile)){
					$strConfigFile = "\\nawesdnifs101v.nadsuswe.nads.navy.mil\NMCIISF02$\ITSS-Tools\SupportFiles\MiscSettings.txt";
				}
				else{
					$strConfigFile = $strConfigFile + "MiscSettings.txt";
				}

				$strDestDomain = $strDomain;
				$strTempError = $null;
				$Error.Clear();
				#$strDestRIDMaster = (Get-ADDomain $strDestDomain -ErrorAction SilentlyContinue).RIDMaster;
				$strDestRIDMaster = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $strDestDomain))).RidRoleOwner.Name;
				if (($strDestRIDMaster -eq "") -or ($strDestRIDMaster -eq $null)){
					$strTempError = $Error;
					$strDestRIDMaster = $strDomain;
					#Can NOT use just the Domain Name.
					##So read the RidMaster info from my "MiscSettings" file.
					#if ((Test-Path -Path $strConfigFile)){
					#	foreach ($strLine in [System.IO.File]::ReadAllLines($strConfigFile)) {
					#		if ($strLine.StartsWith($strDomain)){
					#			$strDestRIDMaster = $strLine.SubString($strLine.IndexOf("=") + 1).Trim();
					#			break;
					#		}
					#	}
					#}
				}

				#$objUser.CanonicalName   =   "nadsusea.nads.navy.mil/NAVRESFOR/NRFK/USERS/redirect.test"
				#$objUser.DistinguishedName = "CN=redirect.test,OU=USERS,OU=NRFK,OU=NAVRESFOR,DC=nadsusea,DC=nads,DC=navy,DC=mil"
				$strSrcOU = $objUser.DistinguishedName;
				$strSrcDomain = $strSrcOU.SubString($strSrcOU.IndexOf(",DC=") + 4);
				$strSrcDomain = $strSrcDomain.SubString(0, $strSrcDomain.IndexOf(",DC="));
				#$strSrcRIDMaster = (Get-ADDomain $strSrcDomain -ErrorAction SilentlyContinue).RIDMaster;
				$strSrcRIDMaster = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $strSrcDomain))).RidRoleOwner.Name;
				if (($strSrcRIDMaster -eq "") -or ($strSrcRIDMaster -eq $null)){
					$strSrcRIDMaster = $strSrcDomain;
					#Can NOT use just the Domain Name.
					##So read the RidMaster info from my "MiscSettings" file.
					#if ((Test-Path -Path $strConfigFile)){
					#	foreach ($strLine in [System.IO.File]::ReadAllLines($strConfigFile)) {
					#		if ($strLine.StartsWith($strDomain)){
					#			$strSrcRIDMaster = $strLine.SubString($strLine.IndexOf("=") + 1).Trim();
					#			break;
					#		}
					#	}
					#}
				}

				if ($objUser.Name.StartsWith("^")){
					#Remove ^ (Rename account).
					#Next line changes/updates the "DistinguishedName", "CanonicalName", "CN", and "Name" fields
					Rename-ADObject -Server $strSrcRIDMaster $objUser.DistinguishedName -NewName $objUser.SamAccountName;
					#$objUser = FindUser $objUser.SamAccountName;
					#$objUser = Get-ADUser -Server $strSrcRIDMaster -Identity $objUser.SamAccountName;
					$objUser = Get-ADUser -Server $strSrcRIDMaster -Identity $objUser.SamAccountName -Properties *;
					$objReturn.Message = "Renamed the AD object from '" + $objUser.Name + "' to '" + $objUser.SamAccountName + "'.`r`n";
				}

				if ((($strDestRIDMaster -ne "") -and ($strDestRIDMaster -ne $null)) -and ($strDestRIDMaster -ne $strSrcRIDMaster)){
					#Check if Src OU and Dest OU are the same
					If ($objUser.DistinguishedName.Contains($DestOU)){
						$objReturn.Message = $objReturn.Message + "The AD object is already at the OU Path provided.`r`n";
						$objReturn.Results = $True;
					}
					else{
						$Error.Clear();
						if ($strDestRIDMaster -ne $strSrcRIDMaster){
							$(Try {Get-ADUser -Server $strSrcRIDMaster -Identity $objUser.DistinguishedName | Move-ADObject -Server $strSrcRIDMaster -TargetPath $DestOU -TargetServer $strDestRIDMaster} Catch {$null});
						}
						else{
							$(Try {Get-ADUser -Server $strSrcRIDMaster -Identity $objUser.DistinguishedName | Move-ADObject -Server $strSrcRIDMaster -TargetPath $DestOU} Catch {$null});
						}
						if ($Error){
							$objReturn.Message = $objReturn.Message + "Failed to move the AD object '" + $objUser.SamAccountName + "' to the Destination OU `r`n'$DestOU'.`r`n";
							$objReturn.Message = $objReturn.Message + $Error + "`r`n";
							$objReturn.Results = $False;
						}
						else{
							$objReturn.Results = $True;
							$strMessage = "Successfully moved the AD object '" + $objUser.SamAccountName + "' to the Destination OU.`r`n"
							if ($strDestRIDMaster -ne $strSrcRIDMaster){
								$objUser = Get-ADUser -Server $strDestRIDMaster -Identity $objUser.SamAccountName -Properties *;
								$strMessage = $strMessage + $strDestRIDMaster
							}
							else{
								$objUser = Get-ADUser -Server $strSrcRIDMaster -Identity $objUser.SamAccountName -Properties *;
								$strMessage = $strMessage + $strSrcRIDMaster
							}

							#$objReturn.Message = $objReturn.Message + $objUser.DistinguishedName + "`r`n";
							$strMessage = $strMessage + "   " + $objUser.DistinguishedName + "`r`n";
							$objReturn.Message = $objReturn.Message + $strMessage;
						}
					}
				}
				else{
					#Could not get the RID Master for the Destination Domain.
					$objReturn.Results = $False;
					$objReturn.Message = $objReturn.Message + "Failed to move the AD object '" + $objUser.SamAccountName + "' to the Destination OU `r`n'$DestOU'.`r`nCould not determine the RID/Ops Master DC from the network.`r`n";
					if ($strTempError){
						$objReturn.Message = $objReturn.Message + $strTempError + "`r`n";
					}
				}
			}
			else{
				$objReturn.Results = $False;
				$objReturn.Message = "Could not find an AD object with a name of '" + $ADUser + "' in any Domain on the network.`r`n";
			}
		}

		return $objReturn;
	}

	function TSRead{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strUserDN, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$Property
		)
		#Description....
		#Return.....
		#$strUserDN = The DistinguishedName of the account. (i.e. CN=redirect.test,OU=USERS,OU=NRFK,OU=NAVRESFOR,DC=nadsusea,DC=nads,DC=navy,DC=mil)
		#$Property = What TS Property to get. (blank or $null returns all) "allowLogon", "TerminalServicesHomeDirectory", "TerminalServicesHomeDrive", "TerminalServicesProfilePath"

		<#
		#Setup a PSObject to return.
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
		#>

		$InitializeDefaultDrives=$False;
		if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};

		[System.Collections.ArrayList]$arrValues = @();
		#$arrValues = @();

		$Error.Clear();
		#Prep to interact with Term Serv Attributes.
		$arrTSProperties = "allowLogon","TerminalServicesHomeDirectory","TerminalServicesHomeDrive","TerminalServicesProfilePath";
		if ($strUserDN.IndexOf(",") -gt 0){
			$strLDAP = "LDAP://$($strUserDN.SubString($strUserDN.IndexOf(",") + 1))";						#"LDAP://OU=USERS,OU=SDNI,OU=COMPACFLT,DC=nadsuswe,DC=nads,DC=navy,DC=mil"
			$strUserDN = $strUserDN.SubString(0, $strUserDN.IndexOf(","));									#"CN=redirect.test"
			$objOU = [ADSI]$strLDAP;
			$objADSIUser = $objOU.PSBase.get_children().find($strUserDN);

			#READ Term Serv Attributes
			if (!($Error)){
				foreach($strProperty in $arrTSProperties){
					if (($Property -ne "") -and ($Property -ne $null)){
						if ($Property -eq $($strProperty)){
							$arrValues += $($objADSIUser.PSBase.InvokeGet($strProperty)).ToString();
						}
					}
					else{
						#$strMessage = "$($strProperty) value: $($objADSIUser.PSBase.InvokeGet($strProperty))";
						#MsgBox $strMessage;
						$arrValues += "$($strProperty) = $($objADSIUser.PSBase.InvokeGet($strProperty))";
					}
				}

				#$objReturn.Results = $True;
				#$objReturn.Message = "Success";
			}
			else{
				#$objReturn.Message = "Error `r`n" + $Error;
			}
		}

		return $arrValues;
		#$objReturn.Returns = $arrValues;
		#return $objReturn;
	}

	function TSSet{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strUserDN, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$Attribute, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$Value
		)
		#Description....
		#Return.....
		#$strUserDN = The DistinguishedName of the account. (i.e. CN=redirect.test,OU=USERS,OU=NRFK,OU=NAVRESFOR,DC=nadsusea,DC=nads,DC=navy,DC=mil)
		#$Attribute = What TS Property to update.  "allowLogon", "TerminalServicesHomeDirectory", "TerminalServicesHomeDrive", "TerminalServicesProfilePath"
		#$Value = The Value to populate $Attribute with.

		$bolSuccess = $False;

		$InitializeDefaultDrives=$False;
		if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};

		#Prep to interact with Term Serv Attributes.
		#$arrTSProperties = "allowLogon","TerminalServicesHomeDirectory","TerminalServicesHomeDrive","TerminalServicesProfilePath";
		if ($strUserDN.IndexOf(",") -gt 0){
			$Error.Clear();
			$strLDAP = "LDAP://$($strUserDN.SubString($strUserDN.IndexOf(",") + 1))";						#"LDAP://OU=USERS,OU=SDNI,OU=COMPACFLT,DC=nadsuswe,DC=nads,DC=navy,DC=mil"
			$strUserDN = $strUserDN.SubString(0, $strUserDN.IndexOf(","));									#"CN=redirect.test"
			$objOU = [ADSI]$strLDAP;
			$objADSIUser = $objOU.PSBase.get_children().find($strUserDN);

			#SET Term Serv Attributes
			if (!($Error)){
				#$objADSIUser.PSBase.invokeSet("allowLogon", $Value);
				#$objADSIUser.PSBase.invokeSet("TerminalServicesHomeDirectory", $Value);
				#$objADSIUser.PSBase.invokeSet("TerminalServicesProfilePath", $Value);
				#$objADSIUser.PSBase.invokeSet("TerminalServicesHomeDrive", $Value);
				$objADSIUser.PSBase.invokeSet($Attribute, $Value);
				$objADSIUser.setinfo();

				if (!($Error)){
					$bolSuccess = $True;
				}
			}
		}

		return $bolSuccess;
	}

	function UpdateADField{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$ADUserDN, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$FieldName, 
			[ValidateNotNull()][Parameter(Mandatory=$True)]$NewValue, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$MultiVal = $False, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$RIDMaster
		)
		#Description....
		#No Return.
		#$ADUserDN = AD User DistinguishedName.  (i.e. CN=redirect.test,OU=USERS,OU=SDNI,OU=COMPACFLT,DC=nadsuswe,DC=nads,DC=navy,DC=mil)
		#$FieldName = AD field name to update.
		#$NewValue = The new value to put in FieldName.
		#$MultiVal = If FieldName is a MultiValue field, then should pass $True, $False, "Add", "Remove".
		#$RIDMaster = The Ops Master / RID Master, or domain, to do the work on.

		$InitializeDefaultDrives=$False;
		if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};

		if (($RIDMaster -eq "") -or ($RIDMaster -eq $null)){
			$RIDMaster = $ADUserDN.SubString($ADUserDN.IndexOf(",DC=") + 4);
			$RIDMaster = $RIDMaster.SubString(0, $RIDMaster.IndexOf(",DC="));

			if (($frmAScIIGUI -ne "") -and ($frmAScIIGUI -ne $null)){
				$RIDMaster = GetOpsMaster2WorkOn $RIDMaster;
			}
			else{
				#$RIDMaster = (Get-ADDomain $RIDMaster).RIDMaster;
				$RIDMaster = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $RIDMaster))).RidRoleOwner.Name;
			}
		}

		#http://windowsitpro.com/active-directory/meet-set-aduser-all-purpose-hammer
		#Set-ADUser -Identity $UserADInfo.DistinguishedName -Server $strRIDMaster -Add @{wWWHomePage = "value"}
		#Set-ADUser -Identity $UserADInfo.DistinguishedName -Server $strRIDMaster -Replace @{wWWHomePage = "value2"}
		#Set-ADUser -Identity $UserADInfo.DistinguishedName -Server $strRIDMaster -Remove @{wWWHomePage = "value2"}
		#Set-ADUser -Identity $UserADInfo.DistinguishedName -Server $strRIDMaster -Clear wWWHomePage;

		if (($MultiVal -ne $False) -and ($MultiVal -ne $null) -and ($MultiVal -ne "")){
			if ((($MultiVal -eq "yes") -or ($MultiVal -eq $True) -and ($NewValue -ne "")) -or ($MultiVal -eq "Add")){
				Set-ADUser -Identity $ADUserDN -Server $RIDMaster -Add @{$FieldName = $NewValue};
			}
			if ((($MultiVal -eq "yes") -or ($MultiVal -eq $True) -and ($NewValue -ne "")) -or ($MultiVal -eq "Remove")){
				Set-ADUser -Identity $ADUserDN -Server $RIDMaster -Remove @{$FieldName = $NewValue};
			}
		}
		else{
			if (($NewValue -eq "") -or ($NewValue -eq $null)){
				Set-ADUser -Identity $ADUserDN -Server $RIDMaster -Clear $FieldName;
			}
			else{
				$strCheckVal = [String](Get-ADUser -Identity $ADUserDN -Server $RIDMaster -Properties * | Select $FieldName);
				if ($strCheckVal.Contains("Microsoft.ActiveDirectory.Management.ADPropertyValueCollection")){
					#The field name supplied probably does not exist.
					$strCheckVal = "";
				}
				else{
					$strCheckVal = $strCheckVal.Replace("@{", "");
					$strCheckVal = $strCheckVal.Replace("}", "");
					$strCheckVal = $strCheckVal.Replace($FieldName + "=", "").Trim();
				}

				if (($strCheckVal.Trim() -eq "") -or ($strCheckVal -eq $null)){
					Set-ADUser -Identity $ADUserDN -Server $RIDMaster -Add @{$FieldName = $NewValue};
				}
				else{
					Set-ADUser -Identity $ADUserDN -Server $RIDMaster -Replace @{$FieldName = $NewValue};
				}
			}
		}
	}


#All the code below here is the C Sharp User Class Chris did (.NET 4+ required).
if ($PSVersionTable.CLRVersion.Major -ge 4){
	$cs = @"
	using System;
	using System.Collections.Generic;
	using System.DirectoryServices;
	using System.DirectoryServices.AccountManagement;
	using System.DirectoryServices.ActiveDirectory;
	using System.Text;
	using System.Threading.Tasks;

	namespace NMCI.AD
	{
		[DirectoryRdnPrefix("CN")]
		[DirectoryObjectClass("User")]
		public class NMCIUserPrincipal : UserPrincipal
		{
			// Inplement the constructor using the base class constructor. 
			public NMCIUserPrincipal(PrincipalContext context) : base(context)
			{
			}
			// Implement the constructor with initialization parameters.    
			public NMCIUserPrincipal(PrincipalContext context, 
								 string samAccountName, 
								 string password, 
								 bool enabled)
								 : base(context, 
										samAccountName, 
										password, 
										enabled)
			{
			}
		   // Create the other home phone property.  
			[DirectoryProperty("otherHomePhone")]
			public string[] HomePhoneOther
			{
				get
				{
					int len = ExtensionGet("otherHomePhone").Length;
					if (len == 0) return null;
				   string[] otherHomePhone = new string[len];
					object[] otherHomePhoneRaw = ExtensionGet("otherHomePhone");
				   for (int i = 0; i < len; i++)
					{
						otherHomePhone[i] = (string)otherHomePhoneRaw[i];
					}
					return otherHomePhone;
				}
				set
				{
					ExtensionSet("otherHomePhone", value);
				}
			}
			// Create the logoncount property.    
			[DirectoryProperty("LogonCount")]
			public Nullable<int> LogonCount
			{
				get
				{
					if (ExtensionGet("LogonCount").Length != 1)
						return null;
					return ((Nullable<int>)ExtensionGet("LogonCount")[0]);
				}
			}
			// Create the assistant property.
			[DirectoryProperty("assistant")]
			public string Assistant
			{
				get
				{
					if (ExtensionGet("assistant").Length != 1)
						return null;
					return (string)ExtensionGet("assistant")[0];
				}
				set
				{
					ExtensionSet("assistant", value);
				}
			}
			// Create the base property.
			[DirectoryProperty("base")]
			public string Base
			{
				get
				{
					if (ExtensionGet("base").Length != 1)
						return null;
					return (string)ExtensionGet("base")[0];
				}
				set
				{
					ExtensionSet("base", value);
				}
			}
			// Create the building property.
			[DirectoryProperty("building")]
			public string Building
			{
				get
				{
					if (ExtensionGet("building").Length != 1)
						return null;
					return (string)ExtensionGet("building")[0];
				}
				set
				{
					ExtensionSet("building", value);
				}
			}
			// Create the citizenship property.
			[DirectoryProperty("citizenship")]
			public string Citizenship
			{
				get
				{
					if (ExtensionGet("citizenship").Length != 1)
						return null;
					return (string)ExtensionGet("citizenship")[0];
				}
				set
				{
					ExtensionSet("citizenship", value);
				}
			}
			// Create the CN property.		//hjs
			[DirectoryProperty("CN")]
			public string CN
			{
				get
				{
					if (ExtensionGet("CN").Length != 1)
						return null;
					return (string)ExtensionGet("CN")[0];
				}
				set
				{
					ExtensionSet("CN", value);
				}
			}
			// Create the co property.
			[DirectoryProperty("co")]
			public string Co
			{
				get
				{
					if (ExtensionGet("co").Length != 1)
						return null;
					return (string)ExtensionGet("co")[0];
				}
				set
				{
					ExtensionSet("co", value);
				}
			}
			// Create the company property.
			[DirectoryProperty("company")]
			public string Company
			{
				get
				{
					if (ExtensionGet("company").Length != 1)
						return null;
					return (string)ExtensionGet("company")[0];
				}
				set
				{
					ExtensionSet("company", value);
				}
			}
			// Create the department property.
			[DirectoryProperty("department")]
			public string Department
			{
				get
				{
					if (ExtensionGet("department").Length != 1)
						return null;
					return (string)ExtensionGet("department")[0];
				}
				set
				{
					ExtensionSet("department", value);
				}
			}
			//// Create the displayName property.
			//[DirectoryProperty("displayName")]
			//public string DisplayName
			//{
			//    get
			//    {
			//        if (ExtensionGet("displayName").Length != 1)
			//            return null;
			//        return (string)ExtensionGet("displayName")[0];
			//    }
			//    set
			//    {
			//        ExtensionSet("displayName", value);
			//    }
			//}
			// Create the division property.
			[DirectoryProperty("division")]
			public string Division
			{
				get
				{
					if (ExtensionGet("division").Length != 1)
						return null;
					return (string)ExtensionGet("division")[0];
				}
				set
				{
					ExtensionSet("division", value);
				}
			}
			// Create the doduid property.
			[DirectoryProperty("doduid")]
			public string Doduid
			{
				get
				{
					if (ExtensionGet("doduid").Length != 1)
						return null;
					return (string)ExtensionGet("doduid")[0];
				}
				set
				{
					ExtensionSet("doduid", value);
				}
			}
			// Create the eDIPI property.
			[DirectoryProperty("eDIPI")]
			public string eDIPI
			{
				get
				{
					if (ExtensionGet("eDIPI").Length != 1)
						return null;
					return (string)ExtensionGet("eDIPI")[0];
				}
				set
				{
					ExtensionSet("eDIPI", value);
				}
			}
			// Create the extensionAttribute1 property.
			[DirectoryProperty("extensionAttribute1")]
			public string ExtensionAttribute1
			{
				get
				{
					if (ExtensionGet("extensionAttribute1").Length != 1)
						return null;
					return (string)ExtensionGet("extensionAttribute1")[0];
				}
				set
				{
					ExtensionSet("extensionAttribute1", value);
				}
			}
			// Create the extensionAttribute2 property.
			[DirectoryProperty("extensionAttribute2")]
			public string ExtensionAttribute2
			{
				get
				{
					if (ExtensionGet("extensionAttribute2").Length != 1)
						return null;
					return (string)ExtensionGet("extensionAttribute2")[0];
				}
				set
				{
					ExtensionSet("extensionAttribute2", value);
				}
			}
			// Create the extensionAttribute3 property.
			[DirectoryProperty("extensionAttribute3")]
			public string ExtensionAttribute3
			{
				get
				{
					if (ExtensionGet("extensionAttribute3").Length != 1)
						return null;
					return (string)ExtensionGet("extensionAttribute3")[0];
				}
				set
				{
					ExtensionSet("extensionAttribute3", value);
				}
			}
			// Create the extensionAttribute4 property.
			[DirectoryProperty("extensionAttribute4")]
			public string ExtensionAttribute4
			{
				get
				{
					if (ExtensionGet("extensionAttribute4").Length != 1)
						return null;
					return (string)ExtensionGet("extensionAttribute4")[0];
				}
				set
				{
					ExtensionSet("extensionAttribute4", value);
				}
			}
			// Create the extensionAttribute5 property.
			[DirectoryProperty("extensionAttribute5")]
			public string ExtensionAttribute5
			{
				get
				{
					if (ExtensionGet("extensionAttribute5").Length != 1)
						return null;
					return (string)ExtensionGet("extensionAttribute5")[0];
				}
				set
				{
					ExtensionSet("extensionAttribute5", value);
				}
			}
			// Create the extensionAttribute6 property.
			[DirectoryProperty("extensionAttribute6")]
			public string ExtensionAttribute6
			{
				get
				{
					if (ExtensionGet("extensionAttribute6").Length != 1)
						return null;
					return (string)ExtensionGet("extensionAttribute6")[0];
				}
				set
				{
					ExtensionSet("extensionAttribute6", value);
				}
			}
			// Create the extensionAttribute7 property.
			[DirectoryProperty("extensionAttribute7")]
			public string ExtensionAttribute7
			{
				get
				{
					if (ExtensionGet("extensionAttribute7").Length != 1)
						return null;
					return (string)ExtensionGet("extensionAttribute7")[0];
				}
				set
				{
					ExtensionSet("extensionAttribute7", value);
				}
			}
			// Create the extensionAttribute8 property.
			[DirectoryProperty("extensionAttribute8")]
			public string ExtensionAttribute8
			{
				get
				{
					if (ExtensionGet("extensionAttribute8").Length != 1)
						return null;
					return (string)ExtensionGet("extensionAttribute8")[0];
				}
				set
				{
					ExtensionSet("extensionAttribute8", value);
				}
			}
			// Create the extensionAttribute4 property.
			[DirectoryProperty("extensionAttribute9")]
			public string ExtensionAttribute9
			{
				get
				{
					if (ExtensionGet("extensionAttribute9").Length != 1)
						return null;
					return (string)ExtensionGet("extensionAttribute9")[0];
				}
				set
				{
					ExtensionSet("extensionAttribute9", value);
				}
			}
			// Create the extensionAttribute10 property.
			[DirectoryProperty("extensionAttribute10")]
			public string ExtensionAttribute10
			{
				get
				{
					if (ExtensionGet("extensionAttribute10").Length != 1)
						return null;
					return (string)ExtensionGet("extensionAttribute10")[0];
				}
				set
				{
					ExtensionSet("extensionAttribute10", value);
				}
			}
			// Create the extensionAttribute11 property.
			[DirectoryProperty("extensionAttribute11")]
			public string ExtensionAttribute11
			{
				get
				{
					if (ExtensionGet("extensionAttribute11").Length != 1)
						return null;
					return (string)ExtensionGet("extensionAttribute11")[0];
				}
				set
				{
					ExtensionSet("extensionAttribute11", value);
				}
			}
			// Create the extensionAttribute12 property.
			[DirectoryProperty("extensionAttribute12")]
			public string ExtensionAttribute12
			{
				get
				{
					if (ExtensionGet("extensionAttribute12").Length != 1)
						return null;
					return (string)ExtensionGet("extensionAttribute12")[0];
				}
				set
				{
					ExtensionSet("extensionAttribute12", value);
				}
			}
			// Create the extensionAttribute13 property.
			[DirectoryProperty("extensionAttribute13")]
			public string ExtensionAttribute13
			{
				get
				{
					if (ExtensionGet("extensionAttribute13").Length != 1)
						return null;
					return (string)ExtensionGet("extensionAttribute13")[0];
				}
				set
				{
					ExtensionSet("extensionAttribute13", value);
				}
			}
			// Create the extensionAttribute14 property.
			[DirectoryProperty("extensionAttribute14")]
			public string ExtensionAttribute14
			{
				get
				{
					if (ExtensionGet("extensionAttribute14").Length != 1)
						return null;
					return (string)ExtensionGet("extensionAttribute14")[0];
				}
				set
				{
					ExtensionSet("extensionAttribute14", value);
				}
			}
			// Create the extensionAttribute15 property.
			[DirectoryProperty("extensionAttribute15")]
			public string ExtensionAttribute15
			{
				get
				{
					if (ExtensionGet("extensionAttribute15").Length != 1)
						return null;
					return (string)ExtensionGet("extensionAttribute15")[0];
				}
				set
				{
					ExtensionSet("extensionAttribute15", value);
				}
			}
			// Create the facsimileTelephoneNumber property.
			[DirectoryProperty("facsimileTelephoneNumber")]
			public string FacsimileTelephoneNumber
			{
				get
				{
					if (ExtensionGet("facsimileTelephoneNumber").Length != 1)
						return null;
					return (string)ExtensionGet("facsimileTelephoneNumber")[0];
				}
				set
				{
					ExtensionSet("facsimileTelephoneNumber", value);
				}
			}
			// Create the floor property.
			[DirectoryProperty("floor")]
			public string Floor
			{
				get
				{
					if (ExtensionGet("floor").Length != 1)
						return null;
					return (string)ExtensionGet("floor")[0];
				}
				set
				{
					ExtensionSet("floor", value);
				}
			}
			// Create the generationQualifier property.
			[DirectoryProperty("generationQualifier")]
			public string GenerationQualifier
			{
				get
				{
					if (ExtensionGet("generationQualifier").Length != 1)
						return null;
					return (string)ExtensionGet("generationQualifier")[0];
				}
				set
				{
					ExtensionSet("generationQualifier", value);
				}
			}
			// Create the homeMDB property.
			[DirectoryProperty("homeMDB")]
			public string HomeMDB
			{
				get
				{
					if (ExtensionGet("homeMDB").Length != 1)
						return null;
					return (string)ExtensionGet("homeMDB")[0];
				}
				set
				{
					ExtensionSet("homeMDB", value);
				}
			}
			// Create the homeMTA property.
			[DirectoryProperty("homeMTA")]
			public string HomeMTA
			{
				get
				{
					if (ExtensionGet("homeMTA").Length != 1)
						return null;
					return (string)ExtensionGet("homeMTA")[0];
				}
				set
				{
					ExtensionSet("homeMTA", value);
				}
			}
			// Create the homeMTA property.
			[DirectoryProperty("homePhone")]
			public string HomePhone
			{
				get
				{
					if (ExtensionGet("homePhone").Length != 1)
						return null;
					return (string)ExtensionGet("homePhone")[0];
				}
				set
				{
					ExtensionSet("homePhone", value);
				}
			}
			// Create the info property.
			[DirectoryProperty("info")]
			public string Info
			{
				get
				{
					if (ExtensionGet("info").Length != 1)
						return null;
					return (string)ExtensionGet("info")[0];
				}
				set
				{
					ExtensionSet("info", value);
				}
			}
			// Create the sn property.
			[DirectoryProperty("sn")]
			public string sn
			{
				get
				{
					if (ExtensionGet("sn").Length != 1)
						return null;
					return (string)ExtensionGet("sn")[0];
				}
				set
				{
					ExtensionSet("sn", value);
				}
			}
			// Create the initials property.
			[DirectoryProperty("initials")]
			public string Initials
			{
				get
				{
					if (ExtensionGet("initials").Length != 1)
						return null;
					return (string)ExtensionGet("initials")[0];
				}
				set
				{
					ExtensionSet("initials", value);
				}
			}
			// Create the ipPhone property.
			[DirectoryProperty("ipPhone")]
			public string IPPhone
			{
				get
				{
					if (ExtensionGet("ipPhone").Length != 1)
						return null;
					return (string)ExtensionGet("ipPhone")[0];
				}
				set
				{
					ExtensionSet("ipPhone", value);
				}
			}
			// Create the l property.
			[DirectoryProperty("l")]
			public string City
			{
				get
				{
					if (ExtensionGet("l").Length != 1)
						return null;
					return (string)ExtensionGet("l")[0];
				}
				set
				{
					ExtensionSet("l", value);
				}
			}
			[DirectoryProperty("l")]		//hjs
			public string l
			{
				get
				{
					if (ExtensionGet("l").Length != 1)
						return null;
					return (string)ExtensionGet("l")[0];
				}
				set
				{
					ExtensionSet("l", value);
				}
			}
			// Create the legacyExchangeDN property.
			[DirectoryProperty("legacyExchangeDN")]
			public string LegacyExchangeDN
			{
				get
				{
					if (ExtensionGet("legacyExchangeDN").Length != 1)
						return null;
					return (string)ExtensionGet("legacyExchangeDN")[0];
				}
				set
				{
					ExtensionSet("legacyExchangeDN", value);
				}
			}
			// Create the mailNickName property.		//hjs
			[DirectoryProperty("mailNickName")]
			public string MailNickName
			{
				get
				{
					if (ExtensionGet("mailNickName").Length != 1)
						return null;
					return (string)ExtensionGet("mailNickName")[0];
				}
				set
				{
					ExtensionSet("mailNickName", value);
				}
			}
			// Create the manager property.
			[DirectoryProperty("manager")]
			public string Manager
			{
				get
				{
					if (ExtensionGet("manager").Length != 1)
						return null;
					return (string)ExtensionGet("manager")[0];
				}
				set
				{
					ExtensionSet("manager", value);
				}
			}
			// Create the mAPIRecipient property.
			[DirectoryProperty("mAPIRecipient")]
			public string MAPIRecipient
			{
				get
				{
					if (ExtensionGet("mAPIRecipient").Length != 1)
						return null;
					return (string)ExtensionGet("mAPIRecipient")[0];
				}
				set
				{
					ExtensionSet("mAPIRecipient", value);
				}
			}
			// Create the mobile phone property.
			[DirectoryProperty("mobile")]
			public string MobilePhone
			{
				get
				{
					if (ExtensionGet("mobile").Length != 1)
						return null;
					return (string)ExtensionGet("mobile")[0];
				}
				set
				{
					ExtensionSet("mobile", value);
				}
			}
			// Create the telephoneNumber property.		//hjs
			[DirectoryProperty("telephoneNumber")]
			public string telephoneNumber
			{
				get
				{
					if (ExtensionGet("telephoneNumber").Length != 1)
						return null;
					return (string)ExtensionGet("telephoneNumber")[0];
				}
				set
				{
					ExtensionSet("telephoneNumber", value);
				}
			}
			// Create the AccountDisabled property.		//hjs - Adding this does not allow the checkbox to work, even if I use "AccountDisabled" as the name.
			[DirectoryProperty("Enabled")]
			public string AccountDisabled
			{
				get
				{
					if (ExtensionGet("Enabled").Length != 1)
						return null;
					return (string)ExtensionGet("Enabled")[0];
				}
				set
				{
					ExtensionSet("Enabled", value);
				}
			}
			[DirectoryProperty("mail")]		//hjs
			public string mail
			{
				get
				{
					if (ExtensionGet("mail").Length != 1)
						return null;
					return (string)ExtensionGet("mail")[0];
				}
				set
				{
					ExtensionSet("mail", value);
				}
			}
			// Create the mDBUseDefaults property.		//hjs
			[DirectoryProperty("mDBUseDefaults")]
			public string mDBUseDefaults
			{
				get
				{
					if (ExtensionGet("mDBUseDefaults").Length != 1)
						return null;
					return (string)ExtensionGet("mDBUseDefaults")[0];
				}
				set
				{
					ExtensionSet("mDBUseDefaults", value);
				}
			}
			// Create the msExchAssistantName property.
			[DirectoryProperty("msExchAssistantName")]
			public string MSxchAssistantName
			{
				get
				{
					if (ExtensionGet("msExchAssistantName").Length != 1)
						return null;
					return (string)ExtensionGet("msExchAssistantName")[0];
				}
				set
				{
					ExtensionSet("msExchAssistantName", value);
				}
			}
			// Create the msExchExpansionServerName property.
			[DirectoryProperty("msExchExpansionServerName")]
			public string MSExchExpansionServerName
			{
				get
				{
					if (ExtensionGet("msExchExpansionServerName").Length != 1)
						return null;
					return (string)ExtensionGet("msExchExpansionServerName")[0];
				}
				set
				{
					ExtensionSet("msExchExpansionServerName", value);
				}
			}
			// Create the msExchHideFromAddressList property.
			[DirectoryProperty("msExchHideFromAddressLists")]
			public bool MSExchHideFromAddressLists
			{
				get
				{
					try
					{
						return (bool)ExtensionGet("msExchHideFromAddressList")[0];
					}
					catch
					{
						Console.WriteLine("Error querying msExchHideFromAddressList");
						return false;
					}
				}
				set
				{
					ExtensionSet("msExchHideFromAddressList",  value);
				}
			}
			// Create the msExchHomeServerName property.
			[DirectoryProperty("msExchHomeServerName")]
			public string MSExchHomeServerName
			{
				get
				{
					if (ExtensionGet("msExchHomeServerName").Length != 1)
						return null;
					return (string)ExtensionGet("msExchHomeServerName")[0];
				}
				set
				{
					ExtensionSet("msExchHomeServerName", value);
				}
			}
			// Create the msExchMasterAccountSid property.
			[DirectoryProperty("msExchMasterAccountSid")]
			public string MSExchMasterAccountSid
			{
				get
				{
					if (ExtensionGet("msExchMasterAccountSid").Length != 1)
						return null;
					return (string)ExtensionGet("msExchMasterAccountSid")[0];
				}
				set
				{
					ExtensionSet("msExchMasterAccountSid", value);
				}
			}
			// Create the other msExchOriginatingForest property.  
			[DirectoryProperty("msExchOriginatingForest")]
			public string[] MSExchOriginatingForest
			{
				get
				{
					int len = ExtensionGet("msExchOriginatingForest").Length;
					if (len == 0) return null;
					string[] msExchOriginatingForest = new string[len];
					object[] msExchOriginatingForestRaw = ExtensionGet("msExchOriginatingForest");
					for (int i = 0; i < len; i++)
					{
						msExchOriginatingForest[i] = (string)msExchOriginatingForestRaw[i];
					}
					return msExchOriginatingForest;
				}
				set
				{
					ExtensionSet("msExchOriginatingForest", value);
				}
			}
			// Create the nMCIAssetTag property.
			[DirectoryProperty("nMCIAssetTag")]
			public string NMCIAssetTag
			{
				get
				{
					if (ExtensionGet("nMCIAssetTag").Length != 1)
						return null;
					return (string)ExtensionGet("nMCIAssetTag")[0];
				}
				set
				{
					ExtensionSet("nMCIAssetTag", value);
				}
			}
			// Create the pager property.
			[DirectoryProperty("pager")]
			public string Pager
			{
				get
				{
					if (ExtensionGet("pager").Length != 1)
						return null;
					return (string)ExtensionGet("pager")[0];
				}
				set
				{
					ExtensionSet("pager", value);
				}
			}
			// Create the personalTitle property.
			[DirectoryProperty("personalTitle")]
			public string PersonalTitle
			{
				get
				{
					if (ExtensionGet("personalTitle").Length != 1)
						return null;
					return (string)ExtensionGet("personalTitle")[0];
				}
				set
				{
					ExtensionSet("personalTitle", value);
				}
			}
			// Create the physicalDeliveryOfficeName property.
			[DirectoryProperty("physicalDeliveryOfficeName")]
			public string PhysicalDeliveryOfficeName
			{
				get
				{
					if (ExtensionGet("physicalDeliveryOfficeName").Length != 1)
						return null;
					return (string)ExtensionGet("physicalDeliveryOfficeName")[0];
				}
				set
				{
					ExtensionSet("physicalDeliveryOfficeName", value);
				}
			}
			// Create the postalCode property.
			[DirectoryProperty("postalCode")]
			public string PostalCode
			{
				get
				{
					if (ExtensionGet("postalCode").Length != 1)
						return null;
					return (string)ExtensionGet("postalCode")[0];
				}
				set
				{
					ExtensionSet("postalCode", value);
				}
			}
			// Create the profilePath property.
			[DirectoryProperty("profilePath")]
			public string ProfilePath
			{
				get
				{
					if (ExtensionGet("profilePath").Length != 1)
						return null;
					return (string)ExtensionGet("profilePath")[0];
				}
				set
				{
					ExtensionSet("profilePath", value);
				}
			}
			// Create the roomCube property.
			[DirectoryProperty("roomCube")]
			public string RoomCube
			{
				get
				{
					if (ExtensionGet("roomCube").Length != 1)
						return null;
					return (string)ExtensionGet("roomCube")[0];
				}
				set
				{
					ExtensionSet("roomCube", value);
				}
			}
			// Create the other seeAlso property.  
			[DirectoryProperty("seeAlso")]
			public string[] SeeAlso
			{
				get
				{
					int len = ExtensionGet("seeAlso").Length;
					if (len == 0) return null;
					string[] seeAlso = new string[len];
					object[] seeAlsoRaw = ExtensionGet("seeAlso");
					for (int i = 0; i < len; i++)
					{
						seeAlso[i] = (string)seeAlsoRaw[i];
					}
					return seeAlso;
				}
				set
				{
					ExtensionSet("seeAlso", value);
				}
			}
			// Create the st property.
			[DirectoryProperty("st")]
			public string State
			{
				get
				{
					if (ExtensionGet("st").Length != 1)
						return null;
					return (string)ExtensionGet("st")[0];
				}
				set
				{
					ExtensionSet("st", value);
				}
			}
			[DirectoryProperty("st")]		//hjs
			public string st
			{
				get
				{
					if (ExtensionGet("st").Length != 1)
						return null;
					return (string)ExtensionGet("st")[0];
				}
				set
				{
					ExtensionSet("st", value);
				}
			}
			// Create the streetAddress property.
			[DirectoryProperty("streetAddress")]
			public string StreetAddress
			{
				get
				{
					if (ExtensionGet("streetAddress").Length != 1)
						return null;
					return (string)ExtensionGet("streetAddress")[0];
				}
				set
				{
					ExtensionSet("streetAddress", value);
				}
			}
			// Create the terminalServer property.
			[DirectoryProperty("terminalServer")]
			public string TerminalServer
			{
				get
				{
					if (ExtensionGet("terminalServer").Length != 1)
						return null;
					return (string)ExtensionGet("terminalServer")[0];
				}
				set
				{
					ExtensionSet("terminalServer", value);
				}
			}
			// Create the title property.
			[DirectoryProperty("title")]
			public string Title
			{
				get
				{
					if (ExtensionGet("title").Length != 1)
						return null;
					return (string)ExtensionGet("title")[0];
				}
				set
				{
					ExtensionSet("title", value);
				}
			}
			// Create the uAITChanged property.
			[DirectoryProperty("uAITChanged")]
			public string UAITChanged
			{
				get
				{
					if (ExtensionGet("uAITChanged").Length != 1)
						return null;
					return (string)ExtensionGet("uAITChanged")[0];
				}
				set
				{
					ExtensionSet("uAITChanged", value);
				}
			}
			// Create the uIC property.
			[DirectoryProperty("uIC")]
			public string UIC
			{
				get
				{
					if (ExtensionGet("uIC").Length != 1)
						return null;
					return (string)ExtensionGet("uIC")[0];
				}
				set
				{
					ExtensionSet("uIC", value);
				}
			}
			// Create the unicodePwd property.
			[DirectoryProperty("unicodePwd")]
			public string UnicodePwd
			{
				get
				{
					if (ExtensionGet("unicodePwd").Length != 1)
						return null;
					return (string)ExtensionGet("unicodePwd")[0];
				}
				set
				{
					ExtensionSet("unicodePwd", value);
				}
			}
			// Create the userParameters property.
			[DirectoryProperty("userParameters")]
			public string UserParameters
			{
				get
				{
					if (ExtensionGet("userParameters").Length != 1)
						return null;
					return (string)ExtensionGet("userParameters")[0];
				}
				set
				{
					ExtensionSet("userParameters", value);
				}
			}
			// Create the primaryComputer property.
			[DirectoryProperty("primaryComputer")]
			public string primaryComputer
			{
				get
				{
					if (ExtensionGet("primaryComputer").Length != 1)
						return null;
					return (string)ExtensionGet("primaryComputer")[0];
				}
				set
				{
					ExtensionSet("primaryComputer", value);
				}
			}
			// Create the userWorkstations property.
			[DirectoryProperty("userWorkstations")]
			public string UserWorkstations
			{
				get
				{
					if (ExtensionGet("userWorkstations").Length != 1)
						return null;
					return (string)ExtensionGet("userWorkstations")[0];
				}
				set
				{
					ExtensionSet("userWorkstations", value);
				}
			}
			// Create the wWWHomePage property.
			[DirectoryProperty("wWWHomePage")]
			public string WWWHomePage
			{
				get
				{
					if (ExtensionGet("wWWHomePage").Length != 1)
						return null;
					return (string)ExtensionGet("wWWHomePage")[0];
				}
				set
				{
					ExtensionSet("wWWHomePage", value);
				}
			}
			//Create the formData property.
			[DirectoryProperty("formData")]
			public string FormData
			{
				get
				{
					if(ExtensionGet("formData").Length != 1)
						return null;
					return System.Text.Encoding.UTF8.GetString((byte[])ExtensionGet("formData")[0]);
				}
				set
				{
					((DirectoryEntry)this.GetUnderlyingObject()).Properties["formData"].Value = (System.Text.Encoding.UTF8).GetBytes(value);
				}
			}
			// Create the msTSProfilePath property.
			[DirectoryProperty("msTSProfilePath")]
			public string msTSProfilePath
			{
				get
				{
					if (ExtensionGet("msTSProfilePath").Length != 1)
						return null;
					return (string)ExtensionGet("msTSProfilePath")[0];
				}
				set 
				{
					ExtensionSet("msTSProfilePath", value);
				}
			}
			// Create the msTSAllowLogon property.
			[DirectoryProperty("msTSAllowLogon")]
			public string msTSAllowLogon
			{
				get
				{
					if (ExtensionGet("msTSAllowLogon").Length != 1)
						return null;
					return (string)ExtensionGet("msTSAllowLogon")[0];
				}
				set 
				{
					ExtensionSet("msTSAllowLogon", value);
				}
			}
			// Create the msTSHomeDirectory property.
			[DirectoryProperty("msTSHomeDirectory")]
			public string msTSHomeDirectory
			{
				get
				{
					if (ExtensionGet("msTSHomeDirectory").Length != 1)
						return null;
					return (string)ExtensionGet("msTSHomeDirectory")[0];
				}
				set 
				{
					ExtensionSet("msTSHomeDirectory", value);
				}
			}
			// Create the msTSHomeDrive property.
			[DirectoryProperty("msTSHomeDrive")]
			public string msTSHomeDrive
			{
				get
				{
					if (ExtensionGet("msTSHomeDrive").Length != 1)
						return null;
					return (string)ExtensionGet("msTSHomeDrive")[0];
				}
				set 
				{
					ExtensionSet("msTSHomeDrive", value);
				}
			}

			// Implement the overloaded search method FindByIdentity.
			public static new NMCIUserPrincipal FindByIdentity(PrincipalContext context,
														   string identityValue)
			{
				return (NMCIUserPrincipal)FindByIdentityWithType(context,
															 typeof(NMCIUserPrincipal),
															 identityValue);
			}
			// Implement the overloaded search method FindByIdentity. 
			public static new NMCIUserPrincipal FindByIdentity(PrincipalContext context,
														   IdentityType identityType,
														   string identityValue)
			{
				return (NMCIUserPrincipal)FindByIdentityWithType(context,
															 typeof(NMCIUserPrincipal),
															 identityType,
															 identityValue);
			}
		}
	}
"@

	$assemblies = @('System.DirectoryServices', 'System.DirectoryServices.AccountManagement')
	Add-Type -TypeDefinition $cs -Language 'CSharp' -ReferencedAssemblies $assemblies -IgnoreWarnings
}
