####################################################
# Updated Date:		10 September 2015
# Purpose:			Group (AD and Exchange) routines.
# Requirements:		AddUserToGroup(), CreateGroup() require "AD-Routines.ps1".
####################################################

	function AddUserToGroup{
		#Adds a User/computer to a Group as a Member.
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$GroupName, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$UserName, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$DomainOrDC
		)
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
		}else{
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
			}else{
				#DL
				#Import exchange commands for the DL actions.
				$Session = Get-PSSession | Select Name;
				if (($Session -ne "") -and ($Session -ne $null)){
					#Write-Host "have at least one session";
				}else{
					if (!(Get-Command "SetupConn" -ErrorAction SilentlyContinue)){
						$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;
						if ((Test-Path ($ScriptDir + "\Exchange.ps1"))){
							. ($ScriptDir + "\Exchange.ps1")
						}
					}
					SetupConn "w" "Random";
				}

				if (([String]($objGroup).DistinguishedName -eq "") -or ([String]($objGroup).DistinguishedName -eq $null)){
					$strGroupDN = [String]$objGroup;
				}else{
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
			}else{
				$objReturn.Results = 1;
				$strMessage = "Success";
			}
		}else{
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
		}else{
			$objGroup =  $(Try {Get-ADGroup -Identity $GroupName -Server $DomainOrDC;} Catch {$null});
		}

		#Check if found an existing Group
		if (($objGroup -ne "") -and ($objGroup -ne $null)){
			#Found an existing Group
			$objReturn.Results = $False;
			$strResults = "Error Found a Group named '" + $GroupName + "' already exists.`r`n" + ($objGroup.DistinguishedName)
		}else{
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
							if ((Test-Path ($ScriptDir + "\Exchange.ps1"))){
								. ($ScriptDir + "\Exchange.ps1")
							}
						}
						if ($Session.State -ne "Opened"){
							#CleanUpConn;
						}
						SetupConn "w" "Random";
					}else{
						#Write-Host "have at least one session";
						#if ($Session -is [array]){
						#	For ($i=0; $i -lt $Session.length; $i++){
						#		Write-Host $Session[$i].Name;
						#	}
						#}else{
						#	$Session = (Get-PSSession).Name;
						#	Write-Host "Session is: " $Session;
						#}
					}

					if (($GroupAlias -eq "") -or ($GroupAlias -eq $null)){
						if (($GroupDisp -eq "") -or ($GroupDisp -eq $null)){
							$GroupAlias = $GroupName;
							$GroupDisp = $GroupName;
						}else{
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
						}else{
							if ($Members){
								#NO Manager & Has Users
									$strResults = New-DistributionGroup -Name $GroupName -Type "Distribution" -DomainController $DomainOrDC -DisplayName $GroupDisp -SamAccountName $GroupName -OrganizationalUnit $OUPath -Alias $GroupAlias -Notes $GroupNotes -Members $Members;
									#$objJobCode = [scriptblock]::create({param($strGroupName, $strOpsMaster, $strGroupDisplayName, $strPath, $strGroupAlias, $strGroupEmail, $strGroupNotes, $strGrpMembers); New-DistributionGroup -Name $strGroupName -Type "Distribution" -DomainController $strOpsMaster -DisplayName $strGroupDisplayName -SamAccountName $strGroupName -OrganizationalUnit $strPath -Alias $strGroupAlias -PrimarySmtpAddress $strGroupEmail -Notes $strGroupNotes -Members $strGrpMembers;});
									#$arrArgs = @($GroupName, $strGroupDomain, $GroupDisp, $OUPath, $GroupAlias, ($GroupAlias + "@navy.mil"), $GroupNotes, $Members);
							}else{
								#NO Manager & NO Users
									$strResults = New-DistributionGroup -Name $GroupName -Type "Distribution" -DomainController $DomainOrDC -DisplayName $GroupDisp -SamAccountName $GroupName -OrganizationalUnit $OUPath -Alias $GroupAlias -Notes $GroupNotes;
									#$objJobCode = [scriptblock]::create({param($strGroupName, $strOpsMaster, $strGroupDisplayName, $strPath, $strGroupAlias, $strGroupEmail, $strGroupNotes); New-DistributionGroup -Name $strGroupName -Type "Distribution" -DomainController $strOpsMaster -DisplayName $strGroupDisplayName -SamAccountName $strGroupName -OrganizationalUnit $strPath -Alias $strGroupAlias -PrimarySmtpAddress $strGroupEmail -Notes $strGroupNotes;});
									#$arrArgs = @($GroupName, $strGroupDomain, $GroupDisp, $OUPath, $GroupAlias, $strGroupFullEmail, $GroupNotes);
							}
						}
					}else{
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
							}else{
								#Has Manager & NO Users
									$strResults = New-DistributionGroup -Name $GroupName -Type "Security" -DomainController $DomainOrDC -DisplayName $GroupDisp -SamAccountName $GroupName -OrganizationalUnit $OUPath -Alias $GroupAlias -Notes $GroupNotes -ManagedBy $ManagedBy;
									#$objJobCode = [scriptblock]::create({param($strGroupName, $strOpsMaster, $strGroupDisplayName, $strPath, $strGroupAlias, $strGroupEmail, $strGroupNotes, $strManagedBy); New-DistributionGroup -Name $strGroupName -Type "Security" -DomainController $strOpsMaster -DisplayName $strGroupDisplayName -SamAccountName $strGroupName -OrganizationalUnit $strPath -Alias $strGroupAlias -PrimarySmtpAddress $strGroupEmail -Notes $strGroupNotes -ManagedBy $strManagedBy;});
									#$arrArgs = @($GroupName, $strGroupDomain, $GroupDisp, $OUPath, $GroupAlias, ($GroupAlias + "@navy.mil"), $GroupNotes, $ManagedBy);
							}
						}else{
							if ($Members){
								#NO Manager & Has Users
									$strResults = New-DistributionGroup -Name $GroupName -Type "Security" -DomainController $DomainOrDC -DisplayName $GroupDisp -SamAccountName $GroupName -OrganizationalUnit $OUPath -Alias $GroupAlias -Notes $GroupNotes -Members $Members;
									#$objJobCode = [scriptblock]::create({param($strGroupName, $strOpsMaster, $strGroupDisplayName, $strPath, $strGroupAlias, $strGroupEmail, $strGroupNotes, $strGrpMembers); New-DistributionGroup -Name $strGroupName -Type "Security" -DomainController $strOpsMaster -DisplayName $strGroupDisplayName -SamAccountName $strGroupName -OrganizationalUnit $strPath -Alias $strGroupAlias -PrimarySmtpAddress $strGroupEmail -Notes $strGroupNotes -Members $strGrpMembers;});
									#$arrArgs = @($GroupName, $strGroupDomain, $GroupDisp, $OUPath, $GroupAlias, ($GroupAlias + "@navy.mil"), $GroupNotes, $Members);
							}else{
								#NO Manager & NO Users
									$strResults = New-DistributionGroup -Name $GroupName -Type "Security" -DomainController $DomainOrDC -DisplayName $GroupDisp -SamAccountName $GroupName -OrganizationalUnit $OUPath -Alias $GroupAlias -Notes $GroupNotes;
									#$objJobCode = [scriptblock]::create({param($strGroupName, $strOpsMaster, $strGroupDisplayName, $strPath, $strGroupAlias, $strGroupEmail, $strGroupNotes); New-DistributionGroup -Name $strGroupName -Type "Security" -DomainController $strOpsMaster -DisplayName $strGroupDisplayName -SamAccountName $strGroupName -OrganizationalUnit $strPath -Alias $strGroupAlias -PrimarySmtpAddress $strGroupEmail -Notes $strGroupNotes;});
									#$arrArgs = @($GroupName, $strGroupDomain, $GroupDisp, $OUPath, $GroupAlias, ($GroupAlias + "@navy.mil"), $GroupNotes);
							}
						}
					}
				}else{
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
								}else{
									$strResults = New-ADGroup -Name $GroupName -GroupScope $Scope -Server $DomainOrDC -SamAccountName $GroupName -Path $OUPath -ManagedBy $ManagedBy -OtherAttributes @{'info'=$GroupNotes};
								}
								#$objJobCode = [scriptblock]::create({param($strGroupName, $strGrpType, $strOpsMaster, $strPath, $strManagedBy); New-ADGroup -Name $strGroupName -GroupScope $strGrpType -Server $strOpsMaster -SamAccountName $strGroupName -Path $strPath -ManagedBy $strManagedBy;});
								#$arrArgs = @($GroupName, $strGroupScope, $strGroupDomain, $OUPath, $ManagedBy);
						}else{
							#NO Manager
								#$strResults = New-ADGroup -Name $GroupName -GroupScope $Scope -Server $DomainOrDC -SamAccountName $GroupName -Path $OUPath;
								if ($GroupType -eq "Distribution"){
									$strResults = New-ADGroup -Name $GroupName -GroupScope $Scope -GroupCategory "Distribution" -Server $DomainOrDC -SamAccountName $GroupName -Path $OUPath -OtherAttributes @{'info'=$GroupNotes};
								}else{
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
				}else{
					$objReturn.Results = $True;
					$objReturn.Message = "Success";
					$objReturn.Message = $objReturn.Message + $strMessage;
					if (($strResults -ne "") -and ($strResults -ne $null)){
						$objReturn.Returns = $strResults;
					}
				}
			}else{
				$objReturn.Results = $False;

				if ($objOUReturn.Results -gt 1){
					$strTemp = $objOUReturn.Returns[0];
					for ($intX = 1; $intX -lt $objOUReturn.Results; $intX++){
						$strTemp = $strTemp + ", " + $objOUReturn.Returns[$intX];
					}

					$strResults = "The OU provided, to create the group in, was found on multiple Domains.`r`n $strTemp";
				}else{
					#OU path provided does not exist.
					$strResults = "The OU path provided, to create the group in, could not be found found on any available Domains.";
				}

				$objReturn.Message = $strResults;
			}
		}

		return $objReturn;
	}

	function GetGroups{
		#Based heavily on code from:
		#http://www.reich-consulting.net/2013/12/05/retrieving-recursive-group-memberships-powershell/
		#Which we got to from:
		#http://stackoverflow.com/questions/5072996/how-to-get-all-groups-that-a-user-is-a-member-of

		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)]$ADObject, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Array]$arrList, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bolRecurse = $False
		)
		#$ADObject = An AD object, or the sAMAccountName (String) of the AD object to get.
		#$arrList = The Array, of strings, that will be updated/returned, that will have the list of Memberships $ADObject has.
		#$bolRecurse = Get the Groups any Groups are Members Of as well.

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
			}else{
				#$objADObject = (Get-ADUser -Identity $ADObject -Properties MemberOf, DistinguishedName, sAMAccountName | Select-Object MemberOf, DistinguishedName, sAMAccountName);
				$objADObject = $(Try {(Get-ADUser -Identity $ADObject -Properties MemberOf, DistinguishedName, sAMAccountName | Select-Object MemberOf, DistinguishedName, sAMAccountName)} Catch {$null});
			}

			if (($objADObject -eq $null) -or ($objADObject -eq "")){
				#Could not find an AD Object, check if it is a Group.
				if ($ADObject.Contains("\")){
					$objADObject = $(Try {Get-ADObject $ADObject.Split("\")[-1] -Server $ADObject.Split("\")[0] -Properties MemberOf, sAMAccountName, DistinguishedName} Catch {$null});
				}else{
					$objADObject = $(Try {Get-ADObject $ADObject -Properties MemberOf, sAMAccountName, DistinguishedName} Catch {$null});
				}
			}

			if (($objADObject -eq $null) -or ($objADObject -eq "")){
				#Could not find an AD Object, check if it is a Machine.
				if ($ADObject.Contains("\")){
					$objADObject = $(Try {Get-ADComputer -Identity $ADObject.Split("\")[-1] -Server $ADObject.Split("\")[0] -Properties MemberOf, sAMAccountName, DistinguishedName} Catch {$null});
				}else{
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
				}else{
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
								}else{
									if (!(($strRIDMasterW -ne $null) -and ($strRIDMasterW -ne ""))){
										$strStatus = CheckRunSpaceJob "GetRIDMasterW" $global:objJobs;
										if ($strStatus -ne "Failed"){
											$strRIDMasterW = WaitForRunSpaceJob "GetRIDMasterW" $global:objJobs $txbRidW;
										}else{
											$strRIDMasterW = (Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
											$txbRidW.Text = $strRIDMasterW;
										}
									#}else{
										#Have $strRIDMasterW already.
									}
								}
							}else{
								#This Function is probably running in the background.
								if (!(($strRIDMasterW -ne $null) -and ($strRIDMasterW -ne ""))){
									$strRIDMasterW = (Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
								#}else{
									#Have $strRIDMasterW already.
								}
							}
							$strRIDMaster = $strRIDMasterW;
						}
						"nadsusea"{
							if ($txbRidE -ne $null){
								if ($txbRidE.Text -ne ""){
									$strRIDMasterE = $txbRidE.Text.Trim();
								}else{
									if (!(($strRIDMasterE -ne $null) -and ($strRIDMasterE -ne ""))){
										$strStatus = CheckRunSpaceJob "GetRIDMasterE" $global:objJobs;
										if ($strStatus -ne "Failed"){
											$strRIDMasterE = WaitForRunSpaceJob "GetRIDMasterE" $global:objJobs $txbRidE;
										}else{
											$strRIDMasterE = (Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
											$txbRidE.Text = $strRIDMasterE;
										}
									#}else{
										#Have $strRIDMasterE already.
									}
								}
							}else{
								#This Function is probably running in the background.
								if (!(($strRIDMasterE -ne $null) -and ($strRIDMasterE -ne ""))){
									$strRIDMasterE = (Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
								#}else{
									#Have $strRIDMasterE already.
								}
							}
							$strRIDMaster = $strRIDMasterE;
						}
						"pads"{
							if ($txbRidP -ne $null){
								if ($txbRidP.Text -ne ""){
									$strRIDMasterP = $txbRidP.Text.Trim();
								}else{
									if (!(($strRIDMasterP -ne $null) -and ($strRIDMasterP -ne ""))){
										$strStatus = CheckRunSpaceJob "GetRIDMasterP" $global:objJobs;
										if ($strStatus -ne "Failed"){
											$strRIDMasterP = WaitForRunSpaceJob "GetRIDMasterP" $global:objJobs $txbRidP;
										}else{
											$strRIDMasterP = (Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
											$txbRidP.Text = $strRIDMasterP;
										}
									#}else{
										#Have $strRIDMasterP already.
									}
								}
							}else{
								#This Function is probably running in the background.
								if (!(($strRIDMasterP -ne $null) -and ($strRIDMasterP -ne ""))){
									$strRIDMasterP = (Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
								#}else{
									#Have $strRIDMasterP already.
								}
							}
							$strRIDMaster = $strRIDMasterP;
						}
						"nmci-isf"{
							if ($txbRidN -ne $null){
								if ($txbRidN.Text -ne ""){
									$strRIDMasterN = $txbRidN.Text.Trim();
								}else{
									if (!(($strRIDMasterN -ne $null) -and ($strRIDMasterN -ne ""))){
										$strStatus = CheckRunSpaceJob "GetRIDMasterN" $global:objJobs;
										if ($strStatus -ne "Failed"){
											$strRIDMasterN = WaitForRunSpaceJob "GetRIDMasterN" $global:objJobs $txbRidN;
										}else{
											$strRIDMasterN = (Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
											$txbRidN.Text = $strRIDMasterN;
										}
									#}else{
										#Have $strRIDMasterN already.
									}
								}
							}else{
								#This Function is probably running in the background.
								if (!(($strRIDMasterN -ne $null) -and ($strRIDMasterN -ne ""))){
									$strRIDMasterN = (Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
								#}else{
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
					}else{
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
		}else{
			$arrList += "Error Could not find $ADObject in AD.";
		}
		#$strPSCmds = $strPSCmds.Replace(", ", ",`r`n");

		#$arrList = $arrList | Sort-Object;

		return $arrList;
	}

