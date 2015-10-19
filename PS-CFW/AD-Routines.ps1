###########################################
# Updated Date:	17 July 2015
# Purpose:		Provide a central location for all the PowerShell Active Directory routines.
# Requirements: For the PInvoked Code .NET 4+ is required.
##########################################

	function TestRoutine{
		#Some users are having issues that the command "(Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster" is NOT getting the RIDMaster, 
			#returns an error simular to "internal error, the server is busy."
		#In talking w/ Chris he suggested the .NET methods instead.
		#http://mikefrobbins.com/2013/04/18/powershell-function-to-determine-the-active-directory-fsmo-role-holders-via-the-net-framework/

		#. "\\nawesdnifs08.nadsuswe.nads.navy.mil\NMCIISF\NMCIISF-SDCP-MAC\MAC\Entr_SRM\Support Files\PS-Scripts\PS-AD-Routines.ps1"

		$InitializeDefaultDrives=$False;
		if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};

		$strDomain = "nadsuswe"
		$Error.Clear();
		(Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
		if ($Error){
			Write-Host "WOOT it errored, so now try the .NET method."
			[System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $strDomain))).RidRoleOwner.Name;
		}

		$strDomain = "nadsusea"
		$Error.Clear();
		(Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
		if ($Error){
			Write-Host "WOOT it errored, so now try the .NET method."
			[System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $strDomain))).RidRoleOwner.Name;
		}

		#So then my RIDMaster commands go from:
		#$strRIDMaster = (Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
		#to:
		#$strRIDMaster = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $strDomain))).RidRoleOwner.Name;
	}


	function CreateADComputer{
		#Note: If the SAMAccountName string provided, does not end with a '$', one will be appended (by powershell) if needed.
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strCompName, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strOU, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strDomain, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strDC = "", 
			[ValidateNotNull()][Parameter(Mandatory=$False)]$objADInfo = $null
		)
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
		#$strTemp = "CreateADComputer(" + $strTemp.Trim() + ")";
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
		}else{
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
				#$strDomain = $objResults.Returns[0];
				if ($objResults.Results -gt 1){
					#"The OU was found on multiple Domains."
					$bMatch = $False;
					for ($intX = 1; $intX -lt $objResults.Results; $intX++){
						#$strTemp = $strTemp + ", " + $objResults.Returns[$intX];
						if (($strDomain -eq $objResults.Returns[$intX])){
							#One of the domains found matches the requested domain.
							$bMatch = $True;
							break;
						}
					}
					if ($bMatch -eq $False){
						$objReturn.Message = "OU found on multiple domains, but none of them match the domain supplied.";
						$strDomain = "";
					}
				}else{
					if (!($strDomain -eq $objResults.Returns[0])){
						$objReturn.Message = "OU found, but the domain it was found on does NOT match the supplied domain.";
						$strDomain = "";
					}
				}

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
						#Check if a propertie exists:
						#if ($objADInfo.PSObject.Properties.Match('Test1').Count) {Write-Host "True"} else {Write-Host "False"};

						if (($strProp.Name -ne "") -and ($strProp.Name -ne $null)){
							if (($strProp.Name -eq "SamAccountName") -or ($strProp.Name -eq "Path") -or ($strProp.Name -eq "Server")){
								#Skip these ones.
									#SamAccountName
									#Path
									#Server
							}else{
								if (($strProp.Value -eq $True) -or ($strProp.Value -eq "True") -or ($strProp.Value -eq $False) -or ($strProp.Value -eq "False")){
									if (($strProp.Value -eq $True) -or ($strProp.Value -eq "True")){
										$strPSCmd = $strPSCmd + " -" + $strProp.Name + " $" + $True;
									}else{
										$strPSCmd = $strPSCmd + " -" + $strProp.Name + " $" + $False;
									}
								}else{
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
					}else{
						$objReturn.Results = $True;
						$objReturn.Message = "Success";

						#Now get the new object and return it.
						$objComp = $(Try {Get-ADComputer -Identity $strCompName -Server $strDC -Properties *} Catch {$null});
						if (($objComp.DistinguishedName -ne "") -and ($objComp.DistinguishedName -ne $null)){
							$objReturn.Returns = $objComp;
						}
					}
				}
			}else{
				$objReturn.Message = "OU was not found on any available domains.";
			}
		}

		return $objReturn;
	}

	function CreateADUser{
		#http://www.howtogeek.com/50187/how-to-create-multiple-users-in-server-2008-with-powershell/
		#$objOU = [ADSI]"LDAP://[DomainController/]OU=People,DC=sysadmingeek,DC=com"
		#Taking about 6 seconds in my initial testing.
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)]$oADInfo, 
			[ValidateNotNull()][Parameter(Mandatory=$True)]$strOU, 
			[ValidateNotNull()][Parameter(Mandatory=$True)]$strDomain, 
			[ValidateNotNull()][Parameter(Mandatory=$False)]$strDC
		)
		#Adds/Updates $oADInfo.TheResults with the results of the Create process.
		#$oADInfo = Cutom PowerShell Object that has all the AD fields to be set.
			#$oADInfo = New-Object PSObject;
			#Add-Member -InputObject $oADInfo -MemberType NoteProperty -Name "CN" -Value "";
			#Add-Member -InputObject $oADInfo -MemberType NoteProperty -Name "sAMAccountName" -Value "";
			#Add-Member -InputObject $oADInfo -MemberType NoteProperty -Name "userPrincipalName" -Value "";
			#Add-Member -InputObject $oADInfo -MemberType NoteProperty -Name "givenName" -Value "";
			#Add-Member -InputObject $oADInfo -MemberType NoteProperty -Name "SN" -Value "";
		#$strOU = The LDAP OU path. i.e. "OU=USERS,OU=BASE,OU=CMD".
		#$strDomain = The Domain to create the account on.  i.e. "sysadmingeek", or "sysadmingeek.com".
		#$strDC = The Domain Controller to create the account at.  FQDN or just the server name.

		$strMessage = "";

		#MUST have the following fields, no matter what, to create a User Object.
		if ((($oADInfo.CN -eq "") -or ($oADInfo.CN -eq $null)) -or (($oADInfo.SN -eq "") -or ($oADInfo.SN -eq $null)) -or (($oADInfo.givenName -eq "") -or ($oADInfo.givenName -eq $null)) -or (($oADInfo.userPrincipalName -eq "") -or ($oADInfo.userPrincipalName -eq $null)) -or (($oADInfo.sAMAccountName -eq "") -or ($oADInfo.sAMAccountName -eq $null))){
			#CN, sAMAccountName, userPrincipalName, givenName (First), SN (Last)
			$strMessage = "Required AD fields are missing.`r`n(CN, sAMAccountName, userPrincipalName, givenName, SN)";
			Add-Member -InputObject $oADInfo -MemberType NoteProperty -Name "TheResults" -Value $strMessage -Force;
			return;
		}

		$InitializeDefaultDrives=$False;
		if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};

		#$objOU = [ADSI]"LDAP://DomainController/OU=USERS,OU=NRFK,OU=NAVRESFOR,DC=nadsusea,DC=nads,DC=navy,DC=mil"
		#if (($strDC -eq "") -or ($strDC -eq $null)){
		#	#$objOU = [ADSI]"LDAP://OU=USERS,OU=NRFK,OU=NAVRESFOR,DC=nadsusea,DC=nads,DC=navy,DC=mil";
		#}else{
		#	#$objOU = [ADSI]"LDAP://$strDC/OU=USERS,OU=NRFK,OU=NAVRESFOR,DC=nadsusea,DC=nads,DC=navy,DC=mil";
		#}

		#Get the Domain DistinguishedName from $strDomain.
		if (!($strMessage.StartsWith("DC="))){
			$strDomain = (Get-ADDomain $strDomain).DistinguishedName;
		}

		$objOU = [ADSI]"LDAP://";
		if (($strDC -ne "") -and ($strDC -ne $null)){
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
				#$oProp.Name + " = " + $oProp.Value;
				$objUser.Put($oProp.Name, $oProp.Value);
			}
		}
		#$objUser.Put("sAMAccountName", $sAMAccountName);
		#$objUser.Put("userPrincipalName", $userPrincipalName);
		#$objUser.Put("displayName", $displayName);
		#$objUser.Put("givenName", $FirstName);
		#$objUser.Put("sn", $LastName);
		$objUser.SetInfo();
		if (($oADInfo.password -ne "") -and ($oADInfo.password -ne $null)){
			$objUser.SetPassword($oADInfo.password);
		}else{
			$objUser.SetPassword("S0me.P@$$w0rd4Y0u");
		}
		$objUser.psbase.InvokeSet("AccountDisabled", $False);
		$objUser.SetInfo();

	}

	function Check4OU{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$OUPath, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$RequiredDomain = ""
		)
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
		#$strTemp = "Check4OU(" + $strTemp.Trim() + ")";
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
		#Need to get Domains.  GetDomains() requires "PS-AD-Routines.ps1".
		if (!(Get-Command "GetDomains" -ErrorAction SilentlyContinue)){
			$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;
			if ((Test-Path ($ScriptDir + "\PS-AD-Routines.ps1"))){
				. ($ScriptDir + "\PS-AD-Routines.ps1")
			}
		}
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
					}else{
						$objReturn.Message = "The OU was found on multiple Domains.";
					}
				}else{
					#Remove the entry
					$arrDomains.Remove($arrDomains[$intX]);
				}
			}else{
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
			}else{
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
			}else{
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
		}else{
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

		$InitializeDefaultDrives=$False;
		if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};

		if ((($strDomain -ne "") -and ($strDomain -ne $null)) -or ($ComputerName.Contains("\"))){
			$arrDomains = @($strDomain);
			if ($ComputerName.Contains("\")){
				$arrDomains += $ComputerName.Split("\")[0];
				$ComputerName = $ComputerName.Split("\")[-1];
			}
		}else{
			#Need to get Domains.  GetDomains() requires "PS-AD-Routines.ps1".
			if (!(Get-Command "GetDomains" -ErrorAction SilentlyContinue)){
				$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;
				if ((Test-Path ($ScriptDir + "\PS-AD-Routines.ps1"))){
					. ($ScriptDir + "\PS-AD-Routines.ps1")
				}
			}
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
			}else{
				if ($strDomain -ne "nads"){
					$strProgress = "  Looking in " + $strDomain + " domain for " + $ComputerName + ".`r`n";

					if (($txbResults -ne "") -and ($txbResults -ne $null)){
						UpdateResults $strProgress $False;

						$strRIDMaster = GetOpsMaster2WorkOn $strDomain;
					}else{
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

	function FindUser{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$Username, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strDomain
		)
		#Checks All domains (gotten from the Network) for $Username, or just the ones provided.

		$InitializeDefaultDrives=$False;
		if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};

		if ((($strDomain -ne "") -and ($strDomain -ne $null)) -or ($Username.Contains("\"))){
			$arrDomains = @($strDomain);
			if ($Username.Contains("\")){
				$arrDomains += $Username.Split("\")[0];
				#$arrDomains = @($Username.Split("\")[0]);
				$Username = $Username.Split("\")[-1];
			}
		}else{
			#Need to get Domains.  GetDomains() requires "PS-AD-Routines.ps1".
			if (!(Get-Command "GetDomains" -ErrorAction SilentlyContinue)){
				$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;
				if ((Test-Path ($ScriptDir + "\PS-AD-Routines.ps1"))){
					. ($ScriptDir + "\PS-AD-Routines.ps1")
				}
			}
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
			}else{
				if ($strDomain -ne "nads"){
					$strProgress = "  Looking in " + $strDomain + " domain for " + $Username + ".`r`n";

					if (($txbResults -ne "") -and ($txbResults -ne $null)){
						UpdateResults $strProgress $False;

						$strRIDMaster = GetOpsMaster2WorkOn $strDomain;
					}else{
						#$strProgress;		#Outputs info for when running as background job.

						#$strRIDMaster = (Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
						$strRIDMaster = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain((New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $strDomain))).RidRoleOwner.Name;
					}
					if (($strRIDMaster -eq "") -or ($strRIDMaster -eq $null)){
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

		return $objUser;
	}

	function GetDomains{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bolFQDN = $False, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bolTrusted = $False
		)
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
		}else{
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
			#		}else{
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
					}else{
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

	function MoveUser{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$ADUser, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$DestOU
		)
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
		#$strTemp = "MoveUser(" + $strTemp.Trim() + ")";
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
					}else{
						#Don't know what Domain to use.
						$objReturn.Message = "The OU Path provided was found on more than one Domain," + "`r`n" + "and did not match the Domain specified in the provided OU Path.";
						$objReturn.Results = $False;
						$strDomain = "";
					}
				}else{
					#Don't know what Domain to use.
					$objReturn.Message = "The OU Path provided was found on more than one Domain.";
					$objReturn.Results = $False;
				}
			}else{
				#We can move the User.
				$strDomain = $objResults.Returns;
			}
		}else{
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
			}else{
				#Already have an AD Object
				$objUser = $ADUser;
			}

			if (($objUser -ne "") -and ($objUser -ne $null)){
				#Can do the actual move, once we pull all the parts together
				$strConfigFile = "\\nawesdnifs08.nadsuswe.nads.navy.mil\NMCIISF\NMCIISF-SDCP-MAC\MAC\Entr_SRM\Support Files\MiscSettings.txt";

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
					}else{
						$Error.Clear();
						if ($strDestRIDMaster -ne $strSrcRIDMaster){
							$(Try {Get-ADUser -Server $strSrcRIDMaster -Identity $objUser.DistinguishedName | Move-ADObject -Server $strSrcRIDMaster -TargetPath $DestOU -TargetServer $strDestRIDMaster} Catch {$null});
						}else{
							$(Try {Get-ADUser -Server $strSrcRIDMaster -Identity $objUser.DistinguishedName | Move-ADObject -Server $strSrcRIDMaster -TargetPath $DestOU} Catch {$null});
						}
						if ($Error){
							$objReturn.Message = $objReturn.Message + "Failed to move the AD object '" + $objUser.SamAccountName + "' to the Destination OU `r`n'$DestOU'.`r`n";
							$objReturn.Message = $objReturn.Message + $Error + "`r`n";
							$objReturn.Results = $False;
						}else{
							$objReturn.Results = $True;
							$strMessage = "Successfully moved the AD object '" + $objUser.SamAccountName + "' to the Destination OU.`r`n"
							if ($strDestRIDMaster -ne $strSrcRIDMaster){
								$objUser = Get-ADUser -Server $strDestRIDMaster -Identity $objUser.SamAccountName -Properties *;
								$strMessage = $strMessage + $strDestRIDMaster
							}else{
								$objUser = Get-ADUser -Server $strSrcRIDMaster -Identity $objUser.SamAccountName -Properties *;
								$strMessage = $strMessage + $strSrcRIDMaster
							}

							#$objReturn.Message = $objReturn.Message + $objUser.DistinguishedName + "`r`n";
							$strMessage = $strMessage + "   " + $objUser.DistinguishedName + "`r`n";
							$objReturn.Message = $objReturn.Message + $strMessage;
						}
					}
				}else{
					#Could not get the RID Master for the Destination Domain.
					$objReturn.Results = $False;
					$objReturn.Message = $objReturn.Message + "Failed to move the AD object '" + $objUser.SamAccountName + "' to the Destination OU `r`n'$DestOU'.`r`nCould not determine the RID/Ops Master DC from the network.`r`n";
					if ($strTempError){
						$objReturn.Message = $objReturn.Message + $strTempError + "`r`n";
					}
				}
			}else{
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

		$InitializeDefaultDrives=$False;
		if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};

		[System.Collections.ArrayList]$arrValues = @();
		#$arrValues = @();

		#Prep to interact with Term Serv Attributes.
		$arrTSProperties = "allowLogon","TerminalServicesHomeDirectory","TerminalServicesHomeDrive","TerminalServicesProfilePath";
		$strLDAP = "LDAP://$($strUserDN.SubString($strUserDN.IndexOf(",") + 1))";						#"LDAP://OU=USERS,OU=SDNI,OU=COMPACFLT,DC=nadsuswe,DC=nads,DC=navy,DC=mil"
		$strUserDN = $strUserDN.SubString(0, $strUserDN.IndexOf(","));									#"CN=redirect.test"
		$objOU = [ADSI]$strLDAP;
		$objADSIUser = $objOU.PSBase.get_children().find($strUserDN);
		#READ Term Serv Attributes
		foreach($strProperty in $arrTSProperties){
			if (($Property -ne "") -and ($Property -ne $null)){
				if ($Property -eq $($strProperty)){
					$arrValues += $($objADSIUser.PSBase.InvokeGet($strProperty)).ToString();
				}
			}else{
				#$strMessage = "$($strProperty) value: $($objADSIUser.PSBase.InvokeGet($strProperty))";
				#MsgBox $strMessage;
				$arrValues += "$($strProperty) = $($objADSIUser.PSBase.InvokeGet($strProperty))";
			}
		}

		return $arrValues;
	}

	function TSSet{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strUserDN, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$Attribute, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$Value
		)

		$InitializeDefaultDrives=$False;
		if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};

		#Prep to interact with Term Serv Attributes.
		#$arrTSProperties = "allowLogon","TerminalServicesHomeDirectory","TerminalServicesHomeDrive","TerminalServicesProfilePath";
		$strLDAP = "LDAP://$($strUserDN.SubString($strUserDN.IndexOf(",") + 1))";						#"LDAP://OU=USERS,OU=SDNI,OU=COMPACFLT,DC=nadsuswe,DC=nads,DC=navy,DC=mil"
		$strUserDN = $strUserDN.SubString(0, $strUserDN.IndexOf(","));									#"CN=redirect.test"
		$objOU = [ADSI]$strLDAP;
		$objADSIUser = $objOU.PSBase.get_children().find($strUserDN);
		#SET Term Serv Attributes
		#$objADSIUser.PSBase.invokeSet("allowLogon", $Value);
		#$objADSIUser.PSBase.invokeSet("TerminalServicesHomeDirectory", $Value);
		#$objADSIUser.PSBase.invokeSet("TerminalServicesProfilePath", $Value);
		#$objADSIUser.PSBase.invokeSet("TerminalServicesHomeDrive", $Value);
		$objADSIUser.PSBase.invokeSet($Attribute, $Value);
		$objADSIUser.setinfo();
	}

	function UpdateADField{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$ADUserDN, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$FieldName, 
			[ValidateNotNull()][Parameter(Mandatory=$True)]$NewValue, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$MultiVal, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$RIDMaster
		)
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
			}else{
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
		}else{
			if (($NewValue -eq "") -or ($NewValue -eq $null)){
				Set-ADUser -Identity $ADUserDN -Server $RIDMaster -Clear $FieldName;
			}else{
				$strCheckVal = [String](Get-ADUser -Identity $ADUserDN -Server $RIDMaster -Properties * | Select $FieldName);
				if ($strCheckVal.Contains("Microsoft.ActiveDirectory.Management.ADPropertyValueCollection")){
					#The field name supplied probably does not exist.
					$strCheckVal = "";
				}else{
					$strCheckVal = $strCheckVal.Replace("@{", "");
					$strCheckVal = $strCheckVal.Replace("}", "");
					$strCheckVal = $strCheckVal.Replace($FieldName + "=", "").Trim();
				}

				if (($strCheckVal.Trim() -eq "") -or ($strCheckVal -eq $null)){
					Set-ADUser -Identity $ADUserDN -Server $RIDMaster -Add @{$FieldName = $NewValue};
				}else{
					Set-ADUser -Identity $ADUserDN -Server $RIDMaster -Replace @{$FieldName = $NewValue};
				}
			}
		}
	}

	function SampleUsage{
		$DomainName = "NMCI-ISF";

		#To create a new blank user object:
		Add-Type -AssemblyName System.DirectoryServices.AccountManagement;
		#http://stackoverflow.com/questions/13688779/force-principalcontext-to-connect-to-a-specific-server
		$PrincipalContext = New-Object System.DirectoryServices.AccountManagement.PrincipalContext("DOMAIN", $DomainName_or_DCName);  #$DomainName_or_DCName = the Domain the new account will be create on;
		$NMCIUser = New-Object NMCI.AD.NMCIUserPrincipal($PrincipalContext);
		#$NMCIUser will now contain all NMCI AD Schema Attribs.

		#Write
		$NMCIUser.l = "City";
		$NMCIUser.City = "City";

		#Read
		foreach ($oPropInfo in $NMCIUser.GetType().GetProperties()){
			#$oPropInfo;
			<#
				MemberType    : Property
				Name          : PasswordNotRequired
				DeclaringType : System.DirectoryServices.AccountManagement.AuthenticablePrincipal
				ReflectedType : NMCI.AD.NMCIUserPrincipal
				MetadataToken : 385876020
				Module        : System.DirectoryServices.AccountManagement.dll
				PropertyType  : System.Boolean
				Attributes    : None
				CanRead       : True
				CanWrite      : True
				IsSpecialName : False

				MemberType    : Property
				Name          : Title
				DeclaringType : NMCI.AD.NMCIUserPrincipal
				ReflectedType : NMCI.AD.NMCIUserPrincipal
				MetadataToken : 385876028
				Module        : hnsughcj.dll
				PropertyType  : System.String
				Attributes    : None
				CanRead       : True
				CanWrite      : True
				IsSpecialName : False
			#>
			#$oPropInfo.Name;
			#$oPropInfo.GetValue($NMCIUser, $null);

			$strName = $oPropInfo.Name;
			$strValue = $oPropInfo.Attributes;
			#$strValue = $oPropInfo.GetValue($NMCIUser, $null);
			if (($strName -eq "") -or ($strName -eq $null)){
				$oPropInfo;
			}else{
				Write-Host $strName " = " $strValue;
			}
		}

	}


#All the code below here is the C Sharp User Class Chris did.
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
