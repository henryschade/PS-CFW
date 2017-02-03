###########################################
# Updated Date:	3 February 2017
# Purpose:		Provide a central location for all the PowerShell SM routines.
# Requirements: Core.ps1, and possibly MsgBox() from Forms.ps1.
##########################################

<# ---=== Change Log ===---
	#Changes for 3 February 2017
		#Initial file creation.  Routines still need lots of work.
#>


	#Make sure the routines in Core.ps1 are loaded.
	if (!(Get-Command "GetPathing" -ErrorAction SilentlyContinue)){
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


	function TestRoutine{

		. C:\Projects\PS-CFW\Core.ps1;

		#$InitializeDefaultDrives=$False;
		#if (!(Get-Module "ActiveDirectory")) {Import-Module ActiveDirectory;};

		#---=== Test 1 ===---
			
		#---=== Test 1 ===---


	}


	function CreateSMObj{
		Param(
			[Parameter(mandatory=$True)][String]$strData, 
			[Parameter(mandatory=$True)][String]$Param2, 
			[Parameter(mandatory=$True)][String]$Param3 
		)
		#Create a PS SM custom object
		#Returns: A custom SM object.
		#$strData = The SM data gotten in a Format of: *.xml, or *.txt.
		#$Param2 = ???
		#$Param3 = ???

		#Start of this idea...
		$objSM = New-Object PSObject -Property @{
			SMType = "unknown"
			TicketNum = "n/a"
		}

		#Logic here to parse data

		return $objSM;
	}

	function DoSMWeb{
		Param(
			[Parameter(mandatory=$True)][String]$strRequestNum, 
			[Parameter(mandatory=$True)][String]$strAction, 
			[Parameter(mandatory=$False)]$objIEWindow, 
			[Parameter(mandatory=$False)][String]$strWriteData
		)
		#Do SM Web scrapping
		#Returns: a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters passed in.
			#$objReturn.Results		= $True or $False.  Was the action successful.
			#$objReturn.Message		= "Success" or an error message.
			#$objReturn.Returns		= A Custom PS SM object.
		#$strRequestNum = The SM Request/Ticket # to get.
		#$strAction = "Push" or "Pull"
		#$objIEWindow = ??? The IE window with SM running in it????
		#$strWriteData = The data to be written back to the SM window.

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

		#Start working here.....
		#Updating $objReturn.Results, $objReturn.Message, and $objReturn.Returns as you go.

		return $objReturn;
	}

	function DoSMWSDL{
		Param(
			[Parameter(mandatory=$True)][String]$WSDLUN, 
			[Parameter(mandatory=$True)][String]$WSDLPWD, 
			[Parameter(mandatory=$True)][String]$XMLRequest, 
			[Parameter(mandatory=$True)][String]$XMLResponse, 
			[Parameter(mandatory=$True)][String]$WSDLName, 
			[Parameter(mandatory=$True)][String]$SOAPMethod, 
			[Parameter(mandatory=$False)][String]$WSDLServer = "SDJS06.nmci-isf.com", 
			[Parameter(mandatory=$False)][String]$WSDLPort = "13082", 
			[Parameter(mandatory=$False)]$bolDoSSL = $True
		)
		#Do SM WSDL Single requests
			#Based of "PS-SMSSL.ps1" started by Chris Henderson
		#Returns: a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters passed in.
			#$objReturn.Results		= $True or $False.  Was the action successful.
			#$objReturn.Message		= "Success" or an error message.
			#$objReturn.Returns		= The full path to the file.
		#$WSDLUN = The Username for the WSDL Credentials (Passed in so it is not stored in clear-text)
		#$WSDLPWD = The Password for the WSDL Credentials (Passed in so it is not stored in clear-text)
		#$XMLRequest = The SOAP Envelope passed in via String, or file path\name.
		#$XMLResponse = Location that you want the SOAP Response written out to as a flat file XML
		#$WSDLName = Name of the WSDL that is targeted for work (E.G.; SRMTools, OTD, OTDiBULK, InteractionInfo, etc)
		#$SOAPMethod = Method being called within targeted WSDL (E.G.; Create, Retrieve, RetrieveList, etc)
		#$WSDLServer = Server to target.
			#Prod -> "SDJS06.nmci-isf.com", "NFJS06.nmci-isf.com" 
			#Dev -> "nmcism7app.dadsuswe.dads.navy.mil" (10.10.181.53) 
			#Grat/UAT -> "sm7uatw.dadsusea.dads.navy.mil", "sm7uate.dadsusea.dads.navy.mil" (10.20.11.60)
		#$WSDLPort = Server Port to target. 
			#Prod -> 13082
			#Dev -> 13080
			#Grat/UAT -> 13081

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

		if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")){
			#"No Administrative rights, it will display UAC asking user for Admin rights"
			$arguments = " -ExecutionPolicy Bypass -Command & '" + $myinvocation.mycommand.definition + "' $WSDLUN $WSDLPWD '$XMLRequest' $XMLResponse $WSDLName $SOAPMethod $WSDLServer $WSDLPort"
			Start-Process "$psHome\powershell.exe" -Verb runAs -ArgumentList $arguments
			break
		}

		[System.Reflection.Assembly]::LoadWithPartialName("System.Security")
		[System.Reflection.Assembly]::LoadWithPartialName("System.Security.Cryptography")

		$Error.Clear()
		#$ServerPath = "https://{0}:{1}/SM/7/ws/{2}.wsdl" -f $WSDLServer,$WSDLPort,$WSDLName
		$ServerPath = "https://{0}:{1}/SM/7/{2}.wsdl" -f $WSDLServer,$WSDLPort,$WSDLName

		$webRequest = [System.Net.WebRequest]::Create($ServerPath)
		$httpRequest = [System.Net.HttpWebRequest]$webRequest
		$httpRequest.Method = "POST"
		$httpRequest.Headers.Add("SOAPAction: `"$SOAPMethod`"")
		$httpRequest.ContentType = "text/xml;charset=utf-8"
		$httpRequest.KeepAlive = $False			#False makes it logout when done.
		#$httpRequest.Timeout = 70000
		$httpRequest.Timeout = 25000
		$httpRequest.PreAuthenticate = $True

		#$bolDoSSL = $True;
		if ($bolDoSSL -eq $True){
			$store = New-Object System.Security.Cryptography.X509Certificates.X509Store([System.Security.Cryptography.X509Certificates.StoreName]::My, [System.Security.Cryptography.X509Certificates.StoreLocation]::LocalMachine)
			#$store = New-Object System.Security.Cryptography.X509Certificates.X509Store([System.Security.Cryptography.X509Certificates.StoreName]::My, [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
			$store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)

			if ($store.Certificates.Count -eq 1)
			{
				$clientCert = $store.Certificates[0]
				$httpRequest.ClientCertificates.Add($clientCert) | Out-Null    
			}
			elseif ($store.Certificates.Count -gt 1)
			{
				$certCollection = $store.Certificates
				$selectedCert = [System.Security.Cryptography.X509Certificates.X509Certificate2UI]::SelectFromCollection($certCollection, "Client Cert Selection", "Select a certificate from the following that is both valid (not expired) and matches your FQDN Client Host Name", [System.Security.Cryptography.X509Certificates.X509SelectionFlag]::SingleSelection)
				$httpRequest.ClientCertificates.Add($selectedCert[0]) | Out-Null  
			}
		}
		$httpRequest.Credentials = New-Object System.Net.NetworkCredential($WSDLUN,$WSDLPWD)
		 
		[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls
		[System.Net.ServicePointManager]::Expect100Continue = $True
		 
		$requestStream = $httpRequest.GetRequestStream()
		$streamWriter = New-Object System.IO.StreamWriter($requestStream, [System.Text.Encoding]::UTF8)

		if ((($XMLRequest.SubString(1, 1) -eq ":") -or ($XMLRequest.SubString(1, 1) -eq "\")) -and ([System.IO.File]::Exists($XMLRequest) -eq $True)){
			#passed in a file name
			#$XMLRequest = "C:\Users\Public\ITSS-Tools\Logs\SM-ASCII-SSL-Req.xml";
			[String]$XMLRequest = [System.IO.File]::ReadAllLines($XMLRequest)
		}
		 
		$soapRequest = New-Object System.Text.StringBuilder($XMLRequest)
		$streamWriter.Write($soapRequest.ToString())
		$streamWriter.Close()

		$webResponse = [System.Net.HttpWebResponse]$httpRequest.GetResponse()
		if (($webResponse -ne $null) -and ($webResponse -ne "")){
			$soapResponse = New-Object System.IO.StreamReader($webResponse.GetResponseStream())
			$responseXML = $soapResponse.ReadToEnd()
			$soapResponse.Close()
			if (($responseXML -eq $null) -or ($responseXML -eq "")){
				$responseXML = $Error
			}
		}
		else{
			#$Error += $Error[0].InvocationInfo
			#$responseXML = $Error
			$responseXML = $Error + "`r`n" + $Error[0].InvocationInfo;
		}
		$responseXML | Out-File -filepath $XMLResponse -Encoding Default

		$httpRequest.Abort();
		$webRequest.Abort();
	}

	function DoSMWSDL-Bulk{
		Param(
			[Parameter(mandatory=$True)][String]$WSDLUN,
			[Parameter(mandatory=$True)][String]$WSDLPWD,
			[Parameter(mandatory=$True)][String]$WSDLName,
			[Parameter(mandatory=$True)][String]$SOAPMethod,
			[Parameter(mandatory=$False)][String]$WSDLServer = "SDJS06.nmci-isf.com",
			[Parameter(mandatory=$False)][String]$WSDLPort = "13082",
			[Parameter(mandatory=$True)][String]$Folder
		)
		#Do SM WSDL Bulk requests
			#Based of "PS-SMSSL-Bulk.ps1" started by Chris Henderson
		#Returns: a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters passed in.
			#$objReturn.Results		= $True or $False.  Was the action successful.
			#$objReturn.Message		= "Success" or an error message.
			#$objReturn.Returns		= The full path to the files.
		#$WSDLUN = The Username for the WSDL Credentials (Passed in so it is not stored in clear-text)
		#$WSDLPWD = The Password for the WSDL Credentials (Passed in so it is not stored in clear-text)
		#$WSDLName = Name of the WSDL that is targeted for work (E.G.; OTD, OTDiBULK, InteractionInfo, etc)
		#$SOAPMethod = Method being called within targeted WSDL (E.G.; Create, Retrieve, RetrieveList, etc)
		#$WSDLServer = Server to target (i.e. "SDJS06.nmci-isf.com", "NFJS06.nmci-isf.com", "sm7svr", "sm7uatw", etc.)
			#Prod -> "SDJS06.nmci-isf.com", "NFJS06.nmci-isf.com" 
			#Dev -> "nmcism7app.dadsuswe.dads.navy.mil" (10.10.181.53) 
			#Grat/UAT -> "sm7uatw.dadsusea.dads.navy.mil", "sm7uate.dadsusea.dads.navy.mil" (10.20.11.60)
		#$WSDLPort = Server Port to target (By default 13082 if not provided/passed.)
			#Prod -> 13082
			#Dev -> 13080
			#Grat/UAT -> 13081
		#$Folder = The folder to read the "*.xml" request files from, and write the "*-Response.xml" files to.

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

		#Checking if running as Admin
		if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")){
			#"No Administrative rights, it will display UAC asking user for Admin rights"
			$arguments = " -ExecutionPolicy Bypass -Command & '" + $myinvocation.mycommand.definition + "' $WSDLUN $WSDLPWD $WSDLName $SOAPMethod $WSDLServer $WSDLPort $Folder"
			Start-Process "$psHome\powershell.exe" -Verb runAs -ArgumentList $arguments
			break
		}

		[System.Reflection.Assembly]::LoadWithPartialName("System.Security")
		[System.Reflection.Assembly]::LoadWithPartialName("System.Security.Cryptography")

		$Error.Clear()
		$ServerPath = "https://{0}:{1}/SM/7/{2}.wsdl" -f $WSDLServer,$WSDLPort,$WSDLName

		Get-ChildItem($Folder) | %{
			if ((!($_ -Like "*Response.xml")) -and ($_ -Like "*.xml")){
				Write-Host " "
				$webRequest = [System.Net.WebRequest]::Create($ServerPath)
				Write-Host "[DEBUG]::File $_"
				Write-Host "[DEBUG]::--Start Time-- $(Get-Date)"
				$XMLInPath = "{0}\{1}" -f $Folder,$_
				$httpRequest = [System.Net.HttpWebRequest]$webRequest
				$httpRequest.Method = "POST"
				$httpRequest.Headers.Add("SOAPAction: `"$SOAPMethod`"")
				$httpRequest.ContentType = "text/xml;charset=utf-8"
				$httpRequest.KeepAlive = $True			#False makes it logout when done.  Doing BULK so need True here.
				#$httpRequest.Timeout = 70000
				$httpRequest.Timeout = 7000
				$httpRequest.PreAuthenticate = $True
				Write-Host "[DEBUG]::Cert Store Opening..."
				$store = New-Object System.Security.Cryptography.X509Certificates.X509Store([System.Security.Cryptography.X509Certificates.StoreName]::My, [System.Security.Cryptography.X509Certificates.StoreLocation]::LocalMachine)
				$store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
				Write-Host "[DEBUG]::Cert Store Opening... [COMPLETE]" 
				Write-Host "[DEBUG]::Selecting Machine Cert..." 
				if ($store.Certificates.Count -eq 1)
				{
					$clientCert = $store.Certificates[0]
					$httpRequest.ClientCertificates.Add($clientCert) | Out-Null    
				}
				elseif ($store.Certificates.Count -gt 1)
				{
					$certCollection = $store.Certificates
					$selectedCert = [System.Security.Cryptography.X509Certificates.X509Certificate2UI]::SelectFromCollection($certCollection, "Client Cert Selection", "Select a certificate from the following that is both valid (not expired) and matches your FQDN Client Host Name", [System.Security.Cryptography.X509Certificates.X509SelectionFlag]::SingleSelection)
					$httpRequest.ClientCertificates.Add($selectedCert[0]) | Out-Null  
				}
				Write-Host "[DEBUG]::Selecting Machine Cert... [COMPLETE]" 
				$httpRequest.Credentials = New-Object System.Net.NetworkCredential($WSDLUN,$WSDLPWD)
				 
				[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls
				[System.Net.ServicePointManager]::Expect100Continue = $true
				
				Write-Host "[DEBUG]::Getting Request Stream..." 
				$requestStream = $httpRequest.GetRequestStream()
				Write-Host "[DEBUG]::Getting Request Stream... [COMPLETE]" 
				Write-Host "[DEBUG]::Creating Stream Writer..."
				$streamWriter = New-Object System.IO.StreamWriter($requestStream, [System.Text.Encoding]::UTF8)

				Write-Host "[DEBUG]::Creating Stream Writer... [COMPLETE]" 
				Write-Host "[DEBUG]::Reading XML Request file..."

				#Write-Host "         $XMLInPath"
				#$soapRequest = New-Object System.Text.StringBuilder($(Get-Content($XMLInPath)))		#always errors:   New-Object : Cannot find an overload for "StringBuilder" and the argument count: "80".

				$strFileName = $XMLInPath.Replace("\\", "\")
				Write-Host "         $strFileName"
				$XMLRequest = [System.IO.File]::ReadAllLines($strFileName)
				#If set the soap request to the next line, then it works
				#$XMLRequest = '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:pws="http://servicecenter.peregrine.com/PWS" xmlns:com="http://servicecenter.peregrine.com/PWS/Common"> <soapenv:Header/> <soapenv:Body> <CreateOTDiBulkRequest ignoreEmptyElements="True"> <model> <keys> </keys> <instance> <USNClaimant type="String">SPAWAR</USNClaimant> <OrderingBase type="String">SPOT</OrderingBase> <CLIN type="String">1026AW</CLIN> <OrderingUIC type="String">N00039</OrderingUIC> <BillingUIC type="String">N69666</BillingUIC> <ActivityUIC type="String">N00039</ActivityUIC> <TaskOrder type="String">Test</TaskOrder> <ModNo type="String">69</ModNo> <ServiceID type="String">12083786</ServiceID> <CCEmailAddresses type="String">steve.conklin@navy.mil; hoang.q.tran2.ctr@navy.mil</CCEmailAddresses> <RequestorBy type="String">CTR Hoang Tran</RequestorBy> <LastNameToUser type="String">Conklin</LastNameToUser> <NewLastName type="String">Conklin</NewLastName> <LastName type="String">Conklin</LastName> <FirstNameToUser type="String">Steve</FirstNameToUser> <NewFirstName type="String">Steve</NewFirstName> <FirstName type="String">Steve</FirstName> <MiddleInitialToUser type="String">G</MiddleInitialToUser> <NewMiddleInitial type="String">G</NewMiddleInitial> <MiddleInitial type="String">G</MiddleInitial> <NMCIUserIDToUser type="String">steve.conklin</NMCIUserIDToUser> <NMCIUserID type="String">steve.conklin</NMCIUserID> <NMCIID type="String">steve.conklin</NMCIID> <EmailAddressToUser type="String">steve.conklin@navy.mil</EmailAddressToUser> <RankRate type="String">CIV</RankRate> <MakeModelToAsset type="String">,</MakeModelToAsset> <RequestedDate type="String">11/20/2014 10:58</RequestedDate> <LoginRestrictionStartReal type="String">11/20/2014 10:58</LoginRestrictionStartReal> <TempDesiredDate type="String">11/20/2014 10:58</TempDesiredDate> <RequestID type="String">Test50RowIssue</RequestID> <BaseToLocation type="String">SPOT</BaseToLocation> <ShipToCity type="String">San Diego</ShipToCity> <CityToLocation type="String">San Diego</CityToLocation> <ShipToState type="String">CA</ShipToState> <StateToLocation type="String">CA</StateToLocation> <ShipToPostalCode type="String">92110</ShipToPostalCode> <ZipToLocation type="String">92110</ZipToLocation> <BuildingToLocation type="String">OT3</BuildingToLocation> <FileSharePath type="String">\\Nawespscfs41\C167\SPAWAR_SPOT_N00039</FileSharePath> <Category type="String">Hardware Services</Category> <RequestType type="String">Install - CLIN 23 Peripheral</RequestType> <HPSXCustomerID type="String">iBulk</HPSXCustomerID> <RowID type="String">3</RowID> <RowCount type="String">3</RowCount> <Company type="String">USN</Company> <Organization type="String">USN</Organization> <BuildoutURL type="String">https://some.url.to/the/buildout.xls</BuildoutURL> <CancelNETBuildoutURL type="String">https://net.ahf.nmci.navy.mil/DirectAccess/BuildoutCancel.aspx?buildoutId=12107174</CancelNETBuildoutURL> <ServiceFY type="String">2010</ServiceFY> <ServiceStartDate type="String">11/20/2014 10:58:27</ServiceStartDate> <ServiceEndDate type="String">11/20/2014 10:58:27</ServiceEndDate> <HPSXTransactionType type="String">INSERT</HPSXTransactionType> </instance> </model> </CreateOTDiBulkRequest> </soapenv:Body> </soapenv:Envelope>'
				#Write-Host "         $XMLRequest"
				#$soapRequest = New-Object System.Text.StringBuilder($XMLRequest)		#always errors:   New-Object : Cannot find an overload for "StringBuilder" and the argument count: "80".
				$soapRequest = [string]($XMLRequest)

				Write-Host "[DEBUG]::Reading XML Request file... [COMPLETE]"
				Write-Host "[DEBUG]::Adding XML Data -> Stream Writer..."
				$streamWriter.Write($soapRequest.ToString())
				Write-Host "[DEBUG]::Adding XML Data -> Stream Writer... [COMPLETE]"
				$streamWriter.Close()
				Write-Host "[DEBUG]::Stream Writer Closed"
				Write-Host "[DEBUG]::Get Response..." 
				$webResponse = [System.Net.HttpWebResponse]$httpRequest.GetResponse()
				if (($webResponse -ne $null) -and ($webResponse -ne "")){
					Write-Host "[DEBUG]::Get Response... [COMPLETE]"
					Write-Host "[DEBUG]::Get Response Stream..."
					$soapResponse = New-Object System.IO.StreamReader($webResponse.GetResponseStream())
					Write-Host "[DEBUG]::Get Response Stream... [COMPLETE]"
					$responseXML = $soapResponse.ReadToEnd()
					$soapResponse.Close()
					if (($responseXML -eq $null) -or ($responseXML -eq "")){
						$responseXML = $Error
					}
					$soapResponse.Dispose();
				}
				else{
					#$Error += $Error[0].InvocationInfo
					#$responseXML = $Error
					$responseXML = $Error + "`r`n" + $Error[0].InvocationInfo
				}
				Write-Host "[DEBUG]::Writting Response File..."
				$responseXML | Out-File -filepath $($XMLInPath -Replace '.xml', '-Response.xml') -Encoding Default
				Write-Host "[DEBUG]::Writting Response File... [COMPLETE]"
				Write-Host "[DEBUG]::Renamming XML Request File..."
				Rename-Item -Path $strFileName -NewName ($strFileName -Replace '.xml', '.done-xml')
				Write-Host "[DEBUG]::Renamming XML Request File... [COMPLETE]"
				Write-Host "[DEBUG]::--End Time-- $(Get-Date)"
				Write-Host " "
			}
		}

		$httpRequest.Abort();
		$webRequest.Abort();
	}

