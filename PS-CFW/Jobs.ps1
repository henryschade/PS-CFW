###########################################
# Updated Date:	28 August 2015
# Purpose:		Background Job and RunSpace Functions.
#				All this code is based on info from the following URLs:
#				#http://technet.microsoft.com/en-US/library/hh847783.aspx
#				#http://msdn.microsoft.com/en-us/library/dd878288(v=vs.85).aspx
#				#http://stackoverflow.com/questions/15520404/how-to-call-a-powershell-function-within-the-script-from-start-job
#				#http://stackoverflow.com/questions/7162090/how-do-i-start-a-job-of-a-function-i-just-defined
#				#http://stackoverflow.com/questions/8750813/powershell-start-job-scriptblock-cannot-recognize-the-function-defined-in-the-s
#
# Require -version 2.0
#				#http://www.nivot.org/post/2009/01/22/CTP3TheRunspaceFactoryAndPowerShellAccelerators
#				#http://msdn.microsoft.com/en-us/library/system.management.automation.runspaces.runspacefactory.createrunspacepool(v=vs.85).aspx
#				#http://thesurlyadmin.com/2013/02/11/multithreading-powershell-scripts/
##########################################

	function CheckJob{
		#Returns Job State.
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strJobName
		)

		#$objJob = Get-Job -Id $strJobName.ID;
		$objJob = Get-Job -Name $strJobName;
		if (($objJob -ne $null) -and ($objJob -ne "")){
			#A finished job has a state of "Complete" or "Failed". A job might also be "blocked" or "running".
			#if ($objJob.State -eq "Completed"){
			if ($objJob.State -ne "Running"){
				$strJobResults = Receive-Job -Job $objJob -Keep;
				if (($strJobResults -ne $null) -and ($strJobResults -ne "")){
					#return $strJobResults;
					return "Complete";
				}else{
					#Job finished running, good or bad, and is not returning any results.
					Remove-Job -Job $objJob;
					return "Failed";
				}
			}else{
				return "Running";
			}
		}else{
			return "Failed";
		}
	}

	function CheckRunSpaceJob{
		#Returns Job State.
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strJobName,
			[ValidateNotNull()][Parameter(Mandatory=$True)]$Jobs
		)
		$objJobs = $Jobs

		$bolFoundJob = $False;
		For ($intX = 0; $intX -lt $objJobs.Count; $intX++){
			if ($objJobs[$intX].Name -eq $strJobName){
				#Write-Host $objJobs[$intX].Powershell.EndInvoke($objJobs[$intX].Results);			#Complete the Async job. Which returns the Results of the job.
				$bolFoundJob = $True;
				break;
			}
		}

		if ($bolFoundJob -eq $False){
			return "Failed";
		}else{
			if (($objJobs[$intX].Powershell.InvocationStateInfo.State -eq $null) -or ($objJobs[$intX].Powershell.InvocationStateInfo.State -eq "")){
				return "Failed";
			}else{
				return $objJobs[$intX].Powershell.InvocationStateInfo.State;
			}
		}
	}

	function CleanRunSpace{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strJobName,
			[ValidateNotNull()][Parameter(Mandatory=$True)]$Jobs
		)
		$objJobs = $Jobs

		For ($intX = 0; $intX -lt $objJobs.Count; $intX++){
			if ($objJobs[$intX].Name -eq $strJobName){
				$objJobs[$intX].PowerShell.Dispose();
				$objJobs[$intX].Name = "";
				#$objJobs[$intX] = $null;		#CAN NOT DO THIS!!!!
			}
		}

		#$objJobs = $objJobs | ? {$_ -ne $null};
	}

	function CreateJob{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$JobName,
			[ValidateNotNull()][Parameter(Mandatory=$False)][ScriptBlock]$InitScript,
			[ValidateNotNull()][Parameter(Mandatory=$True)][ScriptBlock]$JobScript,
			$ArgumentList = $null
		)
		#Pass any arguments using the ArgumentList parameter

		#Start the Job
		#$objJob = Start-Job -Name $strJobName -ScriptBlock {Get-Process};
		#$objJob = Start-Job -InitializationScript $objJobCode -ScriptBlock {GetGroups}|
		if ($InitScript -ne $null){
			$objJob = Start-Job -Name $JobName -InitializationScript $InitScript -ScriptBlock $JobScript -ArgumentList $ArgumentList;
		}else{
			$objJob = Start-Job -Name $JobName -ScriptBlock $JobScript -ArgumentList $ArgumentList;
		}

		return $objJob;
	}

	function CreateRunSpace{
		#http://www.nivot.org/post/2009/01/22/CTP3TheRunspaceFactoryAndPowerShellAccelerators
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)]$intMax = 3
		)
		$intMin = 1;
		#$intMax = 3;

		#https://github.com/dfinke/powershell-for-developers/blob/master/chapter07/ShowUI/C%23/WPFJob.cs
		#$objInitSessState clone = sessionState.Clone();

		$objInitSessState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
		$objInitSessState.ImportPSModule("ActiveDirectory");
		#$objInitSessState.ImportPSModule({"ServerManager"})

		#https://gallery.technet.microsoft.com/scriptcenter/Gather-Generic-WMI-Data-474f788b
		#$Creds = Get-Credential;
		#$objInitSessState.Variables.Add(New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList 'Credential', $Creds);

		#objInitSessState.AddParameter("Credential", $Creds);

		$objPool = [RunSpaceFactory]::CreateRunspacePool($intMin, $intMax, $objInitSessState, $Host);		#$Host = local host.
		#$objPool = [RunSpaceFactory]::CreateRunspacePool($intMin, $intMax);
		#If need the RunSpacePool to be STA, instead of the default MTA.
		#http://blogs.msdn.com/b/powershell/archive/2008/05/22/wpf-powershell-part-1-hello-world-welcome-to-the-week-of-wpf.aspx
		#$objPool.ApartmentState, $objPool.ThreadOptions = “STA”, “ReuseThread”;
		$objPool.Open();

		#Write-Host "Created a Pool of $($objPool.GetAvailableRunspaces()) Runspaces.";

		return $objPool;
	}

	function CreateRunSpaceJob{
		#http://www.nivot.org/post/2009/01/22/CTP3TheRunspaceFactoryAndPowerShellAccelerators
		##http://thesurlyadmin.com/2013/02/11/multithreading-powershell-scripts/
		##http://blogs.technet.com/b/heyscriptingguy/archive/2013/09/29/weekend-scripter-max-out-powershell-in-a-little-bit-of-time-part-2.aspx

		#Returns PSObject
			#Name   	: NameSupplied
			#Powershell : System.Management.Automation.PowerShell
				#Commands            : System.Management.Automation.PSCommand
				#Streams             : System.Management.Automation.PSDataStreams
				#InstanceId          : aca20131-5c24-44d4-a272-aaba4e754469
				#InvocationStateInfo : System.Management.Automation.PSInvocationStateInfo
				#IsNested            : False
				#Runspace            :
				#RunspacePool        : System.Management.Automation.Runspaces.RunspacePool
			#Results    : System.Management.Automation.PowerShellAsyncResult
				#CompletedSynchronously : False
				#IsCompleted            : True
				#AsyncState             :
				#AsyncWaitHandle        : System.Threading.ManualResetEvent

		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)]$RSPool,
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$JobName,
			#[ValidateNotNull()][Parameter(Mandatory=$True)][ScriptBlock]$JobScript,
			[ValidateNotNull()][Parameter(Mandatory=$True)]$JobScript,
			[ValidateNotNull()][Parameter(Mandatory=$False)]$Arguments
		)

		#http://stackoverflow.com/questions/11000801/using-a-psdatacollection-with-begininvoke-in-powershell
		#$pipeline = [powershell]::Create().AddScript($JobScript).AddParameter("pauseTime", 5)
		$objPowershell = [powershell]::Create();											#Create a "powershell pipeline runner".
		if (@($JobScript).Count -gt 1){
			foreach ($strJob in $JobScript){
				$objPowershell.AddScript($strJob);
			}
		}else{
			$objPowershell.AddScript($JobScript);											#The PS command/script to run.
		}
		$objPowershell.RunspacePool = $RSPool;												#Assign to our pool of x runspaces to use.
		if (($Arguments -ne "") -and ($Arguments -ne $null)){								#Add any Arguments.
			if (($Arguments.GetType().FullName -eq "System.Object[]") -or ($Arguments.GetType().FullName -eq "System.Collections.ArrayList")){
				foreach ($strArg in $Arguments){
					$objPowershell.AddArgument($strArg);
				}
			}else{
				$objPowershell.AddArgument($Arguments);
			}
		}

		$strResults = $objPowershell.BeginInvoke();											#Start the job.
		#https://github.com/dfinke/powershell-for-developers/blob/master/chapter07/ShowUI/C%23/WPFJob.cs
		#if (psCmd.InvocationStateInfo.Reason != null)
		#Write-Host $objPowershell.Streams.Error.Count;
		if ($objPowershell.InvocationStateInfo.Reason -ne $null){
			$strResults = $objPowershell.InvocationStateInfo.Reason;
			Write-Host "Error " $strResults;
		}

		$objJobReturning = New-Object PSObject -Property @{
			Name = $JobName
			Powershell = $objPowershell
			Results = $strResults
			#Results = $objPowershell.BeginInvoke()
		}

		return $objJobReturning;
	}

	function WaitForJob{
		#Waits for a background job to finish then returns the Job Results.
		#If $objControl is provided it is updated w/ the progress and the results.
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)]$objJob, 
			[ValidateNotNull()][Parameter(Mandatory=$False)]$objControl, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$bolKeep
		)

		if (($bolKeep -eq "") -or ($bolKeep -eq $null)){
			$bolKeep = $False;
		}

		$strLastResults = "";
		$strJobResults = "";
		$objJobDetails = Get-Job -ID $objJob.ID;
		#$intX = 0;
		Do{
			if (($objControl -ne $null) -and ($objControl -ne "")){
				#$intX++;
				if ($bolKeep -eq $True){
					$strJobResults = Receive-Job -Job $objJobDetails -Keep;
				}else{
					$strJobResults = Receive-Job -Job $objJobDetails;
				}
				if (($strJobResults -ne $null) -and ($strJobResults -ne "")){
					##if ($strJobResults -ne $strLastResults){
					#if (!($objControl.Text.Contains($strLastResults))){
						$objControl.Text = $objControl.Text + $strJobResults + "`r`n";
					#}
					$strLastResults = $strJobResults;
					#$strJobResults = "";
				}
				[System.Windows.Forms.Application]::DoEvents();
				#$objControl.Refresh;
			}
		}Until (($objJobDetails.State -eq "Completed") -or ($objJobDetails.State -eq "Complete"))
		#}Until (($objJobDetails.State -eq "Completed") -or ($intX -gt 10000))
		if (($strJobResults -eq "") -or ($strJobResults -eq $null)){
			$strJobResults = Receive-Job -Job $objJobDetails -Keep;
			if (($strJobResults -eq "") -or ($strJobResults -eq $null)){
				#Delete the job.
				Remove-Job -Job $objJobDetails;
			}
		}

		if ((($strJobResults -eq "") -or ($strJobResults -eq $null)) -and ($strLastResults -ne "")){
			$strJobResults = $strLastResults;
		}

		return $strJobResults;
	}

	function WaitForRunSpaceJob{
		#Waits for a RunSpace background job to finish then returns the Job Results.
		#If $objControl is provided it is updated w/ the progress and the results.
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strJobName,
			[ValidateNotNull()][Parameter(Mandatory=$True)]$Jobs,
			[ValidateNotNull()][Parameter(Mandatory=$False)]$objControl, 
			[ValidateNotNull()][Parameter(Mandatory=$False)]$bIgnoreErr = $False
		)
		$objJobs = $Jobs;

		$bolFoundJob = $False;
		For ($intX = 0; $intX -lt $objJobs.Count; $intX++){
			if ($objJobs[$intX].Name -eq $strJobName){
				$bolFoundJob = $True;
				break;
			}
		}

		if ($bolFoundJob -eq $False){
			return "Failed";
		}else{
			$strJobResults = $null;
			Do{
				if (($objControl -ne $null) -and ($objControl -ne "")){
					$strJobResults = $null;
					[String]$strJobResults = [String]$objJobs[$intX].Powershell.EndInvoke($objJobs[$intX].Results);				#Complete the Async job. Which returns the Results of the job.
					if (($strJobResults -ne $null) -and ($strJobResults -ne "")){
						$objControl.Text = $objControl.Text + $strJobResults + "`r`n";
					}
					#[System.Windows.Forms.Application]::DoEvents();
					$objControl.Refresh;
				}
			}Until (($objJobs[$intX].Results.IsCompleted))

			#Write-Host "`r`nResults1: " $strJobResults;			#Blank
			#Write-Host "`r`nResults2: " $objJobs[$intX].Powershell.EndInvoke($objJobs[$intX].Results);		#Results

			if (($strJobResults -eq "") -or ($strJobResults -eq $null)){
				#[String]$strJobResults = [String]$objJobs[$intX].Powershell.EndInvoke($objJobs[$intX].Results);				#Complete the Async job. Which returns the Results of the job.
				#if (($strJobResults -eq "") -or ($strJobResults -eq $null)){
					$strJobResults = $objJobs[$intX].Powershell.EndInvoke($objJobs[$intX].Results);				#Complete the Async job. Which returns the Results of the job.
				#}
				#Write-Host ($strJobResults.GetType());			#System.Management.Automation.PSDataCollection`1[System.Management.Automation.PSObject]
			}

			if (($objControl -ne $null) -and ($objControl -ne "")){
				if ((!($objControl.Text.Contains($strJobResults))) -and (($strJobResults -ne "") -and ($strJobResults -ne $null))){
					$objControl.Text = $objControl.Text + $strJobResults + "`r`n";
				}
			}

			#if ((($strJobResults -eq $null) -or ($strJobResults -eq "")) -or (($objJobs[$intX].Powershell.Streams.Error.Count -gt 0) -and (($objJobs[$intX].Powershell.Streams.Error -ne $null) -and ($objJobs[$intX].Powershell.Streams.Error -ne "")))){
			if (($strJobResults -eq $null) -or ($strJobResults -eq "") -or (($objJobs[$intX].Powershell.Streams.Error.Count -gt 0) -and ($bIgnoreErr -ne $True))){
				#$strJobResults = "Job finished running, good or bad, and is not returning any results."
				#Write-Host $strJobResults;
				#Write-Host "Error??";

				$strIsComp = $objJobs[$intX].Results.IsCompleted;
				$strState = $objJobs[$intX].Powershell.InvocationStateInfo.State;
				$strJobResults = $objJobs[$intX].Powershell.InvocationStateInfo.Reason;									#The Reason that Powershell failed.
				if (($strJobResults -eq "") -or ($strJobResults -eq $null)){
					#Write-Host $objJobs[$intX].Powershell.Streams.Error.Count;
					#Write-Host $objJobs[$intX].Powershell.Streams.Error[0];
					#Write-Host $objJobs[$intX].Powershell.Streams.Error[($objJobs[$intX].Powershell.Streams.Error.Count - 1)];

					if (($objJobs[$intX].Powershell.Streams.Error.Count -gt 0) -and (($objJobs[$intX].Powershell.Streams.Error -ne $null) -and ($objJobs[$intX].Powershell.Streams.Error -ne ""))){
						$strJobResults = "Error:`r`n" + $objJobs[$intX].Powershell.Streams.Error;
					}
				}
				#$strJobResults = "Job IsComplete:$strIsComp; Job State:$strState; Job Reason:$strJobResults; Job Results:`r`n"
				$strJobResults = $strJobResults + $objJobs[$intX].Powershell.EndInvoke($objJobs[$intX].Results);		#Complete the Async job. Which returns the Results of the job.

				##$objJobs[$intX].PowerShell.Dispose();
				#CleanRunSpace $strJobName $objJobs;
			}else{
				#$strTestName = $objJobs[$intX].Name;
				#Write-Host "Have some results for $strTestName" -NoNewLine;
				#Write-Host ($strJobResults.GetType());
				#Write-Host $strJobResults;
				#if (($strJobResults.GetType().Name -ne "String") -and ($strJobResults.GetType().Name -ne "System.String")){
				#	[String]$strJobResults = $strJobResults;
				#	Write-Host $strJobResults.GetType();
				#}
			}

			#MsgBox $strJobResults;
			return $strJobResults;
		}
	}


	#Sample Job calls
	function SampleJobCall1{
		#Pass any arguments using the ArgumentList parameter
		$objJob1 = CreateJob -JobName "FindUserDomain" -JobScript {
			Param(
				[ValidateNotNull()][Parameter(Mandatory=$True)][String]$Username
			)
			$strSrcDomain = "";
			Import-Module ActiveDirectory -ErrorAction SilentlyContinue;

			#$arrDomains = @("nadsusea", "nadsuswe", "pads", "nmci-isf");
			#Get Domain List
			$objForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest();
			$DomainList = @($objForest.Domains | Select-Object Name);
			#$Domains = $DomainList | foreach {$_.Name};
			$arrDomains = @($DomainList | foreach {($_.Name).split(".")[0]});
			#Write-Host $arrDomains;
			if ($arrDomains -eq ""){
				$arrDomains = @("nadsusea", "nadsuswe", "pads", "nmci-isf");
			}

			foreach ($strDomain in $arrDomains){
				if (($strDomain -eq $null) -or ($strDomain -eq "")){
					break;
				}

				$strProgress = "Looking in $strDomain domain.";
				$strProgress;
				$strRIDMaster = (Get-ADDomain $strDomain -ErrorAction SilentlyContinue).RIDMaster;
				#$objUser = Get-ADUser -Server $strRIDMaster -Identity $UserName;
				#$objComp = Get-ADComputer -Server $strRIDMaster -Identity "MachineName";
				$objUser = $(Try {Get-ADUser -Server $strRIDMaster -Identity $UserName} Catch {$null});
				If (($objUser.DistinguishedName -ne "") -and ($objUser.DistinguishedName -ne $null)){
					$strSrcDomain = $strDomain;
					#Write-Host "Found $UserName in $strSrcDomain domain.";
					#$objTxbResults.TEXT = $objTxbResults.TEXT + "  -- Found " + $UserName + " in " + $strDomain + " domain." + "`r`n";
					break;
				}
			}
			
			return $strSrcDomain;} -ArgumentList $UserName;
	}
	function SampleJobCall2{
		#http://stackoverflow.com/questions/15520404/how-to-call-a-powershell-function-within-the-script-from-start-job
		$objJobCode = [scriptblock]::create($function:FindUserDomain);
		#Pass any arguments using the ArgumentList parameter
		$objJob1 = CreateJob -JobName "FindUserDomain" -JobScript $objJobCode -ArgumentList $UserName;

		#or
		#http://stackoverflow.com/questions/7162090/how-do-i-start-a-job-of-a-function-i-just-defined
		#http://stackoverflow.com/questions/8750813/powershell-start-job-scriptblock-cannot-recognize-the-function-defined-in-the-s
		$objJobCode = [scriptblock]::create("function GetGroups {" + $function:GetGroups + "}");
		#Pass any arguments using the ArgumentList parameter
		$objJob1 = CreateJob -JobName "GetMemberships" -InitScript $objJobCode -JobScript {param($Name) GetGroups $Name} -ArgumentList @($env:UserName, @());
	}

	function SampleRunSpaceCalls{
		$objPool = CreateRunSpace 3;
		$objJobs = @();								#Collection of the Jobs that get run in the RunSpacePool.

		$objJobCode = [scriptblock]::create("Get-Command");
		#or
		#$objJobCode = [scriptblock]::create({Get-Command});
		$objJobs += CreateRunSpaceJob -RSPool $objPool -JobName "GetCmds" -JobScript $objJobCode;
		#Write-Host $objJobs[($objJobs.Count) - 1].Name;
		#CheckRunSpaceJob "GetCmds" $objJobs;
		$strResults = WaitForRunSpaceJob "GetCmds" $objJobs;
		Write-Host $strResults;
		CleanRunSpace "GetCmds" $objJobs;




		$objJobCode = [scriptblock]::create({Import-Module ActiveDirectory; (Get-ADDomain "nmci-isf").RIDMaster;});
		$objJobs += CreateRunSpaceJob -RSPool $objPool -JobName "RIDMaster" -JobScript $objJobCode;
		$strResults = CheckRunSpaceJob "RIDMaster" $objJobs;
		if (($strResults) -eq "Completed"){
			$strResults = WaitForRunSpaceJob "RIDMaster" $objJobs;
		}
		Write-Host $strResults;
		CleanRunSpace "RIDMaster" $objJobs;




		#If I create an Exchange Session in the Pool, then the Exchange cmdlets are only available in the Pool.
		$objJobCode = [scriptblock]::create($function:GetPSCmds);		#create an Exchange Session, and returns all imported commands.
		$objJobs += CreateRunSpaceJob -RSPool $objPool -JobName "GetPSCmds" -JobScript $objJobCode;
		$strResults = WaitForRunSpaceJob $objJobs[($objJobs.Count) - 1].Name $objJobs;
		#Write-Host $strResults;
		#CleanRunSpace "GetPSCmds" $objJobs;

		#$objJobCode = [scriptblock]::create("Get-MailboxStatistics -Identity ""henry.schade"" | Select DisplayName, TotalItemSize, ItemCount, TotalDeletedItemSize, StorageLimitStatus, ServerName, DatabaseName, MailboxGUID;");
		$objJobCode = [scriptblock]::create({param($Username, $strDC); Get-MailboxStatistics -Identity $Username -DomainController $strDC | Select DisplayName, TotalItemSize, ItemCount, TotalDeletedItemSize, StorageLimitStatus, ServerName, DatabaseName, MailboxGUID;});
		$arrArgs = @("henry.schade", "SomeDC");
		$objJobs += CreateRunSpaceJob $objPool "MailBoxSize" $objJobCode -Arguments $arrArgs;
		#$objJobs += CreateRunSpaceJob $objPool "MailBoxSize" $objJobCode -Arguments @("henry.schade", "SomeDC");
		$objMailBoxInfo = WaitForRunSpaceJob "MailBoxSize" $objJobs;
		Write-Host $objMailBoxInfo.DisplayName
		Write-Host $objMailBoxInfo.TotalItemSize
		Write-Host $objMailBoxInfo.ItemCount
		Write-Host $objMailBoxInfo.TotalDeletedItemSize
		CleanRunSpace "MailBoxSize" $objJobs;




		$objJobCode1 = [scriptblock]::create("function GetGroups {" + $function:GetGroups + "}");
		$objJobCode2 = [scriptblock]::create({param($Name) GetGroups $Name;});
		$objJobs += CreateRunSpaceJob -RSPool $objPool -JobName "GetUserMemberships" -JobScript @($objJobCode1, $objJobCode2) -Arguments @($env:UserName);
		Write-Host "Name: " $objJobs[($objJobs.Count) - 1].Name;
		$strResults = WaitForRunSpaceJob "GetUserMemberships" $objJobs;
		Write-Host $strResults;
		CleanRunSpace "GetUserMemberships" $objJobs;
	}
