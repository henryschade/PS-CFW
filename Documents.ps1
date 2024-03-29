###########################################
# Updated Date:	11 May 2023
# Purpose:		Code to manipulate Documents.
# Requirements: None
##########################################

<# ---=== Change Log ===---
	#Changes for 28 June 2016:
		#Added Change Log.
	#Changes for 21 November 2016
		#Updated ParseLogFile() documentation, and added a PS progress bar.
		#Added Close-OpenFile() routine.
	#Changes for 8 December 2016
		#Add "#Returns: " to functions, for routine documentation.
	#Changes for 11 May 2023
		#Add CSV functions.
#>



	function SampleExcelReadCOM{
		#Option 3, is the desired option at this time.

		#. C:\Projects\PS-CFW\Documents.ps1
		$strFilePath = "C:\Projects\PS-Scripts\Testing\CIVMAR Bulk MAC.xls";
		$WorkSheetName = "Create User Account";
		$bCreateIfNotExist = $False;
		$bVisible = $False;
		$bReadOnly = $False;
		#Google search "powershell Excel.Application iterate columns"

		#Opening the file.
		($objExcel, $objWorkBook) = ExcelCreateOpenFile $strFilePath $bCreateIfNotExist $bVisible $bReadOnly;

		#Check if got the workbook
		if ($objWorkBook){
			#Get an Array of all available sheets:
			$arrSheets = ExcelGetWorksheet $objWorkBook;

			#Get a WorkSheet object.
			#$objSheets = ExcelGetWorksheet $objWorkBook $arrSheets[1];
			#$objWorkSheet = $arrSheets[1];
			#or one of the following:
			#$objSheets = ExcelGetWorksheet $objWorkBook "SheetName";
			#$objWorkSheet = ExcelGetWorksheet -Workbook $objWorkBook -SheetName "East";
			$objWorkSheet = ExcelGetWorksheet $objWorkBook $WorkSheetName;

			#---=== Option 1 ===---
			#Get an object of the Cells
			#$objCells = $objWorkSheet.Cells;

			#Read from the worksheet.
			#Write-Host $objCells.Item($intRow, $intCol).Value();
			#Write-Host $objCells.Item(1, 1).Value();					#A1
			#Write-Host $objCells.Item(1, 2).Value();					#B1
			#Write-Host $objCells.Item(2, 1).Value();					#A2
			#Write-Host $objCells.Item(2, 2).Value();					#B2

			##Clean up / Release Cells object
			#[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objCells) | Out-Null;
			#$objCells = $null;
			#---=== Option 1 ===---

			#---=== Option 2 ===---
			#or don't bother with the Cells object and just do the following
			#Write-Host $objWorkSheet.Range("A2").Text;					#A2
			#---=== Option 2 ===---

			#---=== Option 3 ===---
			#Use ExcelGetData().  Returns a DataTable object.
			#$objData = ExcelGetData $objWorkBook $objWorkSheet;
			#---=== Option 3 ===---

			#Clean up / Release WorkSheet object
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objWorkSheet) | Out-Null;
			$objWorkSheet = $null;

			#Close the workbook.
			#Turn off Error messages.
			#$objExcel.DisplayAlerts = $False;
			$objWorkBook.Close();
				#Or could close Workbook with false (Don´t save changes)
				#$objWorkBook.Close($False);
			#Turn Error messages back on.
			#$objExcel.DisplayAlerts = $True;
			#Clean up / Release WorkBook object
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objWorkBook) | Out-Null;
		}

		#Quit/Close Excel.
		$objExcel.Quit();
			#Excel still shows in TaskManager...
			#When close Powershell then Excel will close too (this is a .NET issue using COM objects).
		#Clean up / Release Excel object
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel) | Out-Null;
		$objExcel = $null;
	}

	function SampleExcelUsageCOM{
		# The Excel Code is from following URL, but modified for my use.
		# http://mypowershell.webnode.sk/news/create-or-open-excel-file-in-powershell/

		#$ExcelFilePath = "c:\temp\MyExcelFile2.xlsx"
		#data 'raz, 'dva', 'tri' will be inserted to excel sheet 'Data'
		#ExcelWriteData -InputData @{"raz", "dva", "tri") -ExcelFilePath $ExcelFilePath



		#What file to open/create
		$strFilePath = "\\nawesdnifs101v.nadsuswe.nads.navy.mil\NMCIISF02$\ITSS-Tools\SRM\SupportFiles\BackUpLocation\USN_Server_Farms-Testing.xls";
		$strFilePath = "C:\Projects\PS-Scripts\Testing\CIVMAR Bulk MAC.xls";
		$WorkSheetName = "Create User Account";

		#Opening the file, and/or create it, and visible or not.
		($objExcel, $objWorkBook) = ExcelCreateOpenFile -ExcelFilePath $strFilePath $False $False;

		#Check if got the workbook
		if ($objWorkBook){
			#Got a workbook, so now can get the desired sheet.
			#$objWorkSheet = ExcelGetWorksheet -Workbook $objWorkBook -SheetName "East";
			$objWorkSheet = ExcelGetWorksheet -Workbook $objWorkBook -SheetName $WorkSheetName;

			#Read the worksheet, or write to the worksheet.
			#$objCells = $objWorkSheet.Cells;

			#Read from the worksheet.
			#Write-Host $objCells.Item(1, 1).Value();					#A1
			#Write-Host $objCells.Item(1, 2).Value();					#B1
			#Write-Host $objCells.Item(2, 1).Value();					#A2
			#Write-Host $objCells.Item(2, 2).Value();					#B2
			#Write-Host $objCells.Item($intRow, $intCol).Value();
			#or:
			#Write-Host $objWorkSheet.Range("A2").Text;					#A2

			#Write to the worksheet.
			#$objCells.Item(1, 1) = "A1";
			#$objCells.Item(1, 2) = "B1";
			#$objCells.Item(2, 1) = "A2";
			#$objCells.Item(2, 2) = "B2";
			#or:
			#$objCells.Item(1, 1).Value() = "A1";
			#$objCells.Item(1, 2).Value() = "B1";
			#$objCells.Item(2, 1).Value() = "A2";
			#$objCells.Item(2, 2).Value() = "B2";

			#Turn off Error message for replacing existing file when saving it.
			#$objExcel.DisplayAlerts = $False;
			#$objWorkBook.SaveAs($strFilePath);																		#appears to default to Excel 2007/2010 format
			#$objWorkBook.SaveAs($strFilePath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlExcel12);			#Excel 2007/2010
			#$objWorkBook.SaveAs($strFilePath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlExcel8);			#Excel 95/97/2003
			#Turn Error message back on.
			#$objExcel.DisplayAlerts = $True;

			#Clean up / Release Cells object
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objCells) | Out-Null;

			#Clean up / Release WorkSheet object
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objWorkSheet) | Out-Null;

			#Close the workbook.
			$objWorkBook.Close();
			#Clean up / Release WorkBook object
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objWorkBook) | Out-Null;
		}
		#Quit/Close Excel.
		$objExcel.Quit();
		#Clean up / Release Excel object
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel) | Out-Null;
	}

	function SampleExcelUsageXML1{
		#Updated the Code from Trev, that treats an Excel doc like an xml doc to pull data.

		$strDocPath = "C:\Projects\PS-Scripts\Testing\CIVMAR Bulk MAC.xls";
		$WorkSheetName = "Create User Account";
		$Query = "";
		$ListValues = "";
		#Create the scriptblock to run in a job
		$JobCode = [scriptblock]::create($function:ExcelXML);


		#Create
		$strDocPath = "C:\Projects\PS-Scripts\Testing\Test.xls";
		$WorkSheetName = "TestSheet";
		$ListValues = "Col1, Col2, Col3";		#Not sure what this should be.
		#Create the scriptblock to run in a job
			#$JobCode = [scriptblock]::create($function:ExcelXML);		#$strDocPath $WorkSheetName $ListValues;
		#Run the code in a 32bit job, since the provider is 32bit only
		$job = Start-Job $JobCode -RunAs32 -ArgumentList $strDocPath, $Query, $WorkSheetName, $ListValues
		$job | Wait-Job | Receive-Job
		Remove-Job $job


		#GetSheets  (Gets all sheets)
		#Create the scriptblock to run in a job
			#$JobCode = [scriptblock]::create($function:ExcelXML);		#$strDocPath $Query;
		#Run the code in a 32bit job, since the provider is 32bit only
		$job = Start-Job $JobCode -RunAs32 -ArgumentList $strDocPath
		$job | Wait-Job | Receive-Job
		Remove-Job $job


		#GetSheetData
		#Create the scriptblock to run in a job
			#$JobCode = [scriptblock]::create($function:ExcelXML);		# $strDocPath $Query;
		$WorksheetName = "Sheet1";
		$Query = 'SELECT * FROM [Sheet1$]';
		#$Query = 'SELECT * FROM [' + $WorksheetName + '$]';
		#Run the code in a 32bit job, since the provider is 32bit only
		$job = Start-Job $JobCode -RunAs32 -ArgumentList $strDocPath, $Query
		$job | Wait-Job | Receive-Job
		Remove-Job $job
	}

	function SampleExcelUsageXML2{
		$strDocPath = "C:\Projects\PS-Scripts\Testing\CIVMAR Bulk MAC.xls";
		$WorkSheetName = "Create User Account";
		$Query = "";
		$ListValues = "";

		#If I can figure out the 64 Bit PS with 32 Bit Office DLL's issue, this will work.

		#Create
		$strDocPath = "C:\Projects\PS-Scripts\Testing\Test.xls";
		$WorkSheetName = "TestSheet";
		$ListValues = "Col1, Col2, Col3";		#Not sure what this should be.
		$objData = ExcelXML $strDocPath $Query $WorkSheetName $ListValues;


		#GetSheets
		$objData = ExcelXML $strDocPath;


		#GetSheetData
		$Query = "SELECT * FROM [" + $WorkSheetName + '$]';
		#or 		$Query = 'SELECT * FROM [Create User Account$]';
		$objData = ExcelXML $strDocPath $Query;
		#or
		#$Query = "";
		#$objData = ExcelXML $strDocPath $Query $WorkSheetName;
	}

	function SampleEncodeDecode{
		#From a PowerShell window run one of the following commands:
		. "C:\Projects\PS-CFW\Documents.ps1"

		Encode "C:\Settings.txt" "C:\EncSet.txt"
		Encode "C:\Settings.txt" "Display"
		#"";

		DeCode "....JUeXBlID0gbXNzcWwNCnN0ckRCU2VydmVyID0gT......" "C:\Settings.txt"
		DeCode "....JUeXBlID0gbXNzcWwNCnN0ckRCU2VydmVyID0gT......" "Display"

		$strEncode = "c3RyREJUeXBlID0gbXNzcWwNCnN0ckRCU2VydmVyID0gTkFFQU5SRktTUTc1VkFcU1E3NVZBSU5TVDAxDQpzdHJEQk5hbWUgPSBTaXRlQ29kZXMNCnN0ckRCTG9naW5SID0gS0J1c2VyDQpzdHJEQlBhc3NSID0ga2M1JHNxMDI=";
		DeCode $strEncode "C:\Users\henry.schade\Desktop\SQL.txt"
		DeCode $strEncode "Display"
		Encode "C:\Users\henry.schade\Desktop\SQL.txt" "C:\Users\henry.schade\Desktop\EncSet.txt"


		$strDir = "C:\Program Files (x86)\Microsoft\Exchange\Web Services\1.2";
		$arrFiles = Get-ChildItem $strDir
		foreach ($strFile in $arrFiles){
			Encode ($strFile.FullName) (($strFile.FullName).Replace($strFile.Extension, ($strFile.Extension + ".enc.txt")))
		}
	}


	function Close-OpenFile{
		[CmdletBinding()]
		Param (
			[Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position=0)][String[]]$filesToClose
		)
		#If a file is locked open, this will close it.
			#http://serverfault.com/questions/718875/close-locked-file-in-windows-share-using-powershell-and-openfiles
		#Returns: 
		#$filesToClose = The file, with path, to close, or an array list of the files, with paths.

		Begin {
			$netFile = net file
			if($netFile.length -lt 7) { Throw "No Files are Open" }
			$netFile = $netFile[4..($netFile.length-3)]
			$netFile = $netFile | ForEach-Object {
				$column = $_ -split "\s+", 4
				New-Object -Type PSObject -Property @{
					ID = $column[0]
					FilePath = $column[1]
					UserName = $column[2]
					Locks = $column[3]
				}
			}
			$count = 0
		} Process {
			ForEach ($file in $filesToClose) {
				ForEach ($openFile in $netFile) {
					if($openFile.FilePath -eq $file) {
						$count++
						net file $openfile.ID /close > $null
					}
				}
			}
		} End { Write-Output "Closed $count Files" }
	}

	function Csv-Merge {
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strBase64String, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strOutPut
		);
		#Description: Merge 2 CSV files into one.
		#Returns: a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters passed in.
			#$objReturn.Results		= $True or $False.
			#$objReturn.Message		= "Success" or the error message.
			#$objReturn.Returns		= 
		#Inputs:
			#$strBase64String = String to be Decoded
			#$strOutPut = How to output.  "Display" or [Output file with full path]

	}

	function Csv-BuildCustomObject {
		Param(
			[Parameter(Mandatory=$True)]$custObject,
			[ValidateNotNull()][Parameter(Mandatory=$True)]$rowData,
			[ValidateNotNull()][Parameter(Mandatory=$True)]$headerColumns1,
			[ValidateNotNull()][Parameter(Mandatory=$True)]$headerColumns2
		)

		if ($null -eq $custObject) {
			$custObject = [PSCustomObject]@{
				$headerColumns1[0] = $rowData.($headerColumns2[0]);
			};
		}

		#Add all the column data (as blanks) from file 1
		for ($x = 1; $x -lt $headerColumns1.Count; $x++) {
			$custObject | Add-Member NoteProperty -Name $headerColumns1[$x] -Value "" -Force;
		}
		#Add all the column data from file 2
		for ($x = 1; $x -lt $headerColumns2.Count; $x++) {
			$custObject | Add-Member NoteProperty -Name $headerColumns2[$x] -Value $rowData.($headerColumns2[$x]) -Force;
		}

		return $custObject;
	}

	function Csv-GetHeaderColumns {
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)]$strFilePath,
			[ValidateNotNull()][Parameter(Mandatory=$True)]$strDelimiter
		)

		[System.IO.StreamReader]$reader = [System.IO.File]::Open($strFilePath, [System.IO.FileMode]::Open);
		$line = $reader.ReadLine();
		$reader.Close();
	
		return $line -split $strDelimiter;
	}

	function Csv-MergeTwoFilesAllData {
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)]$strFile1Path,
			[ValidateNotNull()][Parameter(Mandatory=$True)]$strFile2Path
		)
		#Description: Merge entries from file 2 that match file 1, add non matches as a new row.
		#Returns: a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters passed in.
			#$objReturn.Results		= $True or $False.
			#$objReturn.Message		= "Success" or the error message.
			#$objReturn.Returns		= The merged file results.
		#Inputs:
			#$strFile1Path = First file to merge.
			#$strFile2Path = Second file to merge.

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

		# Verify the 2 files exist.
		if (![System.IO.File]::Exists($strFile1Path)) {
			$objReturn.Message = "Error, file '$strFile1Path' could not be found.";
			return $objReturn;
		}
		if (![System.IO.File]::Exists($strFile2Path)) {
			$objReturn.Message = "Error, file '$strFile2Path' could not be found.";
			return $objReturn;
		}

		$mergeResult = @{};
		$itemCount = 0;
		#$lineCount = (Get-Content -Path $strFile1Path | Measure-Object -Line).Lines;
		$lineCount = 0;
		Get-Content -Path $strFile2Path -ReadCount 100 | ForEach-Object { $lineCount += $_.Count; };
		$lineCount--;		## To remove the count of the header row.
		$delimiter = ",";
		$headerColumns1 = Csv-GetHeaderColumns $strFile1Path $delimiter;
		$headerColumns2 = Csv-GetHeaderColumns $strFile2Path $delimiter;

		#Add all the entries from File1
		Import-Csv $strFile1Path | ForEach-Object { 
			$mergeResult[$_.($headerColumns1[0])] = $_;
		}

		#Add the entries from File2, updating matches and adding unMatched as new entry.
		Import-Csv $strFile2Path | ForEach-Object { 
			$itemCount++;

			if ($mergeResult.ContainsKey($_.($headerColumns2[0]))) {
				#Update Matched entries.
				for ($x = 1; $x -lt $headerColumns2.Count; $x++) {
					$mergeResult[$_.($headerColumns1[0])] | Add-Member NoteProperty -Name $headerColumns2[$x] -Value $_.($headerColumns2[$x]) -Force;
				}
			} else {
				#unMatched entry, so build a Custom PS Object and add it
				$custObject = [PSCustomObject]@{
					$headerColumns1[0] = $_.($headerColumns2[0]);
				};
				$mergeResult[$_.($headerColumns2[0])] = Csv-BuildCustomObject $custObject $_ $headerColumns1 $headerColumns2;
			}

			if (Get-Command ('ShowProgress') -ErrorAction SilentlyContinue) {
				ShowProgress $itemCount $lineCount "Merging files is in Progress" "line";
			}
		}

		$objReturn.Results = $True
		$objReturn.Message = "Success";
		$objReturn.Returns = $mergeResult;

		return $objReturn;
	}

	function Csv-MergeTwoFilesMatchesOnly {
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)]$strFile1Path, 
			[ValidateNotNull()][Parameter(Mandatory=$True)]$strFile2Path
		)
		#Description: Merge entries from file 2 that match file 1, skips non matches.
		#    Not fully coded.
		#Returns: a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters passed in.
			#$objReturn.Results		= $True or $False.  Was a zip file created.
			#$objReturn.Message		= "Success" or the error message.
			#$objReturn.Returns		= The merged file results.
		#Inputs:
			#$strFile1Path = First file to merge.
			#$strFile2Path = Second file to merge.

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

		$file1 = Import-Csv -Path $strFile1Path;
		$file2 = Import-Csv -Path $strFile2Path;

		$mergeResult = @();

		foreach ($line in $file1) {
			$match = $file2 | Where-Object { $_.one -eq $line.one };
			if ($match) {
				$row = "" | Select-Object one, two, three, four, five, six, seven, eight, nine;

				$row.one = $line.one;
				$row.two = $line.two;
				$row.three = $line.three;
				$row.four = $line.four;
				$row.five = $line.five;
				$row.six = $line.six;
				$row.seven = $line.seven;

				$row.eight = $match.eight;
				$row.nine = $match.nine;

				$mergeResult += $row;
			}
		}

		$objReturn.Results = $True
		$objReturn.Message = "Success";
		$objReturn.Returns = $mergeResult;

		return $objReturn;
	}

	function DeCode{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strBase64String, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strOutPut
		);
		#Description....
			#$Content = [System.Convert]::FromBase64String($Base64)
			#Set-Content -Path $env:temp\AM.dll -Value $Content -Encoding Byte
		#Returns: 
		#$strBase64String = String to be Decoded
		#$strOutPut = How to output.  "Display" or [Output file with full path]

		$strContent = [System.Convert]::FromBase64String($strBase64String);
		if (($strOutPut -ne "") -and ($strOutPut -ne $null) -and ($strOutPut -ne "Display")){
			Set-Content -Path $strOutPut -Value $strContent -Encoding Byte;
		}else{
			#Write-Host $strContent;
			$strContent = [System.Text.Encoding]::ASCII.GetString($strContent);
			Write-Host $strContent;
			return $strContent;
		}
	}

	function DeCodeFile{
		#A place holder.  Should be using DeCode() instead of this one.
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strBase64String, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strOutPut
		);
		#A place holder.  Should be using DeCode() instead of this one.
		#Returns: 

		DeCode $strBase64String $strOutPut;
	}

	function Encode{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strFile, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strOutPut
		);
		#Description....
			#$Content = Get-Content -Path C:\AM\AM.dll -Encoding Byte
			#$Base64 = [System.Convert]::ToBase64String($Content)
			#$Base64 | Out-File c:\AM\encoded.txt 
		#Returns: 
		#$strFile = File to be Decoded
		#$strOutPut = How to output.  "Display" or [Output file with full path]

		$strContent = Get-Content -Path $strFile -Encoding Byte;
		$strBase64 = [System.Convert]::ToBase64String($strContent);
		if (($strOutPut -ne "") -and ($strOutPut -ne $null) -and ($strOutPut -ne "Display")){
			$strBase64 | Out-File $strOutPut;
		}else{
			Write-Host $strBase64;
			return $strBase64;
		}
	}

	function EncodeFile{
		#A place holder.  Should be using Encode() instead of this one.
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strFile, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strOutPut
		);
		#A place holder.  Should be using Encode() instead of this one.
		#Returns: 

		Encode $strFile $strOutPut;
	}

	function ExcelCreateOpenFile{
		param(
			[ValidateNotNull()][Parameter(Mandatory = $True, HelpMessage = "Excel file path.")][string] $ExcelFilePath, 
			[ValidateNotNull()][Parameter(Mandatory = $False, HelpMessage = "Create a new Excel file if not exist.")][bool] $CreateNew = $True, 
			[ValidateNotNull()][Parameter(Mandatory = $False, HelpMessage = "Excel Window visibility.")][bool] $ExcelVisible = $True, 
			[ValidateNotNull()][Parameter(Mandatory = $False, HelpMessage = "Open WorkBook ReadOnly.")][bool] $AsReadOnly = $False
		)
		#Description....
		#Returns: 

		$cultureUS = [System.Globalization.CultureInfo]'en-US';
		[System.Threading.Thread]::CurrentThread.CurrentCulture = $cultureUS;

		#temporary continue if error, because it stops even when we want to continue, then return to prior state.
		$ErrorActionPreference = "Continue";
		$application = try{[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')}catch{$null};
		$ErrorActionPreference = "Stop";

		if(-not $application){
			$application = New-Object -comobject Excel.Application;
		}
		else{
			if (($application.Visible -eq $True) -or ($ExcelVisible -eq $True)){
				$ExcelVisible = $True;
			}
			else{
				#Excel is running already, so use the visibility of the current session.
				$ExcelVisible = $application.Visible;
			}
		}

		$application.Visible = $ExcelVisible;

		if (Test-Path $ExcelFilePath){
			if ($AsReadOnly){
				#Open the source file in ReadOnly mode.
				$workbook = $application.Workbooks.Open($ExcelFilePath, $null, $True);
			}else{
				#Open the file normally.
				$workbook = $application.Workbooks.Open($ExcelFilePath, 2, $False);
			}
		}
		else{
			if($CreateNew){
				$workbook = $application.Workbooks.Add();
				$workbook.SaveAs($ExcelFilePath);																		#appears to default to Excel 2007/2010 format
				#$workbook.SaveAs($ExcelFilePath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlExcel12);			#Excel 2007/2010
				#$workbook.SaveAs($ExcelFilePath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlExcel8);				#Excel 95/97/2003
			}else{
				$workbook = $null;
			}
		}

		#we need to return also application, because of option to setup it later
		return ($application, $workbook);
	}

	function ExcelGetData{
		param(
			[ValidateNotNull()][Parameter(Mandatory = $True, HelpMessage = "Excel workbook object.")][object] $WorkBook, 
			[ValidateNotNull()][Parameter(Mandatory = $False)] $WorkSheet, 
			[ValidateNotNull()][Parameter(Mandatory = $False)] $bShowStatus = $False
		)
		#Get data off a WorkSheet.
		#Returns: a DataTable (Returns a DataSet, if no WorkSheet provided).

		if (($WorkSheet -eq "") -or ($WorkSheet -eq $null)){
			#No WorkSheet name provided, so get ALL WorkSheets.
			#$objDataSet = New-Object System.Data.DataSet;
			#---=== repeat as needed ===---
			#$objDataTable = New-Object System.Data.DataTable;
			#objDataSet.Tables.Add($objDataTable);
			#---=== repeat as needed ===---

			#$objData = $objDataSet;
		}
		else{
			#Get the data off $WorkSheet.
			#Code based heavly from:
				#https://podlisk.wordpress.com/2011/11/20/import-excel-spreadsheet-into-powershell/

			#Following has ide on a faster? way to import the data, maybe:
				#http://stackoverflow.com/questions/7023140/how-to-remove-empty-rows-from-datatable

			$objDataTable = New-Object System.Data.DataTable;

			$iNumCols = $WorkSheet.UsedRange.Columns.Count;
			$iNumRows = $WorkSheet.UsedRange.Rows.Count;
			$iHeaderRow = 1;

			#Determine Header Row.
			if ($WorkSheet.Range("A1").MergeCells -eq $True){
				$iHeaderRow = 2;
			}

			#Get the column headers
			for ($intCol = 1; $intCol -le $iNumCols; $intCol ++) {
				$fieldName = $WorkSheet.Cells.Item.Invoke($iHeaderRow, $intCol).Value2;
				if (($fieldName -eq $null) -or ($fieldName -eq "")){
					$fieldName = "Column" + $intCol.ToString();
				}
				#$fieldName = $fieldName.Replace(" ", "");
				$fieldName = $fieldName.Replace("`r`n", "");

				#Write-Host "Adding column $fieldName";
				$Error.Clear();
				$objCol = New-Object System.Data.DataColumn $fieldName,([String]);
				$strResult = $objDataTable.Columns.Add($objCol);
				#$strResult = $objDataTable.Columns.Add($objCol) | Out-Null;
				if ($Error){
					$Error.Clear();
					$fieldName = $fieldName + "1";
					#Write-Host "Adding column $fieldName again";
					#Try again.
					$objCol = New-Object System.Data.DataColumn $fieldName,([String]);
					$strResult = $objDataTable.Columns.Add($objCol);
				}
			}

			#Get the rows of data
			$intNumBlanks = 0;
			for ($line = ($iHeaderRow + 1); $line -le $iNumRows; $line ++) {
				$objRow = $objDataTable.NewRow();

				#Write-Host "Adding row $line...";
				for ($intCol = 1; $intCol -le $iNumCols; $intCol ++) {
					$strVal = $WorkSheet.Cells.Item.Invoke($line, $intCol).Value2;
					#$objRow.($objDataTable.Columns[($intCol - 1)].ColumnName) = $WorkSheet.Cells.Item.Invoke($line, $intCol).Value2;
					$objRow.($objDataTable.Columns[($intCol - 1)].ColumnName) = $strVal;

					if ($intCol -eq 1){
						if (($strVal -eq "") -or ($strVal -eq $null)){
							$intNumBlanks = $intNumBlanks + 1;
						}
						else{
							$intNumBlanks = 0;
						}
					}
				}

				if ($intNumBlanks -gt 2){
					#Found 3 blank rows (Column 1) together, assume rest are blank and break out of For loop.
					#Write-Host "Breaking out.";
					break;
				}

				$objDataTable.Rows.Add($objRow);
				if ($bShowStatus){
					$dblPercent = [math]::round((($line/$iNumRows) * 100), 2);
					Write-Host "$dblPercent % complete.  (Added row $line of $iNumRows)";
				}
			}

			# Remove empty lines
			$Columns = $objDataTable.Columns.Count;
			$Rows = $objDataTable.Rows.Count;
			#for ($r = 0; $r -lt $Rows; $r++) {
			for ($r = $Rows; $r -ge 0; $r--) {
				$Empty = 0;
				if ($objDataTable.Rows[$r] -ne $null) {
					for ($c = 0; $c -lt $Columns; $c++) {
						if ($objDataTable.Rows[$r].IsNull($c)) {
							$Empty++;
						}
					}
					if ($Empty -eq $Columns) {
						# Mark row for deletion:
						$objDataTable.Rows[$r].Delete();
					}
				}
			}
			# Delete marked rows:
			$objDataTable.AcceptChanges();

			#Now return the DataTable
			$objData = $objDataTable;
		}

		return ,$objData;
	}

	function ExcelGetWorksheet{
		param(
			[ValidateNotNull()][Parameter(Mandatory = $True, HelpMessage = "Excel workbook object.")][object] $Workbook, 
			[ValidateNotNull()][Parameter(Mandatory = $False)][string] $SheetName
		)
		#Description....
		#Returns: a WorkSheet object, or an array of available WorkSheets.

		if (($SheetName -eq "") -or ($SheetName -eq $null)){
			#Return an array of the available sheets
			$worksheet = @();
			foreach ($sheet in $Workbook.Worksheets){
				#Write-Host $sheet.Name;
				$worksheet += $sheet.Name;
			}
		}
		else{
			#Return a WorkSheet object.
			$worksheet = $Workbook.Worksheets | where {$_.name -eq $SheetName};
			#$worksheet = $Workbook.Sheets.Item($SheetName);

			if (-not $worksheet){
				$worksheet = $Workbook.Worksheets.Add();
				$worksheet.name = $SheetName;
			}
		}

		return $worksheet;
	}

	function ExcelWriteData{
		param(
			[ValidateNotNull()][Parameter(Mandatory = $True, HelpMessage = "Excel file path.")][string] $ExcelFilePath,
			[ValidateNotNull()][Parameter(Mandatory = $True, HelpMessage = "Object with input data e. g. hashtable, array, ...")][object] $InputData 
		)
		#Description....
			#Has some basic info on writing to Excel
				#http://stackoverflow.com/questions/28101802/powershell-excel-combine-worksheets-into-a-single-worksheet
			#This one might be better
				#https://gist.github.com/fergus/8553443
		#Returns: 

		#Add next sheet for 'Test Case Overview'
		($application, $workbook) = ExcelCreateOpenFile -ExcelFilePath $ExcelFilePath;
		$worksheetTC = ExcelGetWorksheet -Workbook $workbook -SheetName "Data";
		$row = 1;
		$col = 1;
		$cells = $worksheetTC.Cells;

		#if $InputData is simple array
		foreach($data in $InputData){
			#write values
			$cells.item($row,$col) = $data.ToString();
			$row++;

			#define cell name and create hyperlink to other cell
			$cellValue = $data.ToString();
			$cellName = "o1_{0}" -f $cellValue;
			($cells.Item($row,$col)).Name = $cellName;
			$targetCellName = "o2_{0}" -f $cellValue;
			$subAddress = "'{0}'!{1}" -f $sheetName2, $targetCellName;		#"'Test Overview'!A1"
			$void = $worksheetTC.Hyperlinks.Add($cells.Item($row,$col) ,"" , $subAddress, "", $cellValue);
		 
		}

		#read values
		#if($cells.Item($row, $col).Value() -eq "Id"){
		#	$row++;
		#	$cells.item($row, $col) = "cat";
		#}

		#turn off Error message for replacing existing file when saving it
		$application.DisplayAlerts = $False;
		$workbook.SaveAs($ExcelFilePath);
		#$application.Quit();
	}

	function ExcelXML{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$DocPath,
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$Query,
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$WorksheetName,
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$ListValues
		)
		#Description....
		#Returns: 

		$strAction = "";
		#Make sure we have defaults, and know what action is being done.
		if ((($WorksheetName -eq "") -or ($WorksheetName -eq $null)) -and (($Query -eq "") -or ($Query -eq $null)) -and (($ListValues -eq "") -or ($ListValues -eq $null))){
			#Only $DocPath provided, so GetSheet()
			#No need to set defaults.
			$strAction = "GetSheets";
		}
		else{
			if (($ListValues -eq "") -or ($ListValues -eq $null)){
				#No $ListValues provided, so GetData().
				$strAction = "Query";
				if (($WorksheetName -eq "") -or ($WorksheetName -eq $null)){
					#Query.  $Query was provided, so $WorksheetName NOT be needed.
					#$WorksheetName = "Sheet1";
					#Lets make sure $Query is populated too.
					if (($Query -eq "") -or ($Query -eq $null)){
						$Query = 'SELECT * FROM [Sheet1$]';
					}
				}
				else{
					#Query.  $WorksheetName was provided, but $Query may not have been.
					#$WorksheetName was provided, lets make sure $Query is populated too.
					if (($Query -eq "") -or ($Query -eq $null) -or ((!($Query -Match $WorksheetName)) -and ($Query.ToUpper().StartsWith("SELECT")))){
						#Both of the following do the same thing.
						#$Query = 'SELECT * FROM [{0}$]' -f $WorksheetName;
						$Query = "SELECT * FROM [" + $WorksheetName + '$]';
					}
				}

				# Make sure the query is in the correct syntax (e.g. 'SELECT * FROM [SheetName$]')
				$Pattern = '.*from\b\s*(?<Table>\w+).*'
				if($Query -match $Pattern) {
					$Query = $Query -replace $Matches.Table, ('[{0}$]' -f $Matches.Table)
				}
			}
			else{
				#Must be Create()
				$strAction = "Create";
				#$WorksheetName has no default value, but probably should.
				if (($WorksheetName -eq "") -or ($WorksheetName -eq $null)){
					$WorksheetName = "Sheet1";
				}
			}
		}

		#Check what Providers are available.
		#http://stackoverflow.com/questions/6649363/microsoft-ace-oledb-12-0-provider-is-not-registered-on-the-local-machine
		$objOLEProviders = (New-Object system.data.oledb.oledbenumerator).GetElements() | select SOURCES_NAME, SOURCES_DESCRIPTION;
		$objOLEProviders = ($objOLEProviders | Where-Object { $_.SOURCES_NAME -like "Microsoft.*" } | Sort-Object SOURCES_NAME);
		#$Provider = ((New-Object System.Data.OleDb.OleDbEnumerator).GetElements() | Where-Object { $_.SOURCES_NAME -like "Microsoft.ACE.OLEDB*" } | Sort-Object SOURCES_NAME -Descending | Select-Object -First 1 SOURCES_NAME).SOURCES_NAME;
		#$Provider = ($objOLEProviders | Where-Object { $_.SOURCES_NAME -like "Microsoft.ACE.OLEDB*" } | Sort-Object SOURCES_NAME -Descending | Select-Object -First 1 SOURCES_NAME).SOURCES_NAME;
			#Should be able to use "Invoke-Command" to run a ScriptBlock using the "Microsoft.PowerShell32" Configuration.
				#But does not work.
			#http://www.ravichaganti.com/blog/powershell-2-0-remoting-guide-part-9-%E2%80%93-session-configurations-and-creating-custom-configurations/
			#Get-PSSessionConfiguration
			#Register-PSSessionConfiguration Microsoft.PowerShell32 -processorarchitecture x86 -force
			#[ScriptBlock]$ScriptBlock = {((New-Object System.Data.OleDb.OleDbEnumerator).GetElements() | Where-Object { $_.SOURCES_NAME -like "Microsoft.ACE.OLEDB*" } | Sort-Object SOURCES_NAME -Descending | Select-Object -First 1 SOURCES_NAME).SOURCES_NAME;};
			#Invoke-Command -ScriptBlock $ScriptBlock -ConfigurationName Microsoft.PowerShell32;

		#Following errors (in 64 bit PS).
		#[Reflection.Assembly]::LoadFrom("C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE14\ACEOLEDB.DLL")
		#http://stackoverflow.com/questions/31545746/unable-to-load-net-assembly-in-powershell
			#... You can't load a 32 bit dll in a 64 bit process or vice versa unless the dll was compiled for Any Cpu....
			#... bad image format exceptions usually happen when you try to load a non-.net assembly, or if you try to load a differing .net assembly....

		#Check if the file is XLS or XLSX
		#http://danielcai.blogspot.com/2011/02/solution-run-jet-database-engine-on-64.html
		#if ((Get-Item -Path $DocPath).Extension -eq '.xls'){
		if (($strDocPath.EndsWith(".xls")) -and ($env:Processor_Architecture -eq "x86")){
			#32Bit only
			$Provider = "Microsoft.Jet.OLEDB.4.0";
			$ExtendedProperties = "Excel 8.0;HDR=YES;IMEX=1";
		}
		else{
			#32Bit or 64bit, depending on version of office installed.
			#http://blog.sqlauthority.com/2015/06/24/sql-server-fix-export-error-microsoft-ace-oledb-12-0-provider-is-not-registered-on-the-local-machine/
				#32Bit office -> "C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE14\ACEOLEDB.DLL"
				#64Bit office -> "C:\Program Files\Common Files\Microsoft Shared\OFFICE14\ACEOLEDB.DLL"
			$Provider = "Microsoft.ACE.OLEDB.12.0";
			$ExtendedProperties = "Excel 12.0;HDR=YES";
		}

		# Build the connection string and connection object
		#$ConnectionString = 'Provider={0};Data Source={1};Extended Properties="{2}"' -f $Provider, $DocPath, $ExtendedProperties
		$ConnectionString = "Provider=" + $Provider + ";Data Source=" + $DocPath + ";Extended Properties=" + $ExtendedProperties;
		$Connection = New-Object System.Data.OleDb.OleDbConnection $ConnectionString

		try{
			# Open the connection to the file, and fill the datatable
			$Connection.Open()

			#Create
			if ($strAction -eq "Create"){
				$Command = $Connection.CreateCommand()
				$Command.CommandText = "CREATE TABLE [$WorksheetName] ($ListValues)";
				$Command.ExecuteNonQuery();
			}

			#GetData
			if ($strAction -eq "Query"){
				$Adapter = New-Object -TypeName System.Data.OleDb.OleDbDataAdapter $Query, $Connection
				$DataTable = New-Object System.Data.DataTable
				$Adapter.Fill($DataTable) | Out-Null
			}

			#GetSheets
			if ($strAction -eq "GetSheets"){
				$DataTable = $Connection.GetSchema("Tables")
			}
		}
		catch{
			# something went wrong :-(
			Write-Error $_.Exception.Message
		}
		finally{
			# Close the connection
			if ($Connection.State -eq 'Open') {
				$Connection.Close()
			}
		}

		# Return the results NOT as an array
		return ,$DataTable
	}

	function ParseLogFile{
		param(
			[ValidateNotNull()][parameter(Mandatory=$True, HelpMessage='Source File, with path')][String]$SourceFile,
			[ValidateNotNull()][parameter(Mandatory=$True, HelpMessage='Search criteria')][String]$SearchString,
			[ValidateNotNull()][parameter(Mandatory=$False, HelpMessage='Destination Path for Report')][String]$DestPath
		)
		#Description....
		#Returns: 
		#$SourceFile = The log file to search through.  (i.e. "\\Nawespscfs101v.nadsuswe.nads.navy.mil\isf-ios$\ITSS-Tools\Logs\AScII\20160425_BO-12864827_CreatedBy_AScII.log")
		#$SearchString = The string to look for in each line.  (i.e. " INTO Transactions " or "283821")
		#$DestPath = The destination path to generate the "Filtered" log file to.
			#ParseLogFile "\\Nawespscfs101v.nadsuswe.nads.navy.mil\isf-ios$\ITSS-Tools\Logs\AScII\20160426_AScII.log" "283821" "C:\Projects\PS-Scripts\Tests\";
			#ParseLogFile $SourceFile $SearchString $DestPath;
			#ParseLogFile $SourceFile $SearchString;

		Write-Host "`r`n";
		Write-Host "Parsing log '$SourceFile'.";
		$ElapsedTime = [System.Diagnostics.Stopwatch]::StartNew();

		$intLineCount = 0;
		$intLinesFound = 0;
		if (Test-Path $SourceFile){
			$objFile = Get-Item $SourceFile -Force;
			if ($objFile.PSIsContainer -eq $False){
				#File
				if ($objFile.Length -gt 0){
					if (([String]::IsNullOrEmpty($DestPath)) -or ((Test-Path $DestPath) -eq $False)){
						if ([String]::IsNullOrEmpty($DestPath)){
							Write-Host "Destination path was not provided, so using:";
						}
						else{
							Write-Host "Destination path '$DestPath' does not exist, so using:";
						}
						$DestPath = $objFile.Directory.FullName;
					}
					if ($DestPath.EndsWith("\") -eq $False){
						$DestPath = $DestPath + "\";
					}
					Write-Host "  $DestPath";

					$ErrorActionPreference = 'SilentlyContinue';
					$strOutFile = $DestPath + ($objFile.Name).SubString(0, ($objFile.Name.Length - $objFile.Extension.Length)) + "-Filtered-(" + $SearchString + ")" + $objFile.Extension;
					$Error.Clear();
					$objWriter = New-Object System.IO.StreamWriter($strOutFile);
					if ($Error){
						$Error.Clear();
						$strOutFile = $DestPath + ($objFile.Name).SubString(0, ($objFile.Name.Length - $objFile.Extension.Length)) + "-Filtered" + $objFile.Extension;
						$objWriter = New-Object System.IO.StreamWriter($strOutFile);
					}
					$ErrorActionPreference = 'Continue';
					if ($Error){
						Write-Host "Error tying to setup the output file.";
					}
					else{
						$intTotalLines = ([IO.File]::ReadAllLines($SourceFile)).Count;
						Write-Host "Processing $intTotalLines total lines.";
						$intDivisor = [Int]("1" + ("0" * (([String]$intTotalLines).Length - 2)));

						$objStream = $objFile.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite);
						$objReader = New-Object System.IO.StreamReader($objStream);
						while (!$objReader.EndOfStream){
							$intLineCount++;
							#Display a PowerShell progress bar
								#https://msdn.microsoft.com/powershell/reference/5.1/microsoft.powershell.utility/Write-Progress
							$dblPercentComp = (($intLineCount / $intTotalLines) * 100);
							if ($dblPercentComp -gt 100){
								$dblPercentComp = 100;
							}
							Write-Progress -Activity "Parsing log $SourceFile" -Status "Processing $intTotalLines total lines." -PercentComplete $dblPercentComp;

							if (!(([String]($intLineCount / $intDivisor)).Contains("."))){
								Write-Host "." -NoNewline;
							}
							$strLine = $objReader.ReadLine();
							if ($strLine -match $SearchString){
								$intLinesFound++;
								$objWriter.Write($strLine.Trim() + "`r`n");
							}
						}
						Write-Host "`r`n";
						$objWriter.Close();
					}
				}
				else{
					Write-Host "The log file provided '$SourceFile', is empty (size 0).";
				}
			}
			else{
				#Directory
				Write-Host "The log file provided '$SourceFile', is a directory.";
			}
		}
		else{
			#File not found
			Write-Host "The log file provided '$SourceFile' could not be found.";
		}

		Write-Host "Took $($ElapsedTime.Elapsed.ToString()), to filter out $intLinesFound lines (from the $intLineCount total lines).";
		Write-Host "Results have been written to: "
		Write-Host $strOutFile;
		Write-Host "`r`n"
	}

	function URLSaveToFile{
		Param(
			[Parameter(Mandatory=$True)][String]$strUrl,
			[Parameter(Mandatory=$True)][String]$strDestFolder
		)
 		#Description....
		#Returns: 

		$objResponse = $null;
		$Error.Clear();
		$objResponse = Invoke-WebRequest -Uri $strUrl;
		if ((!($Error)) -and (!([String]::IsNullOrEmpty($objResponse)))){
			$strFilename = [System.IO.Path]::GetFileName($objResponse.BaseResponse.ResponseUri.OriginalString);
			$strFilename = $strFilename.Replace("%20", " ");
			$objFilepath = [System.IO.Path]::Combine($strDestFolder, $strFilename);
			try{
				$Error.Clear();
				$objFilename = [System.IO.File]::Create($objFilepath);
				if ($Error){
					$objFilepath = [System.IO.Path]::Combine($strDestFolder, ((([System.DateTime]::Now).Ticks).ToString()));
					$objFilename = [System.IO.File]::Create($objFilepath);
				}
				$objResponse.RawContentStream.WriteTo($objFilename);
				$objFilename.Close();
			}
			finally{
				if ($objFilename){
					$objFilename.Dispose();
				}
			}
		}
	}

	function ZipCreateFile{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$ZipFile, 
			[ValidateNotNull()][Parameter(Mandatory=$True)][Array]$Files
		)
		#Description....
		#Returns: a PowerShell object.
			#$objReturn.Name		= Name of this process, with paramaters passed in.
			#$objReturn.Results		= $True or $False.  Was a zip file created.
			#$objReturn.Message		= "Success" or the error message.
			#$objReturn.Returns		= The full path annd file name of the file created.
		#$ZipFile = The zip file to create. (Full path) [i.e. "c:\path\file.zip"]
		#$Files = An array of the files to add to the zip file. (Full paths) [i.e. @("c:\path\file.one", "c:\path\file.two")]

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

		if ((Test-Path -Path $ZipFile)){
			#File exists, so delete it.
			Remove-Item $ZipFile;
		}

		if (!(Test-Path -Path $ZipFile)){
			$Error.Clear();
			#http://www.adminarsenal.com/admin-arsenal-blog/powershell-zip-up-files-using-.net-and-add-type
			#Above link is for Powershell 3 and .NET 4.5.

			#http://stackoverflow.com/questions/1153126/how-to-create-a-zip-archive-with-powershell
			#This is adding an "extra" xml file, but do I care?
			#Load assemblys.
			$Results = [System.Reflection.Assembly]::Load("WindowsBase, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35");
			#Create the zip file.
			$objZip = [System.IO.Packaging.ZipPackage]::Open($ZipFile, [System.IO.FileMode]"OpenOrCreate", [System.IO.FileAccess]"ReadWrite");
			#Setup the Array of files to loop through.
			#$arrFiles = @("c:\file.one", "c:\file.two");
			#$arrFiles = $Files -Replace "C:", "" -Replace "\\", "/";
			$arrFiles = $Files;
			foreach ($objFile in $arrFiles){
				#For each file you want to add, we must extract the bytes and add them to a part of the zip file.
				#$partName = New-Object System.Uri($objFile, [System.UriKind]"Relative");
				$partName = New-Object System.Uri(($objFile -Replace "C:", "" -Replace "\\", "/"), [System.UriKind]"Relative");
				#$partName = New-Object System.Uri($objFile, [System.UriKind]"Absolute");
				#$partName = New-Object System.Uri(($objFile -Replace "C:", "" -Replace "\\", "/"), [System.UriKind]"Absolute");
				#Create each part. 
				$part = $objZip.CreatePart($partName, "application/zip", [System.IO.Packaging.CompressionOption]"Maximum");
				#$bytes = [System.IO.File]::ReadAllBytes($objFile) | out-null;
				$bytes = [System.IO.File]::ReadAllBytes($objFile);
				$stream = $part.GetStream();
				$stream.Write($bytes, 0, $bytes.Length);
				$stream.Close();
			}
			#Close the zip file when we're done.
			$objZip.Close();

			if ((Test-Path -Path $ZipFile) -and (!$Error)){
				$objReturn.Results = $True
				$objReturn.Message = "Success";
				$objReturn.Returns = $ZipFile;
			}else{
				$objReturn.Message = "Error `r`n" + $Error;
			}
		}else{
			$objReturn.Message = "Error, File exists already.";
		}

		return $objReturn;
	}
