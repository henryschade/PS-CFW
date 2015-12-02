###########################################
# Updated Date:	2 December 2015
# Purpose:		Code to manipulate Documents.
# Requirements: None
##########################################

	function ExcelSampleReadCOM{
		$strFilePath = "C:\Projects\PS-Scripts\Testing\CIVMAR Bulk MAC.xls";
		$WorkSheetName = "Create User Account";

		#Opening the file.
		($objExcel, $objWorkBook) = ExcelCreateOpenFile -ExcelFilePath $strFilePath $False $False;

		#Check if got the workbook
		if ($objWorkBook){
			#Get an Array of all available sheets:
			$arrSheets = ExcelGetWorksheet $objWorkBook;
			#Get a WorkSheet object.
			#$objSheets = ExcelGetWorksheet $objWorkBook $arrSheets[1];
			#or one of the following:
			#$objSheets = ExcelGetWorksheet $objWorkBook "SheetName";
			#$objWorkSheet = ExcelGetWorksheet -Workbook $objWorkBook -SheetName "East";
			$objWorkSheet = ExcelGetWorksheet -Workbook $objWorkBook -SheetName $WorkSheetName;

			#Get an object of the Cells
			#$objCells = $objWorkSheet.Cells;

			#Read from the worksheet.
			#Write-Host $objCells.Item($intRow, $intCol).Value();
			#Write-Host $objCells.Item(1, 1).Value();					#A1
			#Write-Host $objCells.Item(1, 2).Value();					#B1
			#Write-Host $objCells.Item(2, 1).Value();					#A2
			#Write-Host $objCells.Item(2, 2).Value();					#B2

			#or don't bother with the Cells object and just to the following
			#Write-Host $objWorkSheet.Range("A2").Text;					#A2

			##Clean up / Release Cells object
			#[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objCells) | Out-Null;
			#$objCells = $null;

			#Clean up / Release WorkSheet object
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objWorkSheet) | Out-Null;
			$objWorkSheet = $null;

			#Close the workbook.
			#Turn off Error messages.
			#$objExcel.DisplayAlerts = $False;
			$objWorkBook.Close();
			#Turn Error messages back on.
			#$objExcel.DisplayAlerts = $True;
			#Clean up / Release WorkBook object
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objWorkBook) | Out-Null;
		}

		#Quit/Close Excel.
		$objExcel.Quit();
			#Excel still shows in TaskManager...
			#When close Powershell Excel will close too (this is a .NET using COM objects issue).
		#Clean up / Release Excel object
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel) | Out-Null;
		$objExcel = $null;

	}

	function ExcelSampleUsageCOM{
		# The Excel Code is from following URL, but modified for my use.
		# http://mypowershell.webnode.sk/news/create-or-open-excel-file-in-powershell/

		#$ExcelFilePath = "c:\temp\MyExcelFile2.xlsx"
		#data 'raz, 'dva', 'tri' will be inserted to excel sheet 'Data'
		#ExcelWriteData -InputData @{"raz", "dva", "tri") -ExcelFilePath $ExcelFilePath



		#What file to open/create
		$strFilePath = "\\nawesdnifs08.nadsuswe.nads.navy.mil\NMCIISF\NMCIISF-SDCP-MAC\MAC\Entr_SRM\Support Files\BackUpLocation\USN_Server_Farms-Testing.xls";
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

	function ExcelSampleUsageXML1{
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
		# Run the code in a 32bit job, since the provider is 32bit only
		$job = Start-Job $JobCode -RunAs32 -ArgumentList $strDocPath, $Query, $WorkSheetName, $ListValues
		$job | Wait-Job | Receive-Job
		Remove-Job $job


		#GetSheets
		#Create the scriptblock to run in a job
			#$JobCode = [scriptblock]::create($function:ExcelXML);		#$strDocPath $Query;
		# Run the code in a 32bit job, since the provider is 32bit only
		$job = Start-Job $JobCode -RunAs32 -ArgumentList $strDocPath
		$job | Wait-Job | Receive-Job
		Remove-Job $job


		#GetSheetData
		#Create the scriptblock to run in a job
			#$JobCode = [scriptblock]::create($function:ExcelXML);		# $strDocPath $Query;
		# Run the code in a 32bit job, since the provider is 32bit only
		$job = Start-Job $JobCode -RunAs32 -ArgumentList $strDocPath, $Query
		$job | Wait-Job | Receive-Job
		Remove-Job $job
	}

	function ExcelSampleUsageXML2{
		$strDocPath = "C:\Projects\PS-Scripts\Testing\CIVMAR Bulk MAC.xls";
		$WorkSheetName = "Create User Account";
		$Query = "";
		$ListValues = "";


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
		. "C:\Projects\PS-Scripts\Documents.ps1"

		EncodeFile "C:\Settings.txt" "C:\EncSet.txt"
		EncodeFile "C:\Settings.txt" "Display"

		DeCodeFile "....JUeXBlID0gbXNzcWwNCnN0ckRCU2VydmVyID0gT......" "C:\Settings.txt"
		DeCodeFile "....JUeXBlID0gbXNzcWwNCnN0ckRCU2VydmVyID0gT......" "Display"

		$strEncode = "c3RyREJUeXBlID0gbXNzcWwNCnN0ckRCU2VydmVyID0gTkFFQU5SRktTUTc1VkFcU1E3NVZBSU5TVDAxDQpzdHJEQk5hbWUgPSBTaXRlQ29kZXMNCnN0ckRCTG9naW5SID0gS0J1c2VyDQpzdHJEQlBhc3NSID0ga2M1JHNxMDI=";
		DeCodeFile $strEncode "C:\Users\henry.schade\Desktop\SQL.txt"
		DeCodeFile $strEncode "Display"
		EncodeFile "C:\Users\henry.schade\Desktop\SQL.txt" "C:\Users\henry.schade\Desktop\EncSet.txt"
	}


	function DeCode{
		Param(
			[Parameter(Mandatory=$True)][String]$strBase64String, 
			[Parameter(Mandatory=$False)][String]$strOutPut
		);

		#$Content = [System.Convert]::FromBase64String($Base64)
		#Set-Content -Path $env:temp\AM.dll -Value $Content -Encoding Byte

		$strContent = [System.Convert]::FromBase64String($strBase64String);
		if (($strOutPut -ne "") -and ($strOutPut -ne $null) -and ($strOutPut -ne "Display")){
			Set-Content -Path $strOutPut -Value $strContent -Encoding Byte;
		}else{
			#Write-Host $strContent;
			$strContent = [System.Text.Encoding]::ASCII.GetString($strContent);
			Write-Host $strContent;
		}
	}

	function DeCodeFile{
		#A place holder.  Should be using DeCode() instead of this one.
		Param(
			[Parameter(Mandatory=$True)][String]$strBase64String, 
			[Parameter(Mandatory=$False)][String]$strOutPut
		);

		DeCode $strBase64String $strOutPut;
	}

	function Encode{
		Param(
			[Parameter(Mandatory=$True)][String]$strFile, 
			[Parameter(Mandatory=$False)][String]$strOutPut
		);

		#$Content = Get-Content -Path C:\AM\AM.dll -Encoding Byte
		#$Base64 = [System.Convert]::ToBase64String($Content)
		#$Base64 | Out-File c:\AM\encoded.txt 

		$strContent = Get-Content -Path $strFile -Encoding Byte;
		$strBase64 = [System.Convert]::ToBase64String($strContent);
		if (($strOutPut -ne "") -and ($strOutPut -ne $null) -and ($strOutPut -ne "Display")){
			$strBase64 | Out-File $strOutPut;
		}else{
			Write-Host $strBase64;
		}
	}

	function EncodeFile{
		#A place holder.  Should be using Encode() instead of this one.
		Param(
			[Parameter(Mandatory=$True)][String]$strFile, 
			[Parameter(Mandatory=$False)][String]$strOutPut
		);

		Encode $strFile $strOutPut;
	}



	function ExcelCreateWorksheetXML{
		#Code from Trev, that treats an Excel doc like an xml doc to pull data.
		Param(
			[Parameter(Mandatory=$true)][String] $Path,
			[Parameter(Mandatory=$true)][String] $WorksheetName,
			[Parameter(Mandatory=$true)][String] $ListValues
		)
		$JobCode = {
			Param($Path,$WorkSheetName,$ListValues)
		}
	}

	function ExcelGetDataXML{
		#Code from Trev, that treats an Excel doc like an xml doc to pull data.
		[CmdletBinding(DefaultParameterSetName='Worksheet')]
		Param(
			[Parameter(Mandatory=$true, Position=0)][String] $Path,
			[Parameter(Position=1, ParameterSetName='Worksheet')][String] $WorksheetName = 'Sheet1',
			[Parameter(Position=1, ParameterSetName='Query')][String] $Query = 'SELECT * FROM [Sheet1$]'
		)
		$JobCode = {
			Param($Path, $Query)
		}
	}

	function ExcelGetSheetsXML{
		#Code from Trev, that treats an Excel doc like an xml doc to pull data.
		Param(
			[Parameter(Mandatory=$true)][String] $Path
		)
		$JobCode = {
			Param($Path, $Query)
		}
	}


	function ExcelCreateOpenFile{
		param(
			[ValidateNotNull()][Parameter(Mandatory = $True, HelpMessage = "Excel file path.")][string] $ExcelFilePath, 
			[ValidateNotNull()][Parameter(Mandatory = $False, HelpMessage = "Create a new Excel file if not exist.")][bool] $CreateNew = $True, 
			[ValidateNotNull()][Parameter(Mandatory = $False, HelpMessage = "Excel Window visibility.")][bool] $ExcelVisible = $True, 
			[ValidateNotNull()][Parameter(Mandatory = $False, HelpMessage = "Open WorkBook ReadOnly.")][bool] $AsReadOnly = $False
		)

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
			#Excel is running already, so use the visibility of the current session.
			$ExcelVisible = $application.Visible;
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

	function ExcelGetWorksheet{
		param(
			[ValidateNotNull()][Parameter(Mandatory = $True, HelpMessage = "Excel workbook object.")][object] $Workbook, 
			[ValidateNotNull()][Parameter(Mandatory = $False)][string] $SheetName
		)

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


		# Check if the file is XLS or XLSX 
		#if ((Get-Item -Path $DocPath).Extension -eq '.xls'){
		if ($strDocPath.EndsWith("xls")){
			$Provider = 'Microsoft.Jet.OLEDB.4.0'
			$ExtendedProperties = 'Excel 8.0;HDR=YES;IMEX=1'
		} else {
			$Provider = 'Microsoft.ACE.OLEDB.12.0'
			$ExtendedProperties = 'Excel 12.0;HDR=YES'
		}

		# Build the connection string and connection object
		$ConnectionString = 'Provider={0};Data Source={1};Extended Properties="{2}"' -f $Provider, $DocPath, $ExtendedProperties
		$Connection = New-Object System.Data.OleDb.OleDbConnection $ConnectionString

		try {
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
		} catch {
			# something went wrong :-(
			Write-Error $_.Exception.Message
		}
		finally {
			# Close the connection
			if ($Connection.State -eq 'Open') {
				$Connection.Close()
			}
		}

		# Return the results NOT as an array
		return ,$DataTable
	}


	function URLSaveFile{
		Param(
			[Parameter(Mandatory=$True)][String]$strUrl,
			[Parameter(Mandatory=$True)][String]$strDestFolder
		)
 
		$objResponse = $null;
		$Error.Clear();
		$objResponse = Invoke-WebRequest -Uri $strUrl
		if ((!($Error)) -and ($objResponse -ne "") -and ($objResponse -ne $null)){
			$strFilename = [System.IO.Path]::GetFileName($objResponse.BaseResponse.ResponseUri.OriginalString)
			$strFilename = $strFilename.Replace("%20", " ")
			$objFilepath = [System.IO.Path]::Combine($strDestFolder, $strFilename)
			try{
				$Error.Clear();
				$objFilename = [System.IO.File]::Create($objFilepath)
				if ($Error){
					$objFilepath = [System.IO.Path]::Combine($strDestFolder, ((([System.DateTime]::Now).Ticks).ToString()))
					$objFilename = [System.IO.File]::Create($objFilepath)
				}
				$objResponse.RawContentStream.WriteTo($objFilename)
				$objFilename.Close()
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
		#Returns a PowerShell object.
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

