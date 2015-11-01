###########################################
# Updated Date:	1 November 2015
# Purpose:		Code to interact w/ Documents.
# Requirements: None
##########################################

	# The Excel Code is from following URL, but modified for my use.
	# http://mypowershell.webnode.sk/news/create-or-open-excel-file-in-powershell/

	#$ExcelFilePath = "c:\temp\MyExcelFile2.xlsx"
	#data 'raz, 'dva', 'tri' will be inserted to excel sheet 'Data'
	#ExcelWriteData -InputData @{"raz", "dva", "tri") -ExcelFilePath $ExcelFilePath

	function ExcelSampleUsage{
		#What file to open/create
		$strFilePath = "\\nawesdnifs08.nadsuswe.nads.navy.mil\NMCIISF\NMCIISF-SDCP-MAC\MAC\Entr_SRM\Support Files\BackUpLocation\USN_Server_Farms-Testing.xls";

		#Opening the file, and/or create it, and visible or not.
		($objExcel, $objWorkBook) = ExcelCreateOpenFile -ExcelFilePath $strFilePath $False $False;
		#Check if got the workbook
		if ($objWorkBook){
			#Got a workbook, so now can get the desired sheet.
			$objWorkSheet = ExcelGetWorksheet -Workbook $objWorkBook -SheetName "East";

			#Read the worksheet, or write to the worksheet.
			#$objCells = $objWorkSheet.Cells;

			#Write to the worksheet.
			#$objCells.Item(1, 1) = "A1";
			#$objCells.Item(1, 2) = "A2";
			#$objCells.Item(2, 1) = "B1";
			#$objCells.Item(2, 2) = "B2";

			#Turn off Error message for replacing existing file when saving it.
			#$objExcel.DisplayAlerts = $False;
			#$objWorkBook.SaveAs($strFilePath);																		#appears to default to Excel 2007/2010 format
			#$objWorkBook.SaveAs($strFilePath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlExcel12);			#Excel 2007/2010
			#$objWorkBook.SaveAs($strFilePath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlExcel8);			#Excel 95/97/2003
			#Turn Error message back on.
			#$objExcel.DisplayAlerts = $True;

			#Read from the worksheet.
			#Write-Host $objCells.Item(1, 1).Value();
			#Write-Host $objCells.Item(1, 2).Value();
			#Write-Host $objCells.Item(2, 1).Value();
			#Write-Host $objCells.Item(2, 2).Value();
			#Write-Host $objWorkSheet.Range("A2").Text;

			#Clean up WorkSheet object
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objWorkSheet) | Out-Null;

			#Close the workbook.
			$objWorkBook.Close();
			#Clean up WorkBook object
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objWorkBook) | Out-Null;
		}
		#Quit/Close Excel.
		$objExcel.Quit();
		#Clean up Excel object
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel) | Out-Null;
	}


	function ExcelCreateOpenFile(){
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

		$application.Visible = $ExcelVisible;

		if (Test-Path $ExcelFilePath){
			if ($AsReadOnly){
				#Open the source file in ReadOnly mode.
				$workbook = $application.Workbooks.Open($ExcelFilePath, $null, $True);
			}else{
				#Open the file normally.
				$workbook = $application.Workbooks.Open($ExcelFilePath, 2, $False);
			}
		}else{
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

	function ExcelGetWorksheet(){
		param(
			[ValidateNotNull()][Parameter(Mandatory = $True, HelpMessage = "Excel workbook object.")][object] $Workbook, 
			[ValidateNotNull()][Parameter(Mandatory = $False)][string] $SheetName = "Sheet1"
		)

		$worksheet = $Workbook.Worksheets | where {$_.name -eq $SheetName};

		if (-not $worksheet){
			$worksheet = $Workbook.Worksheets.Add();
			$worksheet.name = $SheetName;
		}

		return $worksheet;
	}

	function ExcelWriteData{
		param(
			[Parameter(Mandatory = $True, HelpMessage = "Excel file path.")][string] $ExcelFilePath,
			[Parameter(Mandatory = $True, HelpMessage = "Object with input data e. g. hashtable, array, ...")][object] $InputData 
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
		#$strTemp = "ZipCreateFile(" + $strTemp.Trim() + ")";
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

