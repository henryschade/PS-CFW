###########################################
# Updated Date:	31 March 2015
# Purpose:		Code to interact w/ Excel
#
# Code from following URL, but modified for my use.
# http://mypowershell.webnode.sk/news/create-or-open-excel-file-in-powershell/
##########################################


	function CreateOrOpenExcelFile(){
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

	function GetExcelWorksheet(){
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

	function Write-DataToExcelFile{
		param(
			[Parameter(Mandatory = $True, HelpMessage = "Excel file path.")][string] $ExcelFilePath,
			[Parameter(Mandatory = $True, HelpMessage = "Object with input data e. g. hashtable, array, ...")][object] $InputData 
		)

		#Add next sheet for 'Test Case Overview'
		($application, $workbook) = CreateOrOpenExcelFile -ExcelFilePath $ExcelFilePath;
		$worksheetTC = GetExcelWorksheet -Workbook $workbook -SheetName "Data";
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

	#$ExcelFilePath = "c:\temp\MyExcelFile2.xlsx"
	#data 'raz, 'dva', 'tri' will be inserted to excel sheet 'Data'
	#Write-DataToExcelFile -InputData @{"raz", "dva", "tri") -ExcelFilePath $ExcelFilePath


	function SampleUsage{
		#What file to open/create
		$strFilePath = "\\nawesdnifs08.nadsuswe.nads.navy.mil\NMCIISF\NMCIISF-SDCP-MAC\MAC\Entr_SRM\Support Files\BackUpLocation\USN_Server_Farms-Testing.xls";

		#Opening the file, and/or create it, and visible or not.
		($objExcel, $objWorkBook) = CreateOrOpenExcelFile -ExcelFilePath $strFilePath $False $False;
		#Check if got the workbook
		if ($objWorkBook){
			#Got a workbook, so now can get the desired sheet.
			$objWorkSheet = GetExcelWorksheet -Workbook $objWorkBook -SheetName "East";

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
