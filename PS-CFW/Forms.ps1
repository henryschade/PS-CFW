###########################################
# Updated Date:	5 June 2015
# Purpose:		My functions to create PS Forms and Controls
# Requirements: None
# Web sites that helped me:
#				http://bytecookie.wordpress.com/2011/07/17/gui-creation-with-powershell-the-basics/
#				http://blogs.technet.com/b/heyscriptingguy/archive/2011/07/24/create-a-simple-graphical-interface-for-a-powershell-script.aspx
#				http://social.technet.microsoft.com/Forums/scriptcenter/en-US/4a625d51-3016-4a2b-a643-c5eab6def599/powershell-how-to-return-object-in-a-function?forum=ITCG
#				http://blogs.technet.com/b/stephap/archive/2012/04/23/building-forms-with-powershell-part-1-the-form.aspx
#				http://blogs.technet.com/b/heyscriptingguy/archive/2010/03/24/hey-scripting-guy-march-24-2010.aspx
#				Working w/ ListBoxes - http://technet.microsoft.com/en-us/library/ff730950.aspx
#				ListViews - http://social.technet.microsoft.com/Forums/scriptcenter/en-US/553f06bc-522c-4854-9e28-d0e219a789a6/powershell-and-systemwindowsformslistview?forum=ITCG
#				PictureBox  - http://powershell.com/cs/forums/t/13511.aspx
#				MessageBox - http://powershell-tips.blogspot.com/2012/02/display-messagebox-with-powershell.html
#				InputBox - http://windowsitpro.com/blog/getting-input-and-inputboxes-powershell
#
#				Form Events: - http://msdn.microsoft.com/en-us/library/system.windows.forms.form_events(v=vs.110).aspx
#					List an objects Properties/Events/Methods/etc:  -  http://stackoverflow.com/questions/7377959/how-to-find-properties-of-an-object
##########################################

	#Import the Assemblies
	#Add-Type -AssemblyName System.Windows.Forms;

	#[Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null;
	#[Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null;

	#loading the necessary .net libraries (using void to suppress output)
	[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms");
	[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing");

	Function AddControl($objParent, $strWhatType, $strName, $intWidth, $intHeight, $strText, $intLeft, $intTop, $intTabIndex, $strClickCode, $arrValues, $bolChecked, $bolMultiLine, $strImage){
		#$objParent = The Form (or Parent Control) to put the Control on.  (i.e. for Radio buttons could be the GroupBox).
		#$strWhatType = What type of Control to Add.
			#Current options = Button, Checkbox, ComboBox, GroupBox, Label, Listbox, PictureBox, RadioButton, TextBox
		#$strName = The name of the Control.
		#$intWidth = The width of the Control.
		#$intHeight = The height of the Control.
		#$strText = The text of the Control.
		#$intLeft = The left of the Control.
		#$intTop = The top of the Control.
		#$intTabIndex = The tab index of the Control.
		#$strClickCode = The code to run when the Control is clicked.
		#$arrValues = The values to populate the ComboBox or ListBox with (an array).
		#$bolChecked = Should the Control (RadioButton or CheckBox) be Checked.
		#$bolMultiLine = If a control (TextBox) should be mulitline.
		#$strImage = Path to the image to put on the control (PicturBox).

		#$objSystem_Size = New-Object System.Drawing.Size;
		#$objSystem_Location = New-Object System.Drawing.Point;

		Switch ($strWhatType) {
			"Button"{
				$objControl = New-Object System.Windows.Forms.Button;
				##$objSystem_Size = New-Object System.Drawing.Size;
				#$objSystem_Size.Width = $intWidth;
				#$objSystem_Size.Height = $intHeight;
				#$objControl.Size = $objSystem_Size;
				$objControl.Size = New-Object System.Drawing.Size($intWidth, $intHeight)
				##$objSystem_Location = New-Object System.Drawing.Point;
				#$objSystem_Location.X = $intLeft;
				#$objSystem_Location.Y = $intTop;
				#$objControl.Location = $objSystem_Location;
				$objControl.Location = New-Object System.Drawing.Point($intLeft, $intTop)

				#$objControl.Name = $strName;
				#$objControl.TabIndex = $intTabIndex;
				$objControl.Text = $strText;
				#$objControl.Font = New-Object System.Drawing.Font("Verdana",14,[System.Drawing.FontStyle]::Bold);

				$objControl.BackColor = "#CCCCCC";		# color names are static properties of System.Drawing.Color you can also use ARGB values, such as "#FFFFEBCD"
				#$objControl.UseVisualStyleBackColor = $True;
				$objControl.Cursor = [System.Windows.Forms.Cursors]::Hand;

				#$objControl.Add_Click($strClickCode);
				Break
			}
			"Checkbox"{
				$objControl = New-Object System.Windows.Forms.Checkbox;
				$objControl.Size = New-Object System.Drawing.Size($intWidth, $intHeight);
				#$objControl.Location = new-object System.Drawing.Size(10,10)
				$objControl.Location = New-Object System.Drawing.Point($intLeft, $intTop);

				$objControl.Text = $strText;
				If (($bolChecked -eq "True") -or ($bolChecked -eq "yes")){
					$objControl.Checked = $True;
				}else{
					$objControl.Checked = $False;
				}
			}
			"ComboBox"{
				$objControl = New-Object System.Windows.Forms.ComboBox;
				$objControl.Size = New-Object System.Drawing.Size($intWidth, $intHeight);
				$objControl.Location = New-Object System.Drawing.Point($intLeft, $intTop);

				#$objControl.Name = $strName;
				#$objControl.TabIndex = $intTabIndex;

				$objControl.DropDownHeight = 200;

				#$arrValues=@("Value1", "Value2", "etc.")
				ForEach ($strEntry in $arrValues){
					$objControl.Items.Add($strEntry);
				}
			}
			"GroupBox"{
				$objControl = New-Object System.Windows.Forms.GroupBox;
				$objControl.Size = New-Object System.Drawing.Size($intWidth, $intHeight);
				$objControl.Location = New-Object System.Drawing.Point($intLeft, $intTop);

				#$objControl.Name = $strName;
				#$objControl.TabIndex = $intTabIndex;
				$objControl.Text = $strText;
			}
			"Label"{
				$objControl = New-Object System.Windows.Forms.Label;
				##$objSystem_Size = New-Object System.Drawing.Size;
				#$objSystem_Size.Width = $intWidth;
				#$objSystem_Size.Height = $intHeight;
				#$objControl.Size = $objSystem_Size;
				$objControl.Size = New-Object System.Drawing.Size($intWidth, $intHeight);
				##$objSystem_Location = New-Object System.Drawing.Point;
				#$objSystem_Location.X = $intLeft;
				#$objSystem_Location.Y = $intTop;
				#$objControl.Location = $objSystem_Location;
				$objControl.Location = New-Object System.Drawing.Point($intLeft, $intTop);

				#$objControl.AutoSize = $True
				#$objControl.Name = $strName;
				#$objControl.TabIndex = $intTabIndex;
				$objControl.Text = $strText;

				#$objControl.Add_Click($strClickCode);
				Break
			}
			"Listbox"{
				$objControl = New-Object System.Windows.Forms.Listbox;
				#$objControl.Size = New-Object System.Drawing.Size(260,20);
				$objControl.Size = New-Object System.Drawing.Size($intWidth, $intHeight);
				#$objControl.Location = New-Object System.Drawing.Size(10,40);
				$objControl.Location = New-Object System.Drawing.Point($intLeft, $intTop);

				# there are only two real differences between a standard list box and a multi-select list box:
				#1) in a multi-select list box you must assign a value to the SelectionMode property; and,
				#2) in a multi-select list box you must work with an array of selected items rather than a single selected item.
				#"MultiExtended", "MultiSimple"
				$objControl.SelectionMode = "MultiExtended";

				#[void] $objControl.Items.Add("Item 1")
				#[void] $objControl.Items.Add("Item 2")
				#$arrValues=@("Value1", "Value2", "etc.")
				ForEach ($strEntry in $arrValues){
					$objControl.Items.Add($strEntry);
				}
				#$objControl.SelectedItem = "Item 2";
				#$objControl.SelectedItem = $arrValues[1];

				$objControl.Height = 70
			}
			"PictureBox"{
				$strImage = [System.Drawing.Image]::Fromfile($strImage);

				$objControl = New-Object Windows.Forms.PictureBox;
				#$objControl.Width = $strImage.Size.Width;
				#$objControl.Height =  $strImage.Size.Height;
				if (($intWidth -ne $null) -and ($intWidth -ne "") -and ($intHeight -ne $null) -and ($intHeight -ne "")){
					$objControl.Size = New-Object System.Drawing.Size($intWidth, $intHeight);
				}else{
					$objControl.Size = New-Object System.Drawing.Size($strImage.Size.Width, $strImage.Size.Height);
				}
				$objControl.Image = $strImage;
				$objControl.Location = new-object System.Drawing.Point($intLeft, $intTop);

				if (($strImage.Size.Width -ne $intWidth) -or ($strImage.Size.Height -ne $intHeight)){
					$objControl.SizeMode = "Zoom";
				}
			}
			"RadioButton"{
				$objControl = New-Object System.Windows.Forms.RadioButton;
				$objControl.Size = New-Object System.Drawing.Size($intWidth, $intHeight);
				$objControl.Location = new-object System.Drawing.Point($intLeft, $intTop);

				#$objControl.Name = $strName;
				#$objControl.TabIndex = $intTabIndex;
				$objControl.Text = $strText;
				If (($bolChecked -eq "true") -or ($bolChecked -eq "yes")){
					$objControl.Checked = $True;
				}else{
					$objControl.Checked = $False;
				}
			}
			"TextBox"{
				$objControl = New-Object System.Windows.Forms.TextBox;
				##$objSystem_Size = New-Object System.Drawing.Size;
				#$objSystem_Size.Width = $intWidth;
				#$objSystem_Size.Height = $intHeight;
				#$objControl.Size = $objSystem_Size;
				$objControl.Size = New-Object System.Drawing.Size($intWidth, $intHeight);
				##$objSystem_Location = New-Object System.Drawing.Point;
				#$objSystem_Location.X = $intLeft;
				#$objSystem_Location.Y = $intTop;
				#$objControl.Location = $objSystem_Location;
				$objControl.Location = new-object System.Drawing.Point($intLeft, $intTop);

				#$objControl.Name = $strName;
				#$objControl.TabIndex = $intTabIndex;
				$objControl.Text = $strText;
				#$objControl.Font = New-Object System.Drawing.Font("Verdana",8,[System.Drawing.FontStyle]::Italic);
				#$objControl.Multiline=$False;
				$objControl.Multiline=$bolMultiLine;
				#$objControl.ScrollBars = "Vertical";

				#$objControl.Add_Click($strClickCode);
				Break
			}
			Default {
				Throw "No match for `$strWhatType: $strWhatType";
			}
		}

		#Attributes that are common to ALL controls:
		If (($strName -ne $null) -and ($strName -ne "")){
			$objControl.Name = $strName;
		}
		If (($intTabIndex -ne $null) -and ($intTabIndex -ne "")){
			$objControl.TabIndex = $intTabIndex;
		}
		#$objControl.Size = New-Object System.Drawing.Size($intWidth, $intHeight);
		#$objControl.Location = new-object System.Drawing.Point($intLeft, $intTop);
		If (($strClickCode -ne $null) -and ($strClickCode -ne "")){
			$objControl.Add_Click($strClickCode);
		}

		If ($objControl -ne $null){
			$objParent.Controls.Add($objControl);
		}

		Return $objControl;

	}

	Function Calendar{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strTitle, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Int]$NumV = 1, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Int]$NumH = 1, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Bool]$CircToday = $False
		)

		if (($strTitle -eq "") -or ($strTitle -eq $null)){
			$strTitle = "Select a Date";
		}
		if (($NumV -lt 1) -or ($NumV -eq $null)){
			$NumV = 1;
		}
		if (($NumH -lt 1) -or ($NumH -eq $null)){
			$NumH = 1;
		}
		if (($NumH * $NumV) -gt 12){
			if ($NumH -ge 12){
				$NumH = 12;
				$NumV = 1;
			}elseif ($NumH -ge 6){
				$NumH = 6;
				$NumV = 2;
			}elseif ($NumH -ge 4){
				$NumH = 4;
				$NumV = 3;
			}elseif ($NumH -ge 3){
				$NumH = 3;
				$NumV = 4;
			}elseif ($NumH -ge 2){
				$NumH = 2;
				$NumV = 6;
			}elseif ($NumH -ge 1){
				$NumH = 1;
				$NumV = 12;
			}else{
				MsgBox "def";
				$NumH = 4;
				$NumV = 3;
			}
		}

		#https://technet.microsoft.com/en-us/library/ff730942.aspx
		$objForm = New-Object Windows.Forms.Form 

		#$objForm.Text = "Select a Date"
		$objForm.Text = $strTitle;
		#$objForm.Size = New-Object Drawing.Size @(190,190)
		$intHeader = 55;
		$intPadding = 5;
		$objForm.Size = New-Object Drawing.Size @(($intPadding + ($NumH * 184)), ($intHeader + ($NumV * 141)))
		$objForm.StartPosition = "CenterScreen"

		$objForm.KeyPreview = $True

		$objForm.Add_KeyDown({
			if ($_.KeyCode -eq "Enter"){
				$dtmDate = $objCalendar.SelectionStart;
				$objForm.Close();
			}
			if ($_.KeyCode -eq "Escape"){
				$objForm.Close();
			}
		})

		$objCalendar = New-Object System.Windows.Forms.MonthCalendar 
		$objCalendar.CalendarDimensions = New-Object Drawing.Size @($NumH, $NumV)
		$objCalendar.ShowTodayCircle = $CircToday
		$objCalendar.MaxSelectionCount = 1
		$objForm.Controls.Add($objCalendar) 

		$objForm.Topmost = $True

		$objForm.Add_Shown({$objForm.Activate()})  
		[void] $objForm.ShowDialog() 

		if ($dtmDate){
			#Write-Host "Date selected: $dtmDate"
			return $dtmDate;
		}

	}

	Function CreateForm($strName, $strText, $intWidth, $intHeight, $strOnLoadCode, $strWinState, $bolShowInTaskBar){
		#$strName = The name of the Form.
		#$strText = The text of the Form.
		#$intWidth = The width of the Form.
		#$intHeight = The height of the Form.
		#$strOnLoadCode = The code to run when the Form is loaded.
		#$strWinState = The state the Window/Form should start in.  Maximized, Minimized, Normal (default)
		#$bolShowInTaskBar = Show the Form in the TaksBar.  True or False

		$objForm = New-Object System.Windows.Forms.Form;
		<#
		$strInitFormWindowState = New-Object System.Windows.Forms.FormWindowState;

		$OnLoadForm_StateCorrection = {
			#Correct the initial state of the form to prevent the .Net maximized form issue
			$objForm.WindowState = $strInitFormWindowState;
		}
		#>

		If (($bolShowInTaskBar -eq "") -or ($bolShowInTaskBar -eq $null) -or ($bolShowInTaskBar -eq "True")){
			$bolShowInTaskBar = $True;
		}else{
			$bolShowInTaskBar = $False;
		}
		If (($strWinState -eq "") -or ($strWinState -eq $null)){
			$strWinState = "Normal";
		}

		$objForm.Text = $strText;
			#$objFont = New-Object System.Drawing.Font("Times New Roman",18,[System.Drawing.FontStyle]::Italic)
				# Font styles are: Regular, Bold, Italic, Underline, Strikeout
			#$objForm.Font = $objFont
		$objForm.Name = $strName;
		#$objForm.DataBindings.DefaultDataSourceUpdateMode = 0;
		#$System_Drawing_Size = New-Object System.Drawing.Size;
		#$System_Drawing_Size.Width = $intWidth;
		#$System_Drawing_Size.Height = $intHeight;
		#$objForm.ClientSize = $System_Drawing_Size;
		#$objForm.ClientSize = "284,262";
		#$objForm.ClientSize = $intWidth, $intHeight;
		$objForm.Width = $intWidth;
		$objForm.Height = $intHeight;
		$objForm.MinimumSize = New-Object System.Drawing.Size($intWidth, $intHeight);
		$objForm.StartPosition = "CenterScreen";
			# CenterScreen, Manual, WindowsDefaultLocation, WindowsDefaultBounds, CenterParent

		#$objForm.AutoScroll = $True;
		#Setting AutoSizeMode will prevent the user from manually resizing the form. 
		#$objForm.AutoSizeMode = "GrowAndShrink";
		#	# or GrowOnly
		#$objForm.MinimizeBox = $False;
		#$objForm.MaximizeBox = $False;
		#$objForm.WindowState = "Normal";
		#	# Maximized, Minimized, Normal
		$objForm.WindowState = $strWinState;
		#$objForm.SizeGripStyle = "Hide";
		#	# Auto, Hide, Show
		#$objForm.ShowInTaskbar = $True;
		$objForm.ShowInTaskbar = $bolShowInTaskBar;
		#$objForm.Opacity = 0.7;
		#	# 1.0 is fully opaque; 0.0 is invisible
		if ($bolShowInTaskBar -eq $True){
			#Icon from an image file 
			#$objIcon = New-Object system.drawing.icon ("C:\Program Files\Microsoft Office\Office14\GRAPH.ICO");
			#Icon extracted from a file
			$objIcon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe");
			$objForm.Icon = $objIcon;
		}

		<#
		#Save the initial state of the form
		$strInitFormWindowState = $objForm.WindowState;
		#Init the OnLoad event to correct the initial state of the form
		$objForm.Add_Load($OnLoadForm_StateCorrection);
		#>

		#$objForm.Add_Shown({$objForm.Activate()});

		##Show the Form
		#$objForm.ShowDialog()| Out-Null;

		Return $objForm;

	}

	Function MsgBox{
		Param(
			[ValidateNotNull()][Parameter(Mandatory=$True)][String]$strMessage, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strTitle, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][Int]$intButtons, 
			[ValidateNotNull()][Parameter(Mandatory=$False)][String]$strDefault
		)
		#$strMessage = The message to put in the Message Box.
		#$strTitle = The title to give the Message Box.
		#$intButtons = The buttons and/or type of Message/Input Box to show.
			#0: OK 
			#1: OK Cancel 
			#2: Abort Retry Ignore 
			#3: Yes No Cancel 
			#4: Yes No 
			#5: Retry Cancel
			#6: InputBox
		#$strDefault = The default value to put in an InputBox.

		if (($intButtons -eq $null) -or ($intButtons -eq "") -or ($intButtons -gt 6) -or ($intButtons -lt 0)){
			$intButtons = 0;
		}

		if ($intButtons -eq 6){
			if ($strDefault -eq $null){
				$strDefault = "";
			}

			[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null;
			$strReturn = [Microsoft.VisualBasic.Interaction]::InputBox($strMessage, $strTitle, $strDefault);
		}else{
			$strReturn = [System.Windows.Forms.MessageBox]::Show($strMessage, $strTitle, $intButtons);
		}

		return $strReturn;

		#Another option is:
		##http://gallery.technet.microsoft.com/scriptcenter/1a386b01-b1b8-4ac2-926c-a4986ac94fed
		#$objShell = new-object -comobject wscript.shell;
		#$objPopUp = $objShell.popup("You must run this as admin.", 0, "Not Admin", 1);
	}

	Function ObjectInfo($objObject, $strWantWhat){
		#$objObject = The object to get the Properties/Events/Methods/etc of.
		#$strWantWhat = What to return (wildcards are allowed).  Default is all.  Common options are: "Event", "Method", "Property", "Prop*".

		#Sample usage:
		#[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms");
		#$objForm = CreateForm "frmTestingForm" "Testing Form" 225 280 "" "";
		#C:\..Path..\PS-Forms.ps1 $objForm Events

		if(($strWantWhat -eq $null) -or ($strWantWhat -eq "")){
			$strWantWhat = "";
			$strResults = $objObject | Get-Member | Format-Table -AutoSize -Property MemberType, Name;
		}else{
			$strResults = $objObject | Get-Member | Where {$_.MemberType -Match $strWantWhat} | Format-Table -AutoSize -Property MemberType, Name;
		}

		#$strResults = $objObject | Get-Member | Select -Property MemberType, Name | Format-Table -AutoSize;
		#$strResults = $objObject | Get-Member | Format-Table -AutoSize -Property MemberType, Name;

		return $strResults;
	}



	# All the code below here is for my testing.
	Function btnButton1_Click{
		$objTextBox3.TEXT = "";
		#frmTestingForm
		$objTxbResults = $Window | Get-ChildControl txbResults;

		$strFirstName=$objTextBox1.TEXT;
		$strLastName=$objTextBox2.TEXT;

		#Write-Host txbFirst.TEXT ~ txbLast.TEXT;
		$Username = $strFirstName + "." + $strLastName;
		#Write-Host $Username;
		$objTextBox3.TEXT = $Username;

		##"C:\Users\hschade\Desktop\"
		#$strTest = dir $objTextBox1.TEXT;
		#$strTest2 = "";
		#foreach ($line in $strTest) {
		#	#$strTest2.Appendtext($line + [char]13 + [char]10);
		#	$strTest2 = $strTest2 + $line + [char]13 + [char]10;
		#}
		#Write-Host $strTest2;

		#<#
		if (($Username -ne "") -and ($Username -ne $null) -and ($Username -ne ".")){
			$ScriptDir = Split-Path $MyInvocation.MyCommand.Path
			#$ScriptDir = "\\nawesdnifs08.nadsuswe.nads.navy.mil\NMCIISF\NMCIISF-SDCP-MAC\MAC\Entr_SRM\Support Files\";
			. ($ScriptDir + "\PS-ExchConn.ps1")

			SetupConn "w"
			$strTest = Get-MailboxStatistics $Username | Select DisplayName, TotalItemSize, ItemCount, TotalDeletedItemSize, StorageLimitStatus, ServerName, DatabaseName;
			CleanUpConn

			#Write-Host $strTest;
			$objTextBox3.TEXT = $objTextBox3.TEXT + "`r`n`r`n" + $strTest;
		}
		#>
	}

	Function btnButton2_Click{
		#Close the Form
		$objForm.Close();
		#Kill the Form
		$objForm.Dispose();
	}

	Function MyTestForm{
		$intLeft1 = 10;
		$intLeft2 = 120;
		$intTop1 = 10;
		$intTop2 = 30;
		$intTabIndex = 0;

		$objForm = CreateForm "frmTestingForm" "Testing Form" 225 280 "" "";

		$objLabel1 = AddControl $objForm "Label" "lblFirst" 80 15 "First Name" $intLeft1 $intTop1;
		$objTextBox1 = AddControl $objForm "TextBox" "txbFirst" 80 20 "" $intLeft1 $intTop2 $intTabIndex;
		$intTabIndex += 1;

		$objLabel2 = AddControl $objForm "Label" "lblLast" 80 15 "Last Name" $intLeft2 $intTop1;
		$objTextBox2 = AddControl $objForm "TextBox" "txbLast" 80 20 "" $intLeft2 $intTop2 $intTabIndex;
		$intTabIndex += 1;

		$objButtonClick = {btnButton1_Click};
		$objButton1 = AddControl $objForm "Button" "btnButton1" 90 20 "Get Mail Size" $intLeft1 55 $intTabIndex $objButtonClick;
		$intTabIndex += 1;

		$objButtonClick = {btnButton2_Click};
		$objButton2 = AddControl $objForm "Button" "btnButton2" 80 20 "Close" $intLeft2 55 $intTabIndex $objButtonClick;
		$intTabIndex += 1;

		$objLabel3 = AddControl $objForm "Label" "lblResults" 80 15 "Results" $intLeft1 80;
		$objTextBox3 = AddControl $objForm "TextBox" "txbResults" 190 135 "" $intLeft1 95 "" "" "" "" $True;

		#Show the Form
		#$objForm.ShowDialog()| Out-Null;
		[void] $objForm.ShowDialog();
	}

	if ($args[0] -eq "Form"){
		MyTestForm;

		#Write-Host "Done";
	}else{
		if ($args[0] -eq "Calendar"){
			Calendar "Date" 2;
		}else{
			if ($args[0] -ne $null){
				#Sample usage:
				#[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms");
				#$objForm = CreateForm "frmTestingForm" "Testing Form" 225 280 "" "";
				#C:\..Path..\PS-Forms.ps1 $objForm Events

				ObjectInfo $args[0] $args[1];
			}
		}
	}

	#Write-Host ""
	#Write-Host "Press any key to continue ..."
	#$x = $host.UI.RawUI.ReadKey("NoEcho, IncludeKeyDown")
