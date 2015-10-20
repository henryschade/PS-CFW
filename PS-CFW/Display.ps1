###########################################
# Updated Date:	1 June 2015
# Purpose:		Started out for Hide/Show PowerShell Console Window, then added 
# 				Display Orientation Methods, now it is for all display related routines.
# Requirements: None
##########################################

	#http://www.aspnet-answers.com/microsoft/Powershell/30523953/invisible-windows.aspx

	#http://blogs.msdn.com/b/frankfi/archive/2008/08/13/changing-the-display-resolution-in-a-multi-monitor-environment.aspx

	function DisplaySampleUsage{
		[cDisplaySettings]::GetDispOrientation();

		#Flip display to 180 degrees (Upside down)
		[cDisplaySettings]::SetDispOrientation(2);

		start-sleep 10;
		#Do code here instead of Sleep.

		#Flip display to 0  degrees (normal)
		[cDisplaySettings]::SetDispOrientation(0);

		[cDisplaySettings]::GetDispProps();
	}

	function ConsoleSampleUsage1{
		#Hide the PowerShell Console.
		[ConsoleHelper]::HideConsole();

		start-sleep 10;
		#Do code here instead of Sleep.

		#To show the PowerShell Console.
		[ConsoleHelper]::ShowConsole();
	}

	function ConsoleSampleUsage2{
		#Hide the PowerShell Console.
		$ch::HideConsole();

		start-sleep 10;
		#Do code here instead of Sleep.

		#To show the PowerShell Console.
		$ch::ShowConsole();
	}



#Pinvoke'd C# code.
$DisplayCode = @"
	using System;
	using System.Runtime.InteropServices;

	public class cDisplaySettings
	{
		[StructLayout(LayoutKind.Sequential)]
		public struct DEVMODE
		{
			private const int CCHDEVICENAME = 0x20;
			private const int CCHFORMNAME = 0x20;

			//[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 0x20)]
			[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
			public string dmDeviceName;
			public short dmSpecVersion;
			public short dmDriverVersion;
			public short dmSize;
			public short dmDriverExtra;
			public int dmFields;
			public int dmPositionX;
			public int dmPositionY;
			//public ScreenOrientation dmDisplayOrientation;
			public short dmOrientation;
			public int dmDisplayFixedOutput;
			public short dmColor;
			public short dmDuplex;
			public short dmYResolution;
			public short dmTTOption;
			public short dmCollate;
			//[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 0x20)]
			[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
			public string dmFormName;
			public short dmLogPixels;
			public int dmBitsPerPel;
			public int dmPelsWidth;
			public int dmPelsHeight;
			public int dmDisplayFlags;
			public int dmDisplayFrequency;
			public int dmICMMethod;
			public int dmICMIntent;
			public int dmMediaType;
			public int dmDitherType;
			public int dmReserved1;
			public int dmReserved2;
			public int dmPanningWidth;
			public int dmPanningHeight;
		}

		[DllImport("user32.dll")]
		public static extern int EnumDisplaySettings(string deviceName, int modeNum, ref DEVMODE devMode);
		[DllImport("user32.dll")]
		public static extern int ChangeDisplaySettings(ref DEVMODE devMode, int flags);

		public const int ENUM_CURRENT_SETTINGS = -1;
		public const int CDS_UPDATEREGISTRY = 0x01;
		public const int CDS_TEST = 0x02;
		public const int DISP_CHANGE_SUCCESSFUL = 0;
		public const int DISP_CHANGE_RESTART = 1;
		public const int DISP_CHANGE_FAILED = -1;

		public const int DMDO_DEFAULT = 0;
		public const int DMDO_90 = 1;
		public const int DMDO_180 = 2;
		public const int DMDO_270 = 3;

		//static private DEVMODE GetDevMode(){
		static public DEVMODE GetDevMode(){
			DEVMODE dm = new DEVMODE();
			dm.dmDeviceName = new String(new char[32]);
			dm.dmFormName = new String(new char[32]);
			dm.dmSize = (short)Marshal.SizeOf(dm);
			return dm;
		}

		static public string GetDispOrientation(){
			DEVMODE dm = cDisplaySettings.GetDevMode();

			if (0 != cDisplaySettings.EnumDisplaySettings(null, cDisplaySettings.ENUM_CURRENT_SETTINGS, ref dm)){
				//At this point the DEVMODE structure will be populated with the Display settings and can be modified at any time.
				//return dm.dmOrientation.ToString();
					//0 = normal
					//1 = 90
					//2 = 180
					//3 = 270

					//Following URL has info on getting supported modes
					//http://www.pinvoke.net/default.aspx/coredll.changedisplaysettingsex
					// modes are as follows: 0 = 0, 1 = 90, 2 = 180, 4 = 270 degrees

				switch (dm.dmOrientation.ToString()){
					case "0":
					{
						return "normal";
					}
					case "1":
					{
						return "90";
					}
					case "2":
					{
						return "180";
					}
					case "3":
					{
						return "270";
					}
					default:
					{
						return "UnKnown";
					}
				}
			}
			else{
				return "Failed To Get Current Display Settings.";
			}
		}

		static public DEVMODE GetDispProps(){
			DEVMODE dm = cDisplaySettings.GetDevMode();

			cDisplaySettings.EnumDisplaySettings(null, cDisplaySettings.ENUM_CURRENT_SETTINGS, ref dm);

			return dm;
		}

		static public string SetDispOrientation(int intSetting){
			DEVMODE dm = cDisplaySettings.GetDevMode();

			if (0 != cDisplaySettings.EnumDisplaySettings(null, cDisplaySettings.ENUM_CURRENT_SETTINGS, ref dm)){
				//At this point the DEVMODE structure will be populated with the Display settings and can be modified at any time.
				//dm.dmOrientation = DMDO_180;		//hopefully can do this.  Only if define the constant.
				switch (intSetting){
					case 0:
					{
						dm.dmOrientation = DMDO_DEFAULT;
						break;
					}
					case 1:
					{
						dm.dmOrientation = DMDO_90;
						break;
					}
					case 2:
					{
						dm.dmOrientation = DMDO_180;
						break;
					}
					case 3:
					{
						dm.dmOrientation = DMDO_270;
						break;
					}
					default:
					{
						dm.dmOrientation = DMDO_DEFAULT;
						break;
					}
				}

				//Test that settings can be applied.
				int iRet = cDisplaySettings.ChangeDisplaySettings(ref dm, cDisplaySettings.CDS_TEST);

				if (iRet == cDisplaySettings.DISP_CHANGE_FAILED){
					return "Unable To Change Display Orientation.";
				}
				else{
					//Apply the settings.
					iRet = cDisplaySettings.ChangeDisplaySettings(ref dm, cDisplaySettings.CDS_UPDATEREGISTRY);
					switch (iRet){
						case cDisplaySettings.DISP_CHANGE_SUCCESSFUL:
						{
							return "Success";
						}
						case cDisplaySettings.DISP_CHANGE_RESTART:
						{
							return "You Need To Reboot For The Change To Take Effect.\n If You Have Any Problems After Rebooting\nYou Will Have To Change The Resolution Back, In Safe Mode.";
						}
						default:
						{
							return "Failed";
						}
					}
				}
			}
			else{
				return "Failed To Get Current Display Settings.";
			}
		}
	}
"@;		#This MUST end w/ no leading spaces.


$ConsoleCode = @'
	using System;
	using System.Runtime.InteropServices;
	public class ConsoleHelper {
		private const Int32 SW_HIDE = 0;
		private const Int32 SW_SHOW = 5;

		[DllImport("user32.dll")]
		private static extern Boolean ShowWindow(IntPtr hWnd, Int32 nCmdShow);
		[DllImport("kernel32.dll", SetLastError = true)]
		public static extern bool AllocConsole();
		[DllImport("Kernel32.dll")]
		private static extern IntPtr GetConsoleWindow();

		public static void HideConsole(){
			IntPtr hwnd = GetConsoleWindow();
			if (hwnd != IntPtr.Zero){
				ShowWindow(hwnd, SW_HIDE);
			}
		}

		public static void ShowConsole(){
			IntPtr hwnd = GetConsoleWindow();
			if (hwnd != IntPtr.Zero){
				ShowWindow(hwnd, SW_SHOW);
			}
		}
	}
'@;		#This MUST end w/ no leading spaces.


	#Load the DisplayCode
	$Error.Clear();
	#Add-Type $DisplayCode
	Add-Type -TypeDefinition $DisplayCode -IgnoreWarnings;
	if ($Error){
		#If the Add-Type() commandlet fails, try this:
		#Can build a CSharpCodeProvider object and load the source code (above) into it.
		#[cDisplaySettings] > $null
		#$DisSet = [cDisplaySettings]
		trap {
			if (($MyInvocation.MyCommand.Path -eq "") -or ($MyInvocation.MyCommand.Path -eq $null)){
				$ScriptDir = "C:\SRM_Apps_N_Tools\PS-Scripts";
			}else{
				$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;				#Gets the directory/path the Script was run from.
			}
			$ScriptDir = $ScriptDir + "\PS.tmp";

			# Get an instance of the CSharp code provider
			$CProv = new-object Microsoft.CSharp.CSharpCodeProvider;
			# And compiler parameters...
			$CPara = New-Object System.CodeDom.Compiler.CompilerParameters;
			$CPara.GenerateInMemory = $True;
			$CPara.GenerateExecutable = $False;
			#$CPara.OutputAssembly = "custom";
			#$CPara.OutputAssembly = "C:\PS.tmp";
			$CPara.OutputAssembly = $ScriptDir;
			$Results = $CProv.CompileAssemblyFromSource($CPara, $DisplayCode);

			# display any errors
			if ($Results.Errors.Count){
				$codeLines = $DisplayCode.Split("`n");
				foreach ($CompErr in $Results.Errors){
					write-host "Error: $($codeLines[$($CompErr.Line - 1)])";
					$CompErr | out-default;
				}
				Throw "Compile failed...";
			}else{
				# don't report the exception
				continue;
			}
		}
	}


	#Load the ConsoleCode
	$Error.Clear();
	##Add-Type $ConsoleCode
	Add-Type -TypeDefinition $ConsoleCode -IgnoreWarnings;
	if ($Error){
		#If the Add-Type() commandlet fails, try this:
		#Can build a CSharpCodeProvider object and load the source code (above) into it.
		[ConsoleHelper] > $null
		$ch = [ConsoleHelper]
		trap {
			if (($MyInvocation.MyCommand.Path -eq "") -or ($MyInvocation.MyCommand.Path -eq $null)){
				$ScriptDir = "C:\SRM_Apps_N_Tools\PS-Scripts";
			}else{
				$ScriptDir = Split-Path $MyInvocation.MyCommand.Path;				#Gets the directory/path the Script was run from.
			}
			$ScriptDir = $ScriptDir + "\PS.tmp";

			# Get an instance of the CSharp code provider
			$CProv = new-object Microsoft.CSharp.CSharpCodeProvider;
			# And compiler parameters...
			$CPara = New-Object System.CodeDom.Compiler.CompilerParameters;
			$CPara.GenerateInMemory = $True;
			$CPara.GenerateExecutable = $False;
			#$CPara.OutputAssembly = "custom";
			#$CPara.OutputAssembly = "C:\PS.tmp";
			$CPara.OutputAssembly = $ScriptDir;
			$Results = $CProv.CompileAssemblyFromSource($CPara, $ConsoleCode);

			# display any errors
			if ($Results.Errors.Count){
				$codeLines = $ConsoleCode.Split("`n");
				foreach ($CompErr in $Results.Errors){
					write-host "Error: $($codeLines[$($CompErr.Line - 1)])";
					$CompErr | out-default;
				}
				Throw "Compile failed...";
			}else{
				# don't report the exception
				continue;
			}
		}
	}else{
		$ch = [ConsoleHelper];
	}

