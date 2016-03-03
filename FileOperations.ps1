##########################################
# Updated Date:	4 February 2016
# Purpose:		Routines to set NTFS permissions, set owners on files/folders, create/delete shares, and set share permissions.
##########################################


function WINSHARE_EXAMPLE_USAGE{
	#Local system:
	$sSystem = "127.0.0.1";
	$sLocalPath = "C:\SRM_Apps_N_Tools\PS-Scripts\Testing";
	$sShareName = "Test$";
	$sShareComment = "Share created by PS FileOperations Processes.";

	#FileServer system:
	$strUserName = "redirect.test";
	$sDirectory = "\\NAEAPHILFS11\USER01\usr2\" + $strUserName;
	$sSystem = "NAEAPHILFS11";
	$sLocalPath = "C:\vol\USER01\usr2\" + $strUserName;
	$sShareName = $strUserName + "$";
	$sShareComment = "Share created by PS FileOperations Processes.";


	#Add/Create a Share. ($sShareComment is optional)
	$Results = [FileOperations.WinShare]::Add($sSystem, $sLocalPath, $sShareName, $sShareComment);
	#Check the Results of the Share create.
	#https://msdn.microsoft.com/en-us/library/windows/desktop/ms681381(v=vs.85).aspx
	switch ($Results){
		{(($Results -Contains "NerrSuccess" -or $Results -Contains "1338"))}{
			#Success
			#1338 too
		}
		"ErrorAccessDenied"{
			#AccessDenied
		}
		"ErrorInvalidParameter"{
			#InvalidParameter
		}
		"ErrorInvalidName"{
			#InvalidName
		}
		"ErrorInvalidLevel"{
			#InvalidLevel
		}
		"NerrUnknownDevDir"{
			#UnknownDir
		}
		"NerrRedirectedPath"{
			#RedirectedPath
		}
		"NerrDuplicateShare"{
			#DuplicateShare
		}
		"NerrBufTooSmall"{
			#BufTooSmall
		}
		default{
			#UnexpectedResponse
		}
	}

	#Add default share permissions ("Authenticated Users" w/ "Full Control")
	$Results = [FileOperations.WinShare]::SetDefaultSharePermissions($sSystem, $sShareName);

	#Delete a Share
	$Results = [FileOperations.WinShare]::Delete($sSystem, $sShareName);
	#Check the Results of the Share delete.
	switch ($Results){
		"NerrSuccess"{
			#Success
		}
		"ErrorAccessDenied"{
			#AccessDenied
		}
		"ErrorInvalidParameter"{
			#InvalidParameter
		}
		"ErrorInvalidName"{
			#InvalidName
		}
		"ErrorInvalidLevel"{
			#InvalidLevel
		}
		"NerrUnknownDevDir"{
			#UnknownDir
		}
		"NerrRedirectedPath"{
			#RedirectedPath
		}
		"NerrDuplicateShare"{
			#DuplicateShare
		}
		"NerrBufTooSmall"{
			#BufTooSmall
		}
		default{
			#UnexpectedResponse
		}
	}
}

function SetNTFS_EXAMPLE_USAGE{
	#http://stackoverflow.com/questions/3282656/setting-inheritance-and-propagation-flags-with-set-acl-and-powershell
	#+-------------------------------------------------------------------------------------------------------------------------------------+
	#¦             ¦ folder ¦ folder, sub-folders ¦ folder and	     ¦ folder and    ¦ sub-folders      ¦ sub-folders      ¦ files         ¦
	#¦			   ¦ only   ¦ and files		      ¦ sub-folders      ¦ files         ¦ and files        ¦			       ¦               ¦
	#¦-------------+--------+--------------------------------------------------------------------------------------------------------------¦
	#¦ Propagation ¦ none   ¦ none                ¦ none             ¦ none          ¦ InheritOnly      ¦ InheritOnly      ¦ InheritOnly   ¦
	#¦-------------+--------+--------------------------------------------------------------------------------------------------------------¦
	#¦ Inheritance ¦ none   ¦ ContainerInherit    ¦ ContainerInherit ¦ ObjectInherit ¦ ContainerInherit ¦ ContainerInherit ¦ ObjectInherit ¦
	#¦             ¦        ¦ ObjectInherit       ¦                  ¦               ¦ ObjectInherit    ¦                  ¦               ¦
	#+-------------------------------------------------------------------------------------------------------------------------------------+

	#Local system:
	$bRecursive = $True;
	$bInheritFromParentAce = $True;
	$sSecID = "S-1-5-21-1801674531-2146617017-725345543-3908115";
	$sPath = "C:\SRM_Apps_N_Tools\PS-Scripts\Testing";

	#FileServer system:
	$bRecursive = $True;
	$bInheritFromParentAce = $True;
	$sSecID = "S-1-5-21-1801674531-2146617017-725345543-3908115";
	$strUserName = "redirect.test";
	$sPath = "\\NAEAPHILFS11\USER01\usr2\" + $strUserName;

	#Add the DefaultRequiredPrivileges to the current proccess token.
	#Without DefaultRequiredPrivileges you may not have access/permission to make changes.
	$Results = [FileOperations.Ntfs]::AddPrivilege([FileOperations.Ntfs]::DefaultRequiredPrivileges);
	if ($Results -eq $True){
		#Create the PermissionSet to add.
		#Use one of the DefaultPermissionSets, rather than setting all the ACE fields individually.
			#$PermSetAdmin = [FileOperations.Ntfs]::DefaultPermissionSets["FullControl"];
			#$PermSetUser = [FileOperations.Ntfs]::DefaultPermissionSets["Modify"];
			#$PermSetReader = [FileOperations.Ntfs]::DefaultPermissionSets["ReadExecute"];
		$PermSet = [FileOperations.Ntfs]::DefaultPermissionSets["FullControl"];

		#Set the Sid/Trustee in our PermissionSet.
			#$PermSetUser.UserSid = "S-1-5-21-2362956667-2133453920-1756558540-1001";
			#$PermSetReader.UserSid = "S-1-5-18";
		$PermSet.UserSid = $sSecID
		#Or use one of the "Well Know" default SID's
			#https://msdn.microsoft.com/en-us/library/system.security.principal.wellknownsidtype(v=vs.110).aspx
			#$PermSetAdmin.wellKnown = [System.Security.Principal.WellKnownSidType]::BuiltinAdministratorsSid;
		$PermSet.wellKnown = [System.Security.Principal.WellKnownSidType]::AccountGuestSid;

		#Add the PermissionSet.
		[FileOperations.Ntfs]::SetPermissions($sPath, $bRecursive, $PermSet, $bInheritFromParentAce);
		#No output from the results

		#Stop inheriting from Parent folder
		[FileOperations.Ntfs]::RemoveAllInheritedPermissions($sPath, $bRecursive);
		#No output from the results

		#A routine that will set the default NMCI Home Folder Permissions.  (The above 2 steps).
		#("User" [by SID] and "System Admin" both get "Full Control", remove all others, NO inherit from Parent.)
		[FileOperations.Ntfs]::SetDefaultHomeFolderPermissions($sPath, $sSecID);
		#No output from the results

		#Remove Permissions.
		[FileOperations.Ntfs]::RemovePermission($sPath, $bRecursive, $sSecID);
		#No output from the results
	}else{
		#DefaultRequiredPrivileges could NOT be added to the current proccess token.
		#You may not have enough permissions to make changes.
	}
}

function SetOwner_EXAMPLE_USAGE{

	$strSID = "S-1-5-21-1801674531-2146617017-725345543-3908115";
	$strUserName = "redirect.test";
	$strDirectory = "\\NAEAPHILFS11\USER01\usr2\" + $strUserName;

	$objItem = Get-Item -LiteralPath $strDirectory -Force -ErrorAction Stop;			#force is necessary to get hidden files/folders
	if ($objItem.PSIsContainer){
		#Directory
		$objACL = New-Object System.Security.AccessControl.DirectorySecurity;
	}else{
		#File
		$objACL = New-Object System.Security.AccessControl.FileSecurity;
	}
	#$objACL.SetOwner([System.Security.Principal.NTAccount]"BUILTIN\Administrators");
	#$objACL.SetOwner([System.Security.Principal.NTAccount]$strUserName);
	$objACL.SetOwner($strSID);
	#$objACL.SetOwner([System.Security.Principal.NTAccount]$strSID);

	#Merge the proposed changes (new owner) into the folder's ACL
	$objItem.SetAccessControl($objACL);
	if ($Error){
		$strMessage = "  Error making user '$strUserName' Owner of the Home Directory.`r`n" + $Error;
		$strMessage = "`r`n" + ("-" * 100) + "`r`n" + $strMessage + "`r`n`r`n";
	}else{
		$strMessage = "  Made user Owner of the Home Directory.`r`n";
	}
	Write-Host $strMessage;

}

$WinShare_CS = @"
using System;
using System.Runtime.InteropServices;

namespace FileOperations
{
    public static class WinShare
    {
        #region "Constants"

        private const uint StypeDisktree = 0;
        private const uint SecurityDescriptorRevision = 1;
        private const uint SddlRevision1 = 1;

        #endregion

        #region "Enums"

        public enum NetApiStatus
        {
            NerrSuccess = 0,
            ErrorAccessDenied = 5,
            ErrorInvalidParameter = 87,
            ErrorInvalidName = 123,
            ErrorInvalidLevel = 124,
            NerrUnknownDevDir = 2116,
            NerrRedirectedPath = 2117,
            NerrDuplicateShare = 2118,
            NerrBufTooSmall = 2123
        }

        private enum AccessMode : uint
        {
            NotUsedAccess = 0,
            GrantAccess = 1,
            SetAccess = 2,
            DenyAccess = 3,
            RevokeAccess = 4,
            SetAuditSuccess = 5,
            SetAuditFailure = 6
        }

        private enum MultipleTrusteeOperation : uint
        {
            NoMultipleTrustee = 0,
            TrusteeIsImpersonate = 1
        }

        private enum TrusteeForm : uint
        {
            TrusteeIsSid = 0,
            TrusteeIsName = 1,
            TrusteeBadForm = 2,
            TrusteeIsObjectsAndSid = 3,
            TrusteeIsObjectsAndName = 4
        }

        private enum TrusteeType : uint
        {
            TrusteeIsUnknown = 0,
            TrusteeIsUser = 1,
            TrusteeIsGroup = 2,
            TrusteeIsDomain = 3,
            TrusteeIsAlias = 4,
            TrusteeIsWellKnownGroup = 5,
            TrusteeIsDeleted = 6,
            TrusteeIsInvalid = 7,
            TrusteeIsComputer = 8
        }

        #endregion

        #region "Structures"

        [StructLayout(LayoutKind.Sequential)]
        private struct ExplicitAccess
        {
            private readonly uint grfAccessPermissions;
            private readonly AccessMode grfAccessMode;
            private readonly uint grfInheritance;
            private readonly Trustee Trustee;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct ShareInfo1501
        {
            public Int32 shi1501_reserved;
            public IntPtr shi1501_security_descriptor;
        };

        [StructLayout(LayoutKind.Sequential)]
        private struct SecurityDescriptor
        {
            public Byte Revision;
            private readonly Byte Sbz1;
            private readonly ushort Control;
            private readonly IntPtr Owner;
            private readonly IntPtr Group;
            private readonly IntPtr Sacl;
            private readonly IntPtr Dacl;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct ShareInfo502
        {
            [MarshalAs(UnmanagedType.LPWStr)] public String shi502_netname;
            public uint shi502_type;
            [MarshalAs(UnmanagedType.LPWStr)] public String shi502_remark;
            public int shi502_permissions;
            public int shi502_max_uses;
            public int shi502_current_uses;
            [MarshalAs(UnmanagedType.LPWStr)] public String shi502_path;
            [MarshalAs(UnmanagedType.LPWStr)] public String shi502_passwd;
            public int shi502_reserved;
            public IntPtr shi502_security_descriptor;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct Trustee
        {
            private readonly uint pMultipleTrustee;
            private readonly MultipleTrusteeOperation MultipleTrusteeOperation;
            private readonly TrusteeForm TrusteeForm;
            private readonly TrusteeType TrusteeType;
            [MarshalAs(UnmanagedType.LPTStr)] private readonly String ptstrName;
        }

        #endregion

        #region "Native Methods"

        [DllImport("netapi32.dll", SetLastError = true)]
        private static extern NetApiStatus NetShareAdd(
            [MarshalAs(UnmanagedType.LPWStr)] string strServer,
            Int32 dwLevel,
            ref ShareInfo502 buf,
            out uint parm_err
            );

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern uint InitializeSecurityDescriptor(
            out SecurityDescriptor securityDescriptor,
            uint dwRevision
            );

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern uint SetEntriesInAcl(
            int cCountOfExplicitEntries,
            ref ExplicitAccess pListOfExplicitEntries,
            IntPtr oldAcl,
            out IntPtr newAcl
            );

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern uint SetSecurityDescriptorDacl(
            ref SecurityDescriptor sd,
            bool daclPresent,
            IntPtr dacl,
            bool daclDefaulted
            );

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern uint IsValidSid(
            SecurityDescriptor pSecurityDescriptor
            );

        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool ConvertStringSecurityDescriptorToSecurityDescriptor(
            string stringSecurityDescriptor,
            uint stringSdRevision,
            out IntPtr securityDescriptor,
            out uint securityDescriptorSize
            );

        [DllImport("Netapi32", CharSet=CharSet.Auto)]
        static extern NetApiStatus NetShareGetInfo(
            [MarshalAs(UnmanagedType.LPWStr)] string servername,
            [MarshalAs(UnmanagedType.LPWStr)] string netname,
            int level,
            ref IntPtr bufptr
            );

        [DllImport("Netapi32.dll", SetLastError = true)]
        private static extern Int32 NetShareSetInfo(
            [MarshalAs(UnmanagedType.LPWStr)] string servername,
            [MarshalAs(UnmanagedType.LPWStr)] string netname,
            Int32 level,
            IntPtr bufptr,
            out Int32 parmErr
            );

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr LocalFree(
            IntPtr hMem
            );

        [DllImport("netapi32.dll", SetLastError = true)]
        private static extern NetApiStatus NetShareDel(
            [MarshalAs(UnmanagedType.LPWStr)] string strServer,
            [MarshalAs(UnmanagedType.LPWStr)] string strNetName,
            Int32 reserved //must be 0
            );

        [DllImport("Netapi32", CharSet=CharSet.Auto)]
        static extern int NetApiBufferFree(
            IntPtr Buffer
            );

        #endregion

        #region "Managed Methods"

        /// <summary>
        ///     Shares an existing folder on the local computer or on a remote computer
        /// </summary>
        /// <param name="computerName">The remote computer name to create the share on</param>
        /// <param name="localPath">
        ///     The local path to the folder to be shared. If creating share on a remote computer then
        ///     the path must be local to the remote computer. Do not use UNC paths
        /// </param>
        /// <param name="shareName">The name for the share</param>
        /// <param name="shareComment">An optional comment/description for the share</param>
        public static NetApiStatus Add(
            String computerName,
            String localPath,
            String shareName,
            String shareComment )
        {
            //Argument validation
            if (String.IsNullOrEmpty(shareName) | String.IsNullOrEmpty(localPath))
            {
                throw new ArgumentException(
                    "Invalid argument specified - ShareName, LocalPath and ComputerName arguments must not be empty");
            }

            //This pointer will hold the full ACL (access control list) once the loop below has completed
            var aclPtr = IntPtr.Zero;

            //Create a SECURITY_DESCRIPTOR structure and set the Revision number
            SecurityDescriptor secDesc;
            secDesc.Revision = (byte) SecurityDescriptorRevision;
            //Initialise the SECURITY_DESCRIPTOR instance - returns 0 if an error was encountered
            var decriptorInitResult = InitializeSecurityDescriptor(out secDesc, SecurityDescriptorRevision);
            if (decriptorInitResult == 0)
            {
                throw new ApplicationException(
                    "An error was encountered during the call to the InitializeSecurityDescriptor API. The share has not been created. You may be able to get more information on the error by checking the Err.LastDllError property");
            }
            //Add the ACL to the SECURITY_DESCRIPTOR
            var setSecurityResult = SetSecurityDescriptorDacl(ref secDesc, true, aclPtr, false);
            if (setSecurityResult == 0)
            {
                throw new ApplicationException(
                    "An error was encountered during the call to the SetSecurityDescriptorDacl API. The share has not been created. You may be able to get more information on the error by checking the Err.LastDllError property");
            }
            //Check to make sure the SECURITY_DESCRIPTOR is valid
            if (IsValidSid(secDesc) == 0)
            {
                throw new ApplicationException(
                    "No errors were reported from previous API calls but the security descriptor is not valid. The share has not been created.");
            }
            //Create a pointer for the SECURITY_DESCRIPTOR so that we can pass this in to the SHARE_INFO_502 structure
            var secDescPtr = Marshal.AllocCoTaskMem(Marshal.SizeOf(secDesc));
            Marshal.StructureToPtr(secDesc, secDescPtr, false);
            //Create and populate the SHARE_INFO_502 structure that specifies all of the share settings
            ShareInfo502 shareInfo;
            shareInfo.shi502_netname = shareName;
            shareInfo.shi502_type = StypeDisktree;
            shareInfo.shi502_remark = shareComment;
            shareInfo.shi502_permissions = 0;
            shareInfo.shi502_max_uses = -1;
            shareInfo.shi502_current_uses = 0;
            shareInfo.shi502_path = localPath;
            shareInfo.shi502_passwd = null;
            shareInfo.shi502_reserved = 0;
            shareInfo.shi502_security_descriptor = secDescPtr;
            //Call the NetShareAdd API to create the share
            uint error = 0;
            NetApiStatus result = NetShareAdd(computerName, 502, ref shareInfo, out error);
            //Clean up and return the result of NetShareAdd
            Marshal.FreeCoTaskMem(secDescPtr);
            return result;
        }

        public static bool SetDefaultSharePermissions(
            String computerName,
            String shareName)
        {
            const string sSecurityDescriptor = "D:(A;;FA;;;AU)";
            // REFERENCE FOR THE AFORE DACL STRING (http://www.netid.washington.edu/documentation/domains/sddl.aspx)
            IntPtr lpSecurityDescriptor;
            uint securityDescriptorSize;

            if (
                !ConvertStringSecurityDescriptorToSecurityDescriptor(sSecurityDescriptor, SddlRevision1,
                    out lpSecurityDescriptor, out securityDescriptorSize))
            {
                //Console.WriteLine("ConvertStringSecurityDescriptorToSecurityDescriptor failed with {0}",
                //    Marshal.GetLastWin32Error());
                return false;
            }
            //Console.WriteLine("ConvertStringSecurityDescriptorToSecurityDescriptor SUCCESS, size = {0}",
            //    securityDescriptorSize);

            Int32 paramErr;

            var si1501 = new ShareInfo1501();

            var buffer = Marshal.AllocHGlobal(Marshal.SizeOf(si1501));

            si1501.shi1501_security_descriptor = lpSecurityDescriptor;

            Marshal.StructureToPtr(si1501, buffer, false);

            var nas = NetShareSetInfo(computerName, shareName, 1501, buffer, out paramErr);

            if (lpSecurityDescriptor != IntPtr.Zero)
                LocalFree(lpSecurityDescriptor);

            if (nas != 0)
            {
                //Console.WriteLine("NetShareSetInfo failed with: {0}", nas);
                return false;
            }
            //Console.WriteLine("NetShareSetInfo SUCCESS");
            return true;
        }

        public static NetApiStatus Delete(
            String computerName,
            String shareName)
        {
            var result = NetShareDel(computerName, shareName, 0);
            return result;
        }

        //public static string NetShareGetPath(
        public static string GetShare(
			String computerName, 
			String shareName)
        {
			// <summary> Retrieves the local path for the given server and share name. </summary>
				// 0				No errors encountered.
				// 5				The user has insufficient privilege for this operation.
				// 8				Not enough memory
				// 65				Network access is denied.
				// 87				Invalid parameter specified.
				// 53				The network path was not found.
				// 123				Invalid name
				// 124				Invalid level parameter.
				// 234				More data available, buffer too small.
				// 2100
				// 2102				Device driver not installed.
				// 2106				This operation can be performed only on a server.
				// 2114				Server service not installed.
				// 2116				NerrUnknownDevDir
				// 2117				NerrRedirectedPath
				// 2118				NerrDuplicateShare
				// 2123				Buffer too small for fixed-length data.
				// 2127				Error encountered while remotely.  executing function
				// 2138				The Workstation service is not started.
				// 2141				The server is not configured for this transaction;  IPC$ is not shared.
				// 2310				Sharename not found.
				// (Result + 210)	Sharename not found.
				// 2351				Invalid computername specified.

            string sharePath = null;
            IntPtr ptr = IntPtr.Zero;

            string sFullShareInfo = null;
            string shareComment = null;
            int shareMaxCon = 0;
            int shareCurCon = 0;

            if (String.IsNullOrEmpty(computerName) | String.IsNullOrEmpty(shareName))
            {
                throw new ArgumentException(
                    "Invalid argument specified - computerName and shareName arguments must not be empty");
            }

            //NetApiStatus errCode = NetShareGetInfo(computerName, shareName, 2, ref ptr);
            NetApiStatus errCode = NetShareGetInfo(computerName, shareName, 502, ref ptr);
            if (errCode == NetApiStatus.NerrSuccess)
            {
                ShareInfo502 shareInfo = (ShareInfo502)Marshal.PtrToStructure(ptr, typeof(ShareInfo502));
                sharePath = shareInfo.shi502_path;
				//shareName = shareInfo.shi502_netname;
				//shareType = shareInfo.shi502_type;
				shareComment = shareInfo.shi502_remark;
				//sharePerms = shareInfo.shi502_permissions;
				shareMaxCon = shareInfo.shi502_max_uses;
				shareCurCon = shareInfo.shi502_current_uses;
				//sharePass = shareInfo.shi502_passwd;
				//shareReserved = shareInfo.shi502_reserved;
				//shareDescriptor = shareInfo.shi502_security_descriptor;

                NetApiBufferFree(ptr);

				sFullShareInfo = "ShareName:    " + "\\\\" + computerName + "\\" + shareName + "\r\n" + "SharePath:    " + sharePath + "\r\n" + "ShareComment: " + shareComment + "\r\n" + "ShareMaxCon:  " + shareMaxCon + "\r\n" + "ShareCurCon:  " + shareCurCon + "\r\n";
            }
            else
			{
                //return(errCode);
                return(errCode.ToString());
			}

            return sFullShareInfo;
        }

        #endregion
    }
}
"@


$Ntfs_CS = @"
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;

namespace FileOperations
{
    public static class Ntfs
    {
        #region "Globals"
        public static readonly Dictionary<string,PermissionSet> DefaultPermissionSets = new Dictionary<string,PermissionSet>() {
            {
                "FullControl",
                new PermissionSet {
                    Rights = FileSystemRights.FullControl,
                    InheritanceFlags = InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                    PropagationFlags = PropagationFlags.None,
                    AccessControlType = AccessControlType.Allow 
                }
            },
            {
                "ReadExecute",
                new PermissionSet {
                    Rights = FileSystemRights.ReadAndExecute,
                    InheritanceFlags = InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                    PropagationFlags = PropagationFlags.None,
                    AccessControlType = AccessControlType.Allow
                }
            },
            {
                "Modify",
                new PermissionSet {
                    Rights = FileSystemRights.Modify,
                    InheritanceFlags = InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                    PropagationFlags = PropagationFlags.None,
                    AccessControlType = AccessControlType.Allow
                }
            }
        };

        
        #endregion

        #region "Constants"
        internal const int SePrivilegeDisabled = 0x00000000;
        internal const int SePrivilegeEnabled = 0x00000002;
        internal const int TokenQuery = 0x00000008;
        internal const int TokenAdjustPrivileges = 0x00000020;
        public static readonly string[] DefaultRequiredPrivileges = new string[4] { "SeRestorePrivilege", "SeBackupPrivilege", "SeTakeOwnershipPrivilege", "SeSecurityPrivilege" };
        private const int SW_SHOW = 5;
        private const uint SEE_MASK_INVOKEIDLIST = 12;

        #endregion

        #region "Structures"
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        internal struct TokPriv1Luid
        {
            public int Count;
            public long Luid;
            public int Attr;
        }

        public struct PermissionSet
        {
            public string UserSid;
            public WellKnownSidType wellKnown;
            public FileSystemRights Rights; 
            public InheritanceFlags InheritanceFlags;
            public PropagationFlags PropagationFlags;
            public AccessControlType AccessControlType;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct SHELLEXECUTEINFO
        {
            public int cbSize;
            public uint fMask;
            public IntPtr hwnd;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string lpVerb;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string lpFile;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string lpParameters;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string lpDirectory;
            public int nShow;
            public IntPtr hInstApp;
            public IntPtr lpIDList;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string lpClass;
            public IntPtr hkeyClass;
            public uint dwHotKey;
            public IntPtr hIcon;
            public IntPtr hProcess;
        }

        #endregion

        #region "Native Methods"
        
        [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
        internal static extern bool AdjustTokenPrivileges(IntPtr htok, bool disall, ref TokPriv1Luid newst, int len, IntPtr prev, IntPtr relen);
        
        [DllImport("kernel32.dll", ExactSpelling = true)]
        internal static extern IntPtr GetCurrentProcess();
        
        [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
        internal static extern bool OpenProcessToken(IntPtr h, int acc, ref IntPtr phtok);
        
        [DllImport("advapi32.dll", SetLastError = true)]
        internal static extern bool LookupPrivilegeValue(string host, string name,ref long pluid);

        [DllImport("Shell32.dll", CharSet = CharSet.Auto)]
        static extern bool ShellExecuteEx(ref SHELLEXECUTEINFO lpExecInfo);

        #endregion

        #region "Managed Methods"
        /// <summary>
        /// Used to add Privilege(s) to the current executing Process (E.G.; .exe if dll, powershell.exe if ps1)
        /// For Example: var retVal = Ntfs.AddPrivileges(Ntfs.DefaultRequiredPrivileges);
        /// </summary>
        /// <param name="privileges"></param>
        /// <returns></returns>
        public static bool AddPrivilege(string privilege)
        {
            TokPriv1Luid tp;
            var hproc = GetCurrentProcess();
            var htok = IntPtr.Zero;
            OpenProcessToken(hproc, TokenAdjustPrivileges | TokenQuery, ref htok);
            tp.Count = 1;
            tp.Luid = 0;
            tp.Attr = SePrivilegeEnabled;
            LookupPrivilegeValue(null, privilege, ref tp.Luid);
            var retVal = AdjustTokenPrivileges(htok, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
            return retVal;
        }
        /// <summary>
        /// Used to add Privilege(s) to the current executing Process (E.G.; .exe if dll, powershell.exe if ps1)
        /// For Example: var retVal = Ntfs.AddPrivileges(Ntfs.DefaultRequiredPrivileges);
        /// </summary>
        /// <param name="privileges"></param>
        /// <returns></returns>
        public static bool AddPrivilege(IEnumerable<string> privileges)
        {
            var retVal = true;
            foreach (var privilege in privileges)
            {
                TokPriv1Luid tp;
                var hproc = GetCurrentProcess();
                var htok = IntPtr.Zero;
                retVal &= OpenProcessToken(hproc, TokenAdjustPrivileges | TokenQuery, ref htok);
                tp.Count = 1;
                tp.Luid = 0;
                tp.Attr = SePrivilegeEnabled;
                retVal &= LookupPrivilegeValue(null, privilege, ref tp.Luid);
                retVal &= AdjustTokenPrivileges(htok, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
            }
            return retVal;
        }
        /// <summary>
        /// Used to remove Privilege(s) to the current Process
        /// For Example: var retVal = Ntfs.RemovePrivileges(Ntfs.DefaultRequirePrivileges);
        /// </summary>
        /// <param name="privileges"></param>
        /// <returns></returns>
        private static bool RemovePrivilege(string privilege)
        {
            try
            {
                TokPriv1Luid tp;
                var hproc = GetCurrentProcess();
                var htok = IntPtr.Zero;
                var retVal = OpenProcessToken(hproc, TokenAdjustPrivileges | TokenQuery, ref htok);
                tp.Count = 1;
                tp.Luid = 0;
                tp.Attr = SePrivilegeDisabled;
                retVal &= LookupPrivilegeValue(null, privilege, ref tp.Luid);
                retVal &= AdjustTokenPrivileges(htok, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
                return retVal;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>
        /// Used to remove Privilege(s) to the current Process
        /// For Example: var retVal = Ntfs.RemovePrivileges(Ntfs.DefaultRequirePrivileges);
        /// </summary>
        /// <param name="privileges"></param>
        /// <returns></returns>
        private static bool RemovePrivilege(IEnumerable<string> privileges)
        {
            var retVal = true;
            foreach (var privilege in privileges)
            {
                TokPriv1Luid tp;
                var hproc = GetCurrentProcess();
                var htok = IntPtr.Zero;
                retVal = OpenProcessToken(hproc, TokenAdjustPrivileges | TokenQuery, ref htok);
                tp.Count = 1;
                tp.Luid = 0;
                tp.Attr = SePrivilegeDisabled;
                retVal &= LookupPrivilegeValue(null, privilege, ref tp.Luid);
                retVal &= AdjustTokenPrivileges(htok, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
            }
            return retVal;
        }
        /// <summary>
        /// Add (ACE)AccessControlEntry(s) for the provided PermissionSet (can be manually formed)
        /// </summary>
        /// <param name="path"></param>
        /// <param name="recursive"></param>
        /// <param name="userSid"></param>
        /// <param name="wellKnown"></param>
        /// <param name="fsRights"></param>
        /// <param name="iFlags"></param>
        /// <param name="pFlags"></param>
        /// <param name="acType"></param>
        /// <param name="inheritFromParentAce"></param>
        public static void SetPermissions(string path, bool recursive, string userSid, WellKnownSidType wellKnown, FileSystemRights fsRights, InheritanceFlags iFlags, PropagationFlags pFlags, AccessControlType acType, bool inheritFromParentAce)
        {
            IdentityReference account = null;
            account = userSid != null ? new SecurityIdentifier(userSid) : new SecurityIdentifier(wellKnown, null);
            if (File.GetAttributes(path).HasFlag(FileAttributes.Directory))
            {
                var di = new DirectoryInfo(path);
                var ds = di.GetAccessControl();
                var fsaRule = new FileSystemAccessRule(account, fsRights, iFlags, pFlags, acType);
                ds.AddAccessRule(fsaRule);
                if (!inheritFromParentAce) ds.SetAccessRuleProtection(true, false); else ds.SetAccessRuleProtection(false, true);
                di.SetAccessControl(ds);
            }
            else
            {
                var fi = new FileInfo(path);
                var fs = fi.GetAccessControl();
                var fsaRule = new FileSystemAccessRule(account, fsRights, iFlags, pFlags, acType);
                fs.AddAccessRule(fsaRule);
                if (!inheritFromParentAce) fs.SetAccessRuleProtection(true, false); else fs.SetAccessRuleProtection(false, true);
                fi.SetAccessControl(fs);
            }
            if (!(recursive & File.GetAttributes(path).HasFlag(FileAttributes.Directory))) return;
            var rdi = new DirectoryInfo(path);
            var colDir = rdi.EnumerateDirectories();
            foreach (var subDir in colDir)
            {
                SetPermissions(subDir.FullName, recursive, userSid, wellKnown, fsRights, iFlags, pFlags, acType, inheritFromParentAce);
            }
        }
        /// <summary>
        /// Add (ACE)AccessControlEntry(s) for the provided PermissionSet (can be manually formed)
        /// </summary>
        /// <param name="path"></param>
        /// <param name="recursive"></param>
        /// <param name="permissions"></param>
        /// <param name="inheritFromParentAce"></param>
        public static void SetPermissions(string path, bool recursive, PermissionSet permissions, bool inheritFromParentAce)
        {
            IdentityReference account = null;
            account = permissions.UserSid != null ? new SecurityIdentifier(permissions.UserSid) : new SecurityIdentifier(permissions.wellKnown, null);
            if (File.GetAttributes(path).HasFlag(FileAttributes.Directory))
            { 
                var di = new DirectoryInfo(path);
                var ds = di.GetAccessControl();
                var fsaRule = new FileSystemAccessRule(account, permissions.Rights, permissions.InheritanceFlags, permissions.PropagationFlags, permissions.AccessControlType);
                ds.AddAccessRule(fsaRule);
                if (!inheritFromParentAce) ds.SetAuditRuleProtection(true, false); else ds.SetAccessRuleProtection(false, true);
                di.SetAccessControl(ds);
            }
            else
            {
                var fi = new FileInfo(path);
                var fs = fi.GetAccessControl();
                var fsaRule = new FileSystemAccessRule(account, permissions.Rights, permissions.InheritanceFlags, permissions.PropagationFlags, permissions.AccessControlType);
                fs.AddAccessRule(fsaRule);
                if (!inheritFromParentAce) fs.SetAccessRuleProtection(true, false); else fs.SetAccessRuleProtection(false, true);
                fi.SetAccessControl(fs);
            }
            if (!(recursive & File.GetAttributes(path).HasFlag(FileAttributes.Directory))) return;
            var rdi = new DirectoryInfo(path);
            var colDir = rdi.EnumerateDirectories();
            foreach (var subDir in colDir)
            {
                SetPermissions(subDir.FullName, recursive, permissions, inheritFromParentAce);
            }
        }
        /// <summary>
        /// Add (ACE)AccessControlEntry(s) for the provided PermissionSet (can be manually formed)
        /// </summary>
        /// <param name="path"></param>
        /// <param name="recursive"></param>
        /// <param name="permissions"></param>
        /// <param name="inheritFromParentAce"></param>
        public static void SetPermissions(string path, bool recursive, PermissionSet[] permissions, bool inheritFromParentAce)
        {
            var permCollection = new List<FileSystemAccessRule>();
            foreach (var permission in permissions)
            {
                IdentityReference account = null;
                account = permission.UserSid != null ? new SecurityIdentifier(permission.UserSid) : new SecurityIdentifier(permission.wellKnown, null);
                var fsaRule = new FileSystemAccessRule(account, permission.Rights, permission.InheritanceFlags, permission.PropagationFlags, permission.AccessControlType);
                permCollection.Add(fsaRule);
            }
            if (File.GetAttributes(path).HasFlag(FileAttributes.Directory))
            {
                var di = new DirectoryInfo(path);
                var ds = di.GetAccessControl();
                if (!inheritFromParentAce) ds.SetAccessRuleProtection(true, false); else ds.SetAccessRuleProtection(false, true);
                foreach (var fsaRule in permCollection)
                {
                    ds.SetAccessRule(fsaRule);
                }
                di.SetAccessControl(ds);
            }
            else
            {
                var fi = new FileInfo(path);
                var fs = fi.GetAccessControl();
                if (!inheritFromParentAce) fs.SetAccessRuleProtection(true, false); else fs.SetAccessRuleProtection(false, true);
                foreach (var fsaRule in permCollection)
                {
                    fs.AddAccessRule(fsaRule);
                }
                fi.SetAccessControl(fs);
            }
            if (!(recursive & File.GetAttributes(path).HasFlag(FileAttributes.Directory))) return;
            var rdi = new DirectoryInfo(path);
            var colDir = rdi.EnumerateDirectories();
            foreach (var subDir in colDir)
            {
                SetPermissions(subDir.FullName, recursive, permissions, inheritFromParentAce);
            }
        }
        /// <summary>
        /// Set NTFS Permissions as follows:
        ///  Provided User's sID:  Full Control
        ///  BuiltinAdmins:  Full Control
        /// </summary>
        /// <param name="path"></param>
        /// <param name="sid"></param>
        public static void SetDefaultHomeFolderPermissions(string path, string sid)
        {
            var userPerms = Ntfs.DefaultPermissionSets["FullControl"];
            userPerms.UserSid = sid;
            var adminPerms = Ntfs.DefaultPermissionSets["FullControl"];
            adminPerms.wellKnown = WellKnownSidType.BuiltinAdministratorsSid;
            SetPermissions(path, false, new PermissionSet[] { userPerms, adminPerms }, false);
        }
        /// <summary>
        /// Remove all AccessControlList (ACL) entries for the provided sid(s)
        /// </summary>
        /// <param name="path"></param>
        /// <param name="recursive">Operates on all child directories</param>
        /// <param name="sid"></param>
        public static void RemovePermission(string path, bool recursive, string sid)
        {
            if (File.GetAttributes(path).HasFlag(FileAttributes.Directory))
            {
                var di = new DirectoryInfo(path);
                var ds = di.GetAccessControl();
                ds.PurgeAccessRules(new SecurityIdentifier(sid));
                di.SetAccessControl(ds);
            }
            else
            {
                var fi = new FileInfo(path);
                var fs = fi.GetAccessControl();
                var acls = fs.GetAccessRules(true, false, typeof(SecurityIdentifier));
                fs.PurgeAccessRules(new SecurityIdentifier(sid));
                fi.SetAccessControl(fs);
            }
            if (!(recursive & File.GetAttributes(path).HasFlag(FileAttributes.Directory))) return;
            var rdi = new DirectoryInfo(path);
            var colDir = rdi.EnumerateDirectories();
            foreach (var subDir in colDir)
            {
                RemovePermission(subDir.FullName, recursive, sid);
            }
        }
        /// <summary>
        /// Remove all AccessControlList (ACL) entries for the provided sid(s)
        /// </summary>
        /// <param name="path"></param>
        /// <param name="recursive"></param>
        /// <param name="sid"></param>
        public static void RemovePermission(string path, bool recursive, string[] sid)
        {
            if (File.GetAttributes(path).HasFlag(FileAttributes.Directory))
            {
                var di = new DirectoryInfo(path);
                var ds = di.GetAccessControl();
                var acls = ds.GetAccessRules(true, false, typeof(SecurityIdentifier));
                foreach (var fsaRule in acls.Cast<FileSystemAccessRule>().Where(fsaRule => sid.Contains(fsaRule.IdentityReference.Value)))
                {
                    ds.RemoveAccessRule(fsaRule);
                }
                di.SetAccessControl(ds);
            }
            else
            {
                var fi = new FileInfo(path);
                var fs = fi.GetAccessControl();
                var acls = fs.GetAccessRules(true, false, typeof(SecurityIdentifier));
                foreach (var fsaRule in acls.Cast<FileSystemAccessRule>().Where(fsaRule => sid.Contains(fsaRule.IdentityReference.Value)))
                {
                    fs.RemoveAccessRule(fsaRule);
                }
                fi.SetAccessControl(fs);
            }
            if (!(recursive & File.GetAttributes(path).HasFlag(FileAttributes.Directory))) return;
            var rdi = new DirectoryInfo(path);
            var colDir = rdi.EnumerateDirectories();
            foreach (var subDir in colDir)
            {
                RemovePermission(subDir.FullName, recursive, sid);
            }
        }
        /// <summary>
        /// Removes all inherited AccessControlEntries (ACE) entries from the AccessControlLists (ACL) of the given path.
        /// </summary>
        /// <param name="path"></param>
        /// <param name="recursive"></param>
        public static void RemoveAllInheritedPermissions(string path, bool recursive)
        {
            if (File.GetAttributes(path).HasFlag(FileAttributes.Directory))
            {
                var di = new DirectoryInfo(path);
                var ds = di.GetAccessControl();
                ds.SetAccessRuleProtection(true, false);
                di.SetAccessControl(ds);
            }
            else
            {
                var fi = new FileInfo(path);
                var fs = fi.GetAccessControl();
                fs.SetAccessRuleProtection(true, false);
                fi.SetAccessControl(fs);
            }
        }

        public static bool ShowFileProperties(string FileName)
        {
            SHELLEXECUTEINFO info = new SHELLEXECUTEINFO();
            info.cbSize = System.Runtime.InteropServices.Marshal.SizeOf(info);
            info.lpVerb = "properties";
            info.lpFile = FileName;
            info.nShow = SW_SHOW;
            info.fMask = SEE_MASK_INVOKEIDLIST;
            return ShellExecuteEx(ref info);
        }

        #endregion
    }
}
"@

Add-Type -TypeDefinition $WinShare_CS -IgnoreWarnings;
Add-Type -TypeDefinition $Ntfs_CS -IgnoreWarnings -ReferencedAssemblies @("System.Core");
