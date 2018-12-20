' //=========================================================================
' //*************************************************************************
' // File:       CE16563_800G2SFFUPDATE.vbs
' // Created by: Sam James
' // Modified by: Brian Broda
' // Solution:   Upgrades BIOS on HP 800 G2 SFF
' //
' //
' // Change History:
' // 1.0    - 8/16/2016 - Initial Creation (based off of CE15667 BIOS Upgrade template)
' // 1.1    - 9/19/2016 - Modified to only upgrade video driver and BIOS for BIOS upgrade deployment
' // 1.2    - 10/7/2016 - Modified for all drivers and BIOS for 840 G3 upgrade deployment
' // 1.3    - 10/11/2016 - Modified for use with EliteDesk 800 G2
' //*************************************************************************
' //=========================================================================



Option Explicit
'*---------------------------------------------------='
'* Redirects script to run in native form if needed. ='
'*---------------------------------------------------=' 
If InStr(WScript.FullName,"SysWOW64") > 0 Then
	DIM sw_Argument     : If WScript.Arguments.Count <> 0 Then sw_Argument = WScript.Arguments.Item(0)
	DIM sw_i            : If WScript.Arguments.Count  > 1 Then Do while WScript.Arguments.Count > sw_i + 1 : sw_i = sw_i + 1 : sw_Argument = sw_Argument & " " & WScript.Arguments.Item(sw_i) : Loop
	DIM sw_wshShell     : Set sw_wshShell = WScript.CreateObject("WScript.Shell")
	DIM sw_ENVWindows   : sw_ENVWindows   = sw_wshShell.ExpandEnvironmentStrings("%windir%")
	DIM sw_strScptAgent : sw_strScptAgent = "wscript" : If instr(WScript.FullName,"\cscript.exe") > 0 Then sw_strScptAgent = "cscript"
	DIM sw_SwitchSilent : sw_SwitchSilent = ""        : If Not WScript.Interactive Then sw_SwitchSilent = "//B "
	DIM sw_Return       : sw_Return       = sw_wshShell.Run(sw_ENVWindows & "\sysnative\" & sw_strScptAgent & ".exe " & sw_SwitchSilent & """" & WScript.ScriptFullName & """ " & sw_Argument, 0, True)
	wscript.quit sw_Return
End If

'*------------------='
'* Global Variables ='
'*------------------='
DIM GBL_wshShell        : Set GBL_wshShell      = CreateObject("WScript.Shell")
DIM GBL_wshFileSys      : Set GBL_wshFileSys    = CreateObject("Scripting.FileSystemObject")
DIM GBL_objRegistry     : Set GBL_objRegistry   = GetObject("winmgmts:root\default:StdRegProv")
DIM GBL_objWMIService   : Set GBL_objWMIService = GetObject("winmgmts:\\.\root\cimv2")
DIM GBL_ObjArgs         : Set GBL_ObjArgs       = WScript.Arguments
DIM GBL_WshNetwork      : Set GBL_WshNetwork    = WScript.CreateObject("WScript.Network")
'----------------------------------------------------------------------------------------------------------------------------------------------
Const VBS_VERSION        = 1.0
Const DEBUG_ON           = true
Const ENABLE_LOGGING     = True
Const HKEY_CURRENT_USER  = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS         = &H80000003
Const FOR_READING        = 1
Const FOR_WRITING        = 2
Const FOR_APPENDING      = 8
Const TRISTATETRUE       = -1  ' Opens the file as Unicode
Const TRISTATEFALSE      = 0   ' Opens the file as ASCII
Const TRISTATEUSEDEFAULT = -2  ' Opens the using the default system setting
Const ENH_NUM 			 = "CE16563-800G2-Update"
Const ENH_allClear		 = "CE16563-allSuccess"
Const ENH_BIOS 			 = "CE16563-BIOS-Success"
Const REG_VALFAIL        = "CE16563-BIOS-Retry"
Const ENH_DrvrFAILED	 = "CE16563-Driver-Failed"
Const ENH_DriverSuccess	 = "CE16563-Driver-Success"
const ACTIVE_SETUP_REG 	 = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Active Setup\Installed Components\CE16563\"
const APP_VERSION 				= "1,1"
Const RC_OK = 0
Const RC_FILENOTFOUND_ERROR              = 101
Const RC_SHORTCUTCREATE_ERROR            = 102
Const RC_INVALIDDOMAIN_ERROR             = 103
Const RC_FILEOBJECT_CREATE_ERROR         = 104
Const RC_SHELLOBJECT_CREATE_ERROR        = 105
Const RC_ARGUMENTOBJECT_CREATE_ERROR     = 106
Const RC_NETWORKOBJECT_CREATE_ERROR      = 107
Const RC_FOLDERCREATE_ERROR              = 108
Const RC_REGISTRY_ERROR                  = 109
Const RC_LOGFOLDER_ERROR                 = 110
Const RC_LOGFILE_ERROR                   = 111
Const RC_SHORTCUTDELETE_ERROR            = 112
'----------------------------------------------------------------------------------------------------------------------------------------------
DIM GBL_REG_UNINSTALL_KEY 					'* = Determined by Client OS via registry(X86 or AMD64)
DIM GBL_Name             : GBL_Name            = Left(Wscript.ScriptName, len(Wscript.ScriptName)-4)
DIM GBL_Log              : GBL_Log             = GBL_Name & ".LOG"
DIM GBL_LogFolder        : GBL_LogFolder       = "C:\Logs\CE16563-800G2Upgrade\"
DIM GBL_LogFile          : GBL_LogFile         = "C:\Logs\CE16563-800G2Upgrade\" & GBL_Log
DIM GBL_OSType           : GBL_OSType          = fnCheckOSType()
'DIM GBL_Env_EDWS         : GBL_Env_EDWS        = fnDeployToEDWS	            '* use if deploying to EDWS *'
DIM GBL_ScriptPath       : GBL_ScriptPath      = GBL_wshFileSys.GetFile(wscript.scriptfullname).ParentFolder & "\"
DIM GBL_FailSafeTask	 : GBL_FailSafeTask	   = "FailSafeBIOSCheck"
DIM GBL_MBamOffTask	  	 : GBL_MBamOffTask	   = "BIOSSafeMBamoff"
Dim GBL_strScriptName	 : GBL_strScriptName   = WScript.ScriptName
Dim REG_VALDOUBLE		 :REG_VALDOUBLE 	   = "CE16563_DOUBLEREBOOT"
DIM runLync
DIM strTaskName
DIM GBL_TempXMLFolder
DIM GBL_XML_Name
DIM GBL_TaskName
DIM g_strCurrentBIOSLevel
DIM strBiosCmd
DIM GBL_intBIOSVersion
DIM intReturn
DIM GBL_strBitPath
DIM GBL_strFnErrorCode
DIM sTaskReturn
DIM BiosVersionUpgradeCmd
DIM BiosVersionRollbackCmd
DIM UpgradedBiosVersion
DIM strClientModel
DIM X64BitHost
DIM colItems, objItem, strModel, pcmodel
DIM colBios, objBios
DIM strBattery, strPopUp
DIM nMinutes
DIM nSeconds
DIM sMessage
DIM GBL_REG_RUN_KEY
DIM GBL_REG_POLICY_KEY


'Driver Variables
DIM driveGuardReturn
DIM audioReturn
DIM cameraReturn
DIM chipsetReturn
DIM videoReturn
DIM rapidSotrageReturn
DIM smartCardReturn
DIM USBexhcReturn
DIM wifiReturn
DIM hotkeyReturn
DIM realtekPCIEReturn
DIM DriverOnlyRun : DriverOnlyRun = false
DIM driverFail : driverFail = false

DIM installDriveGuard : installDriveGuard = true
DIM installAudio : installAudio = true
DIM installCamera : installCamera = true
DIM installChipset : installChipset = true	
DIM installVideo : installVideo = true
DIM installRapidStorage : installRapidStorage = true
DIM IMEI : IMEI = true
DIM installUSBexhc : installUSBexhc = true
DIM installWifi : installWifi = true
DIM installSmartCard : installSmartCard = true



sublog "       -------------------------------------------------------------------"
sublog "       |                    Return Codes Hash Table                      |"
sublog "       -------------------------------------------------------------------"
sublog "       #  0     Success (Or during Driver Check if no device found)      #"
sublog "       #  200   Driver is at Higher version                              #"
sublog "       #  199   Driver is at Install version                             #" 
sublog "       #  198   Driver is at Rollback version                            #"
sublog "       #  197   Driver is at Unknown version                             #"
sublog "       #  196   O.S. Not supported                                       #"
sublog "       #  189   No Device Found, Will Upgrade                            #"

'*-------------------------------='
'*-------- Begin New Run --------='
'*-------------------------------=' 
If WScript.Arguments.Named.Exists("Rollback") Then
	subLog "----------------------- ROLLBACK RUN -----------------------"
Else
	subLog "----------------------- INSTALL RUN -----------------------"
End If

'*---------------------------------------------------------------='
'* Sets Uninstall Key location based on OS type (64bit or 32bit) ='
'*---------------------------------------------------------------='	
If GBL_OSType = "AMD64" then
	GBL_REG_UNINSTALL_KEY = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
	runLync = """C:\Program Files (x86)\Microsoft Office\Office15\lync.exe"""
	GBL_REG_RUN_KEY = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Run\"
	GBL_REG_POLICY_KEY = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Policies\System"
	X64BitHost = True
ElseIf GBL_OSType = "x86" Then
	GBL_REG_UNINSTALL_KEY = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
	runLync = """C:\Program Files\Microsoft Office\Office15\lync.exe"""
	GBL_REG_RUN_KEY = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run\"
	GBL_REG_POLICY_KEY =  "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\"
	X64BitHost = False
End If

fnModelSelect
fnGetBIOSVersion



'*==========================='
'*--------------------------='
'* Begin BIOS Installation  ='
'*--------------------------='
'*==========================='

'fnCheckBIOSCriteria

fullBIOSUpgrade


Sub fullBIOSUpgrade()

	If DEBUG_ON = False Then On Error Resume Next
    Err.Clear
	
	'turns off audio during bios upgrade
	GBL_strFnErrorCode = FnRunCmd(GBL_ScriptPath & "BiosConfigUtility.exe", " /setconfig:" & GBL_ScriptPath & "config.txt")

	fnGetBIOSVersion

	DIM prevBIOSVersionKey
	prevBIOSVersionKey = fnWMIRegRead(HKEY_LOCAL_MACHINE, GBL_REG_UNINSTALL_KEY & ENH_NUM,"PreviousBIOSVersion")
	
	If not WScript.Arguments.Named.Exists("Rollback") Then
		If prevBIOSVersionKey = "" or prevBIOSVersionKey > UpgradedBiosVersion then
			AddRegBiosVer
			sublog "INFO:	Adding Previous Version """ & GBL_intBIOSVersion & """ to registry for Rollback Purposes"
		Else
			subLog "INFO:	Previous Version is " & prevBIOSVersionKey
		End If
	
	Else 
		subLog "INFO:	Previous Version is " & prevBIOSVersionKey
	End if
	
	
'*-------------------------------------------------------------------------------------------------------='
'* Determines if BIOS Version meets requirements for Upgrade and sets BIOS upgrade on reboot. 		     ='
'* -If "Rollback"- Checks BIOS Version Meets Requirements, Runs Clean-up, sets BIOS downgrade on Reboot. ='
'*-------------------------------------------------------------------------------------------------------='	
	
	IF (GBL_intBIOSVersion) >= UpgradedBiosVersion Then
'		----If Rollback-----
		If WScript.Arguments.Named.Exists("Rollback") Then
			subRemoveScheduledTask
			if prevBIOSVersionKey <> "" then 
				prevBIOSVersionKey = Cint(prevBIOSVersionKey)
			else
				DeleteRegKey
				subLog "INFO:	No Previous BIOS found, Downgrade not needed"
				subEndScript (6)
			end if
	
			IF (GBL_intBIOSVersion) = UpgradedBiosVersion Then
				If prevBIOSVersionKey => UpgradedBiosVersion Then
					DeleteRegKey
					subLog "INFO:	Previous BIOS is equal or greater than " & UpgradedBiosVersion & " Downgrade not needed"
					subEndScript (6)
				Else
					fnStopServices
					subLog "INFO:	BIOS Meets Requirements, Rollback initiated"
						'Check for Power adapter status
						'fnPowerAdapterCheck
						DeleteRegKey
					WScript.Sleep 3000
						'what it says
						SuspendBitLocker
					WScript.Sleep 3000
						'set scheduled task
						subTaskVerifyFailSafe
					WScript.Sleep 3000
						'set scheduled task
						subTaskVerifyBitLockerOff
					WScript.Sleep 3000
						'Turns off audio alerts in BIOS if needed
						GBL_strFnErrorCode = FnRunCmd(GBL_ScriptPath & "BiosConfigUtility.exe", " /setconfig:" & GBL_ScriptPath & "config.txt")
						'downgrade
						DowngradeBIOS
					WScript.Sleep 180000
						'reboot timer set 
						If WScript.Arguments.Named.Exists("MaintenanceWindow") Then
							GBL_wshShell.run GBL_ScriptPath & "Reboottimer.exe /TIME=1 /REBOOT=YES"
						else
							GBL_wshShell.run GBL_ScriptPath & "Reboottimer.exe /TIME=15 /REBOOT=YES"
						end if
					SubEndScript (3010)
				End if
			Else
				subLog "INFO:	BIOS Version does not Meet Downgrade Requirements, Program Complete. Deleting Registry Key"
				DeleteRegKey
				SubEndScript (0)
			End if
		End If
		subLog "INFO:	BIOS Greater than or equal to " & UpgradedBiosVersion & ", Upgrade not needed. Program complete Adding Registry key" 
		AddRegEntry
		SubEndScript (0)
	Else	
		If WScript.Arguments.Named.Exists("Rollback") Then
			subRemoveScheduledTask
			subLog "INFO:	BIOS Version does not Meet Downgrade Requirements, Program Complete. Deleting Registry Key"
			DeleteRegKey
			SubEndScript (0)

'*-------------------------------------------='
'* Stages upgrade if BIOS meets requirements ='
'*-------------------------------------------='
		Else
			subLog "INFO:	BIOS Needs Upgrade."
				'Check for Power adapter status
				'fnPowerAdapterCheck
				fnStopServices
			WScript.Sleep 3000
				'what it says
				SuspendBitLocker
			WScript.Sleep 3000
				'set scheduled task
				subTaskVerifyFailSafe
			WScript.Sleep 3000
				'set scheduled task
				subTaskVerifyBitLockerOff
			WScript.Sleep 3000
				'turns off audio during bios upgrade
			GBL_strFnErrorCode = FnRunCmd(GBL_ScriptPath & "BiosConfigUtility.exe", " /setconfig:" & GBL_ScriptPath & "config.txt")
				'Upgrades the BIOS
				UpgradeBIOS
			WScript.Sleep 180000
			AddRegEntry
				'reboot timer set
				If WScript.Arguments.Named.Exists("MaintenanceWindow") Then
					GBL_wshShell.run GBL_ScriptPath & "Reboottimer.exe /TIME=1 /REBOOT=YES"
				else
					GBL_wshShell.run GBL_ScriptPath & "Reboottimer.exe /TIME=15 /REBOOT=YES"
				end if
			SubEndScript (3010)
		End if
	
	End if	
	
End Sub

'*==========================='
'*--------------------------='
'*  End BIOS Installation   ='
'*--------------------------='
'*==========================='







'*****************************'
'* Subroutines and Functions *'
'*****************************'

'*================================================================================================================*'
'* Name     : SuspendBitLocker                                                                                    *'
'* Purpose  : Call Script to Suspend Bitlocker and Setup Task Scheduler to Resume Bitlocker Protection on Startup *'
'* Usage	   : Suspend Bitlocker till Reboot                                                                    *'
'*================================================================================================================*'
Sub SuspendBitLocker

	subLog "INFO:	Attempting BitLocker Suspension"
	GBL_strBitPath = GBL_ScriptPath & "BitLockerOnOff.vbs"
	GBL_strFnErrorCode = FnRunCmd(GBL_strBitPath, "DISABLEFORREBOOT C")
	wscript.sleep 3000

	If GBL_strFnErrorCode <> 0 Then
		subLog "ERROR:	Bitlocker Suspend Script Failed. Exiting script"
		SubEndScript(GBL_strFnErrorCode)
	Else
		subLog "INFO:	Bitlocker Suspend Script Succeeded. Continuing with the script"
	End If
End Sub

'*========================================================*'
'* Name     : UpgradeBIOS                                 *'
'* Purpose  : Executes a command line to Upgrade the BIOS *'
'*========================================================*'
Sub UpgradeBIOS
		intReturn = GBL_wshShell.Run(GBL_ScriptPath & strBiosCmd, 0, false)
	If intReturn = 0 then
		subLog "INFO:	Return code:" & intReturn
		subLog "INFO:	PC model within Scope. Upgrading BIOS."
		subLog "INFO:	BIOS command succeeded and will install during reboot"
	Else
		subLog "ERROR:	Return code:" & intReturn
		subLog "ERROR:	BIOS upgrade failed."
		SubEndScript (1)
	End if
End Sub

'*=================================================================*'
'* Name     : DowngradeBIOS                                        *'
'* Purpose  : N\A (Executes a command line to Downgrade the BIOS to A08) *' 
'*=================================================================*'
 Sub DowngradeBIOS
	intReturn = GBL_wshShell.Run(GBL_ScriptPath & strBiosCmd, 0, false)
		If intReturn = 0 then
			subLog "INFO:	Return code:" & intReturn
			subLog "INFO:	PC model within Scope. Downgrading BIOS"
			subLog "INFO:	BIOS command succeeded and will install during reboot"	
		Else
			subLog "ERROR:	Return code:" & intReturn
			subLog "ERROR:	BIOS Downgrade failed."
			SubEndScript (1)
		End if 
	
End Sub
 
'*==========================================================*'
'* Name     : FnRunCmd                                      *'
'* Purpose  : Executes a command line                       *'
'* Input    : strPath - the path of the executable , strArg *'
'* Return   : return value of the executable                *'
'* Usage	:                                               *'
'*==========================================================*'
 Function fnRunCmd (strPath, strArg)
    
	If Not DEBUG_ON Then
        On Error Resume Next
    End If
    
'    DIM wshShell
'	Set wshShell=CreateObject("Wscript.Shell")
 
'	DIM ObjFSO
'	Set ObjFSO=CreateObject("Scripting.FileSystemObject")
	
	DIM retCode
	
    If len(strPath) > 0 and GBL_wshFileSys.FileExists(strPath) Then
		If len(strArg) > 0 Then
			'subLog "INFO:	Running the command " & fnFormatPath (strPath) & " " &  strArg
			retCode = GBL_wshShell.Run (fnFormatPath (strPath) & " " &  strArg ,0 ,True)
			'subLog "INFO:	Returncode:" &  retCode
		Else
			'subLog "INFO:	Running the command " & fnFormatPath (strPath)
			retCode = GBL_wshShell.Run ("%COMSPEC% /c " & fnFormatPath (strPath) ,0 ,True)
			'subLog "INFO:	Returncode:" &  retCode
		End If 
    Else
		'subLog "INFO:	File " & strPath  & " Not Found "
		retCode = -1
		'subLog "INFO:	Returncode:" &  retCode
    End If

	FnRunCmd = retCode
	'subLog "INFO:	The Return code for the Function is : " & FnRunCmd	
End Function

'*==================================================================*'
'* Name      : fnFormatPath                                         *'
'* Purpose   : If there are spaces in the path add quotes around it *'
'* Input     : strPath - the path to be formatted                   *'
'* Return    : formatted path                                       *'
'* Usage	 : fnFormatPath ("C:\Program Files")                    *'
'*==================================================================*'
Function fnFormatPath(strPath)
    If Not DEBUG_ON Then
        On Error Resume Next
    End If
    If InStr(1, strPath, " ", 1) > 0 Then
        fnFormatPath = Chr(34) & trim(strPath) & Chr(34)
    Else
        fnFormatPath = strPath
    End If
End Function

' *===================================================================*'
' * NAME     : subLog                                                 *'
' * PURPOSE  : writes entries to the log file specified at the top of *'
' *            the script if ENABLE_LOGGING = True.  Also, if         *'
' *            DEBUG_ON = True it will echo the entry to the screen.  *'
' *            This reduces the need for extra subDebug calls.        *'
' * INPUT    : strEntry = line to be written to the file              *'
' * OUTPUT   : Writes entry to log file                               *'
' * RETURN   : N/A                                                    *'
' * Errors	 : N/A                                                    *'
' *===================================================================*'
Sub subLog(ByVal strEntry)
	DIM strLogFolder
	DIM strLogFile
	DIM objFile

	If DEBUG_ON = False Then
		On Error Resume Next
	End If

	If ENABLE_LOGGING = False Then
		wscript.echo strEntry
		Exit Sub
	End If
    strLogFile = GBL_LogFile
	strLogFolder = GBL_LogFolder
    
	If Not GBL_wshFileSys.FolderExists(strLogFolder) Then
		GBL_wshFileSys.CreateFolder (strLogFolder)
	End If
	 '* Folder creation failed so exit routine
	If Not GBL_wshFileSys.FolderExists(strLogFolder) Then
		Exit Sub
	End If
	If Not GBL_wshFileSys.FileExists(strLogFile) Then
        Err.Clear
		Set objFile = GBL_wshFileSys.OpenTextFile(strLogFile, FOR_WRITING, True)
        
		objFile.WriteLine "Log File For " & GBL_Name
		DIM colItems, objItem
		Set colItems = GBL_objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
		For Each objItem In colItems
			objFile.WriteLine "Hostname " & objItem.Name
			objFile.WriteLine "Manufacturer " & objItem.Manufacturer
			objFile.WriteLine "Model " & objItem.Model
		Next
		On Error Resume Next
		GBL_objRegistry.GetStringValue HKEY_LOCAL_MACHINE, "Software\GM\Core", "Current Image", strValue
		objFile.WriteLine "Current Image " & strValue
		GBL_objRegistry.GetStringValue HKEY_LOCAL_MACHINE, "Software\GM\Core", "Original Image", strValue
		objFile.WriteLine "Original Image " & strValue
		If DEBUG_ON Then On Error GOTO 0
		objFile.WriteLine fnFormatNow
		objFile.WriteLine "------------------------------"
		objFile.Close
    End IF
    Set objFile = GBL_wshFileSys.OpenTextFile(strLogFile, FOR_APPENDING, True)
	If WScript.Arguments.Named.Exists("CheckOnly") or WScript.Arguments.Named.Exists("CheckRollback") Then strEntry = "CheckOnly: " & strEntry
	objFile.WriteLine "[" & fnFormatNow & "]  " & strEntry
	objFile.Close
End Sub

'*============================*'
'* NAME		: SubEndScript    *'                                     
'* Input	: N/A             *'
'* Return	: N/A             *'                                
'* PURPOSE	: Release objects *'
'*============================*'
Sub SubEndScript(ByVal lExitCode)
	If DEBUG_ON = False Then On Error Resume Next
	DIM oExitCode  : oExitCode = lExitCode
	' This will exit with 709 for EDWS
	'If GBL_Env_EDWS and lExitCode <> 0 Then oExitCode = 709
	subLog "EXIT: Script ended with exit code " & oExitCode
	Set GBL_wshShell = Nothing
	Set GBL_wshFileSys = Nothing
	Set GBL_objRegistry = Nothing
	Set GBL_WshNetwork = Nothing
	Set GBL_ObjArgs = Nothing
	WScript.Quit(oExitCode)
End Sub

'*================================================================*'
'* Name    : fnCheckOSType                                        *' 
'* Purpose : Checks to see the OS version, returns whatever is in *'
'*           PROCESSOR_ARCHITECTURE.                              *'
'* Input   : N/A                                                  *'
'* Output  : N/A                                                  *'
'* Return  : Processor Architecture, expected to be x86 or AMD64. *'
'*================================================================*'
Function fnCheckOSType
	If DEBUG_ON = False Then On Error Resume Next
	DIM strOSType
	DIM wshShell
	Set wshShell = CreateObject("WScript.Shell")
	strOSType = WshShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")
	fnCheckOSType = strOSType
End Function

'*============================================*'
' *  NAME     : fnFormatNow                   *'
' *  PURPOSE  : formats date and time for log *'
' *  INPUT    : N/A                           *'
' *  OUTPUT   : N/A                           *'
' *  RETURN   : MM-DD-YYYY HH:MM:SS           *'
' *  Errors   : N/A                           *'
'*============================================*'
Function fnFormatNow()
	fnFormatNow = Right("0"& Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & "-" & Year(Now) & "  " & Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2)
End Function

'*================================================*'
'* NAME     : AddRegEntry                         *'
'* PURPOSE  : Create a key and add a string value *'
'* INPUT    : N/A                                 *'
'* OUTPUT   : N/A                                 *'
'* RETURN   : N/A                                 *'
'*================================================*'
Sub AddRegEntry
	If DEBUG_ON = False Then On Error Resume Next
	Err.Clear
	GBL_objRegistry.CreateKey HKEY_LOCAL_MACHINE,GBL_REG_UNINSTALL_KEY & ENH_NUM
	If Err.Number = 0 Then 
		GBL_objRegistry.SetStringValue HKEY_LOCAL_MACHINE,GBL_REG_UNINSTALL_KEY & ENH_NUM,"DisplayName",ENH_BIOS
		If Err.Number = 0 Then 
			subLog "INFO:	Registry Value Set - 'HKLM\" & GBL_REG_UNINSTALL_KEY & ENH_NUM & "' - DisplayName " & "(Value: " & ENH_BIOS & ")'"
		Else
			subLog "Error: Failed to set registry value - 'HKLM\" & GBL_REG_UNINSTALL_KEY & ENH_NUM & "' - DisplayName " & "(Value: " & ENH_BIOS & ")'"
			subLog "Error: " & Err.Number & " - " & Err.Description
			SubEndScript(RC_REGISTRY_ERROR)
		End If
	Else
		subLog "Error: Failed to create registry key - 'HKLM\" & GBL_REG_UNINSTALL_KEY & ENH_NUM
		subLog "Error: " & Err.Number & " - " & Err.Description
		SubEndScript(RC_REGISTRY_ERROR)
	End If
End Sub

'*===================================*'
'* NAME     : DeleteRegKey           *'
'* PURPOSE  : Deletes a registry Key *'
'* INPUT    : N/A                    *'
'* OUTPUT   : N/A                    *'
'* RETURN   : N/A                    *'
'*===================================*'
Sub DeleteRegKey
	If DEBUG_ON = False Then On Error Resume Next
	Err.Clear
	GBL_objRegistry.DeleteKey HKEY_LOCAL_MACHINE,GBL_REG_UNINSTALL_KEY & REG_VALFAIL
	GBL_objRegistry.DeleteKey HKEY_LOCAL_MACHINE,GBL_REG_UNINSTALL_KEY & ENH_NUM
	If Err.Number = 0 Then 
		subLog "INFO:	Successfully deleted Registry Key for BIOS Upgrade"
	Else
		subLog "Error: Failed to delete registry key - BIOS Upgrade"
		subLog "Error: " & Err.Number & " - " & Err.Description
		SubEndScript(RC_REGISTRY_ERROR)
	End If	
End Sub 

'*=======================================*'
'* NAME     : OldScriptCleanup           *'
'* PURPOSE  : Cleans Up Previous Package *'
'* INPUT    : N/A                        *'
'* OUTPUT   : N/A                        *'
'* RETURN   : N/A                        *'
'*=======================================*'
Sub OldScriptCleanup
	If DEBUG_ON = False Then On Error Resume Next
	Err.Clear
	GBL_objRegistry.DeleteKey HKEY_LOCAL_MACHINE,GBL_REG_UNINSTALL_KEY & OldReg
	GBL_objRegistry.DeleteKey HKEY_LOCAL_MACHINE,GBL_REG_UNINSTALL_KEY & OldRegFail
	subRemoveScheduledTask
	If Err.Number = 0 Then 
		subLog "INFO:	Successfully deleted old Registry Key and old scheduled task"
	Else
		subLog "Error: Failed to delete old registry key and/or scheduled task"
		subLog "Error: " & Err.Number & " - " & Err.Description
		SubEndScript(RC_REGISTRY_ERROR)
	End If	
End Sub 

'*================================================*'
'* NAME     : AddRegBiosVer                       *'
'* PURPOSE  : Create a key and add a string value *'
'* INPUT    : N/A                                 *'
'* OUTPUT   : N/A                                 *'
'* RETURN   : N/A                                 *'
'*================================================*'
Sub AddRegBiosVer
	If DEBUG_ON = False Then On Error Resume Next
	Err.Clear
	GBL_objRegistry.CreateKey HKEY_LOCAL_MACHINE,GBL_REG_UNINSTALL_KEY & ENH_NUM
	If Err.Number = 0 Then 
		GBL_objRegistry.SetStringValue HKEY_LOCAL_MACHINE,GBL_REG_UNINSTALL_KEY & ENH_NUM,"PreviousBIOSVersion",GBL_intBIOSVersion
		If Err.Number = 0 Then 
			subLog "INFO:	Registry Value Set - 'HKLM\" & GBL_REG_UNINSTALL_KEY & ENH_NUM & "' - PreviousBIOSVersion " & "(Value: " & GBL_intBIOSVersion & ")'"
		Else
			subLog "Error: Failed to set registry value - 'HKLM\" & GBL_REG_UNINSTALL_KEY & ENH_NUM & "' - PreviousBIOSVersion " & "(Value: " & GBL_intBIOSVersion & ")'"
			subLog "Error: " & Err.Number & " - " & Err.Description
			SubEndScript(RC_REGISTRY_ERROR)
		End If
	Else
		subLog "Error: Failed to create registry key - 'HKLM\" & GBL_REG_UNINSTALL_KEY & ENH_NUM
		subLog "Error: " & Err.Number & " - " & Err.Description
		SubEndScript(RC_REGISTRY_ERROR)
	End If	
End Sub

'*=========================================================*'
'* NAME     : fnWMIRegRead                                 *'
'* PURPOSE  : Uses WMI to read a value from the registry   *'
'* INPUT    : RegHive = Root Hive, e.g. HKEY_LOCAL_MACHINE *'
'*             These are defined at the top of the script  *'
'*             strKey = Key name                           *'
'*             strValueName = name of value to be read     *'
'* OUTPUT   : N/A                                          *'
'* RETURN   : The value in the key or an empty string      *'
'*=========================================================*'
Function fnWMIRegRead(RegHive, strKey, strValueName)
	DIM strValue
	DIM lReturn    
	If DEBUG_ON = False Then On Error Resume Next
	strValue = ""
	lReturn = GBL_objRegistry.GetStringValue(RegHive, strKey, strValueName, strValue)	
	If lReturn <> 0 Then
        	strValue = ""
	End If
    	fnWMIRegRead = strValue
End Function

'*=====================================================*'
'* NAME: fnCreateXML                                   *'
'* PURPOSE  : used to set properties of a created task *'
'* INPUT    : N/A                                      *'
'* OUTPUT   : N/A                                      *'
'* RETURN   : N/A                                      *'
'*=====================================================*'
Function fnCreateXML (byVal strXML)
	DIM objFile
	fnCreateXML = False
	If GBL_wshFileSys.FileExists(strXML) Then GBL_wshFileSys.DeleteFile strXML, True
	wscript.sleep 1000
	If GBL_wshFileSys.FileExists(strXML) Then Exit Function
	Set objFile = GBL_wshFileSys.OpenTextFile(strXML, FOR_WRITING, True)
objFile.WriteLine "<?xml version=""1.0"" encoding=""UTF-16""?>"
	objFile.WriteLine "<Task version=""1.2"" xmlns=""http://schemas.microsoft.com/windows/2004/02/mit/task"">"
	objFile.WriteLine "  <RegistrationInfo>"
	objFile.WriteLine "    <Date>" & fnTimeStampNow & "</Date>"
	objFile.WriteLine "    <Author>SYSTEM</Author>"
	objFile.WriteLine "  </RegistrationInfo>"
	objFile.WriteLine "  <Triggers>"
	objFile.WriteLine "    <BootTrigger>"
	objFile.WriteLine "      <StartBoundary>2014-12-02T10:00:00</StartBoundary>"
	objFile.WriteLine "      <Enabled>true</Enabled>"
	objFile.WriteLine "    </BootTrigger>"
	objFile.WriteLine "  </Triggers>"
	objFile.WriteLine "  <Principals>"
    objFile.WriteLine "    <Principal id=""Author"">"
    objFile.WriteLine "      <UserId>S-1-5-18</UserId>"
	objFile.WriteLine "      <RunLevel>HighestAvailable</RunLevel>"
	objFile.WriteLine "    </Principal>"
	objFile.WriteLine "  </Principals>"
	objFile.WriteLine "  <Settings>"
	objFile.WriteLine "    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>"
	objFile.WriteLine "	   <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>"
	objFile.WriteLine "    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>"
	objFile.WriteLine "    <AllowHardTerminate>true</AllowHardTerminate>"
	objFile.WriteLine "    <StartWhenAvailable>false</StartWhenAvailable>"
	objFile.WriteLine "    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>"
	objFile.WriteLine "    <IdleSettings>"
	objFile.WriteLine "      <Duration>PT10M</Duration>"
	objFile.WriteLine "      <WaitTimeout>PT1H</WaitTimeout>"
	objFile.WriteLine "      <StopOnIdleEnd>false</StopOnIdleEnd>"
	objFile.WriteLine "      <RestartOnIdle>false</RestartOnIdle>"
	objFile.WriteLine "    </IdleSettings>"
	objFile.WriteLine "    <AllowStartOnDemand>true</AllowStartOnDemand>"
	objFile.WriteLine "    <Enabled>true</Enabled>"
	objFile.WriteLine "    <Hidden>false</Hidden>"
	objFile.WriteLine "    <RunOnlyIfIdle>false</RunOnlyIfIdle>"
	objFile.WriteLine "    <WakeToRun>false</WakeToRun>"
	objFile.WriteLine "    <ExecutionTimeLimit>PT1H</ExecutionTimeLimit>"
	objFile.WriteLine "    <Priority>7</Priority>"
	objFile.WriteLine "  </Settings>"
	objFile.WriteLine "  <Actions Context=""Author"">"
	objFile.WriteLine "    <Exec>"
	objFile.WriteLine "     <Command>C:\Windows\System32\cscript.exe</Command>"
	If WScript.Arguments.Named.Exists("Rollback") Then
		objFile.WriteLine "        <Arguments>//B " & GBL_ScriptPath & "Failsafe.vbs /rollback</Arguments>"
	Else 
		objFile.WriteLine "        <Arguments>//B " & GBL_ScriptPath & "Failsafe.vbs</Arguments>"
	End if
	objFile.WriteLine "    </Exec>"
	objFile.WriteLine "  </Actions>"
	objFile.WriteLine "</Task>"
	objFile.Close
	If GBL_wshFileSys.FileExists(strXML) Then fnCreateXML = True
End Function

'*=====================================================*'
'* NAME: fnCreateBitlockeroffXML                       *'
'* PURPOSE  : used to set properties of a created task *'
'* INPUT    : N/A                                      *'
'* OUTPUT   : N/A                                      *'
'* RETURN   : N/A                                      *'
'*=====================================================*'
Function fnCreateBitlockeroffXML (byVal strXML)
	DIM objFile
	fnCreateBitlockeroffXML = False
	If GBL_wshFileSys.FileExists(strXML) Then GBL_wshFileSys.DeleteFile strXML, True
	wscript.sleep 1000
	If GBL_wshFileSys.FileExists(strXML) Then Exit Function
	Set objFile = GBL_wshFileSys.OpenTextFile(strXML, FOR_WRITING, True)
	objFile.WriteLine "<?xml version=""1.0"" encoding=""UTF-16""?>"
	objFile.WriteLine "<Task version=""1.2"" xmlns=""http://schemas.microsoft.com/windows/2004/02/mit/task"">"
	objFile.WriteLine "  <RegistrationInfo>"
	objFile.WriteLine "    <Date>" & fnTimeStampNow & "</Date>"
	objFile.WriteLine "    <Author>SYSTEM</Author>"
	objFile.WriteLine "    <Description>Disable bitlocker on shutdown for BIOS upgrade</Description>"
	objFile.WriteLine "  </RegistrationInfo>"
	objFile.WriteLine " <Triggers>"
	objFile.WriteLine "    <EventTrigger>"
	objFile.WriteLine "      <Enabled>true</Enabled>"
	objFile.WriteLine "      <Subscription>&lt;QueryList&gt;&lt;Query Id=""0"" Path=""System""&gt;&lt;Select Path=""System""&gt;*[System[Provider[@Name='User32'] and EventID=1074]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>"
	objFile.WriteLine "    </EventTrigger>"
	objFile.WriteLine "  </Triggers>"
	objFile.WriteLine "  <Principals>"
	objFile.WriteLine "    <Principal id=""Author"">"
	objFile.WriteLine "      <UserId>S-1-5-18</UserId>"
	objFile.WriteLine "      <RunLevel>HighestAvailable</RunLevel>"
	objFile.WriteLine "    </Principal>"
	objFile.WriteLine "  </Principals>"
	objFile.WriteLine "  <Settings>"
	objFile.WriteLine "   <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>"
	objFile.WriteLine "    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>"
	objFile.WriteLine "    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>"
	objFile.WriteLine "    <AllowHardTerminate>true</AllowHardTerminate>"
	objFile.WriteLine "    <StartWhenAvailable>false</StartWhenAvailable>"
	objFile.WriteLine "    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>"
	objFile.WriteLine "    <IdleSettings>"
	objFile.WriteLine "      <StopOnIdleEnd>false</StopOnIdleEnd>"
	objFile.WriteLine "      <RestartOnIdle>false</RestartOnIdle>"
	objFile.WriteLine "    </IdleSettings>"
	objFile.WriteLine "    <AllowStartOnDemand>true</AllowStartOnDemand>"
	objFile.WriteLine "    <Enabled>true</Enabled>"
	objFile.WriteLine "    <Hidden>false</Hidden>"
	objFile.WriteLine "    <RunOnlyIfIdle>false</RunOnlyIfIdle>"
	objFile.WriteLine "    <WakeToRun>false</WakeToRun>"
	objFile.WriteLine "    <ExecutionTimeLimit>PT1H</ExecutionTimeLimit>"
	objFile.WriteLine "    <Priority>7</Priority>"
	objFile.WriteLine "  </Settings>"
	objFile.WriteLine "  <Actions Context=""Author"">"
	objFile.WriteLine "    <Exec>"
	objFile.WriteLine "      <Command>C:\Windows\System32\cscript.exe</Command>"
	objFile.WriteLine "      <Arguments>//B " & GBL_ScriptPath & "BtlkSuspend.vbs</Arguments>"
	objFile.WriteLine "    </Exec>"
	objFile.WriteLine " </Actions>"
	objFile.WriteLine "</Task>"
	objFile.Close
	If GBL_wshFileSys.FileExists(strXML) Then fnCreateBitlockeroffXML = True
End Function

'*=======================================*'
'* NAME     : fnCreateScheduledTask      *'
'* PURPOSE  : Closes the Running Process *'
'* INPUT    : N/A                        *'
'* OUTPUT   : N/A                        *'
'* RETURN   : True/False                 *'
'*=======================================*'
Function fnCreateScheduledTask (ByVal strXML, ByVal strTaskName)

	If DEBUG_ON = False Then On Error Resume Next
	DIM schtasksCreateCommand, xReturn
	DIM dst_wshShell : Set dst_wshShell = CreateObject("WScript.Shell")
	schtasksCreateCommand = "schtasks /Create /XML " & strXML & " /TN " & strTaskName
	xReturn = dst_wshShell.run (schtasksCreateCommand,0, True)
	If xReturn = 0 Then fnCreateScheduledTask = True
	If xReturn <> 0 Then fnCreateScheduledTask = False
End Function

'*=======================================*'
'* NAME     : fnDeleteScheduledTask      *'
'* PURPOSE  : Deletes the Scheduled Task *'
'* INPUT    : N/A                        *'
'* OUTPUT   : N/A                        *'
'* RETURN   : True/False                 *'
'*=======================================*'
Function fnDeleteScheduledTask (ByVal strTaskName)

	If DEBUG_ON = False Then On Error Resume Next
	DIM schtasksDeleteCommand, xReturn
	DIM dst_wshShell : Set dst_wshShell = CreateObject("WScript.Shell")
		'Delete the scheduled task
	schtasksDeleteCommand = "schtasks /Delete /TN " & strTaskName & " /F"
	xReturn = dst_wshShell.run (schtasksDeleteCommand,0,True)
	If xReturn = 0 Then fnDeleteScheduledTask = True
	If xReturn <> 0 Then fnDeleteScheduledTask = False
End Function

'*=========================================*'
'* NAME     : subRemoveScheduledTask       *'
'* PURPOSE  : Removes Task and logs result *'
'* INPUT    : N/A                          *'
'* OUTPUT   : N/A                          *'
'* RETURN   : N/A                          *'
'*=========================================*'
Sub subRemoveScheduledTask
	
	sTaskReturn = fnDeleteScheduledTask(GBL_FailSafeTask)
	sTaskReturn = fnDeleteScheduledTask(GBL_MBamOffTask)
	If sTaskReturn Then
		subLog "INFO:	Successfully uninstalled task "
	Else
		subLog "INFO:	Task Already Removed"
	End If
End Sub
	
'*======================================================*'
'* NAME     : fnCheckScheduledTask                      *'
'* PURPOSE  : Checks for presence of the Scheduled Task *'
'* INPUT    : N/A                                       *'
'* OUTPUT   : N/A                                       *'
'* RETURN   : True/False                                *'
'*======================================================*'
Function fnCheckScheduledTask (ByVal GBL_TaskName)

	If DEBUG_ON = False Then On Error Resume Next
		DIM schtasksCheckCommand, xReturn
		DIM cst_wshShell : Set cst_wshShell = CreateObject("WScript.Shell")
		'Query the scheduled task
		schtasksCheckCommand = "schtasks /Query /TN " & strTaskName
		xReturn = cst_wshShell.run (schtasksCheckCommand,0,True)
		If xReturn = 0 Then fnCheckScheduledTask = True
		If xReturn <> 0 Then fnCheckScheduledTask = False
End Function

'*===============================================================*'
'* NAME     : subTaskVerifyFailSafe                              *'
'* PURPOSE  : Creates Task if task is not found and logs results *'
'*              Used to create Failsafe task to run on startup   *'
'* INPUT    : N/A                                                *'
'* OUTPUT   : N/A                                                *'
'* RETURN   : True/False                                         *'
'*===============================================================*'
sub subTaskVerifyFailSafe
	strTaskName	   	   = "FailSafeBIOSCheck"
	GBL_TempXMLFolder  = "C:\Temp\"
	GBL_XML_Name       = "BIOS_Task.XML"
	GBL_TaskName	   = "FailSafeBIOSCheck"
	
			sTaskReturn = fnCheckScheduledTask (GBL_TaskName)
			If Not sTaskReturn Then
				If Not GBL_wshFileSys.FolderExists (GBL_TempXMLFolder) Then GBL_wshFileSys.CreateFolder TempXMLFolder
				sTaskReturn = fnCreateXML(GBL_TempXMLFolder & GBL_XML_Name)
				If sTaskReturn Then
					sTaskReturn = fnCreateScheduledTask(GBL_TempXMLFolder & GBL_XML_Name, GBL_TaskName)
					If GBL_wshFileSys.FileExists (GBL_TempXMLFolder & GBL_XML_Name) Then GBL_wshFileSys.DeleteFile GBL_TempXMLFolder & GBL_XML_Name, True
					If sTaskReturn Then
						subLog "INFO:	Create FailSafe Scheduled Task Returned success"
					Else
						subLog "INFO:	FailSafe Scheduled Task already installed"
					End If
				Else
					SubEndScript 151
				End If
				
			End If
End sub

'*===============================================================*'
'* NAME     : subTaskVerifyBitLockerOff                          *'
'* PURPOSE  : Creates Task to disable bitlocker on reboot        *'
'* INPUT    : N/A                                                *'
'* OUTPUT   : N/A                                                *'
'* RETURN   : True/False                                         *'
'*===============================================================*'
sub subTaskVerifyBitLockerOff
	strTaskName	       = "BIOSSafeMBamoff"
	GBL_TempXMLFolder  = "C:\Temp\"
	GBL_XML_Name       = "BitLocker_off_Task.XML"
	GBL_TaskName	   = "BIOSSafeMBamoff"
	
			sTaskReturn = fnCheckScheduledTask (GBL_TaskName)
			If Not sTaskReturn Then
				If Not GBL_wshFileSys.FolderExists (GBL_TempXMLFolder) Then GBL_wshFileSys.CreateFolder TempXMLFolder
				sTaskReturn = fnCreateBitlockeroffXML(GBL_TempXMLFolder & GBL_XML_Name)
				If sTaskReturn Then
					sTaskReturn = fnCreateScheduledTask(GBL_TempXMLFolder & GBL_XML_Name, GBL_TaskName)
					'If GBL_wshFileSys.FileExists (GBL_TempXMLFolder & GBL_XML_Name) Then GBL_wshFileSys.DeleteFile GBL_TempXMLFolder & GBL_XML_Name, True
					If sTaskReturn Then
						subLog "INFO:	Create BIOSSafeMBamOff Scheduled Task Returned success"
					Else
						subLog "INFO:	BIOSSafeMBamOff Scheduled Task already installed"
					End If
				Else
					SubEndScript 151
				End If
				
			End If
End sub

'*======================================================================*'
'* Name    : fnTimeStampNow                                             *'
'* Purpose : Formats Date and Time for XML Creation and other functions *'
'* Input   : N/A                                                        *'
'* Output  : N/A                                                        *'
'* Return  : Date and time, 2014-01-01T12:00:00.0145077                 *'
'*======================================================================*'
Function fnTimeStampNow
    DIM strSafeDate
    DIM strSafeTime
    DIM strDateTime
  	strSafeDate = DatePart("yyyy",Date) & "-" & Right("0" & DatePart("m",Date), 2) & "-" & Right("0" & DatePart("d",Date), 2)
    strSafeTime = Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2)
    fnTimeStampNow = strSafeDate & "T" & strSafeTime
End Function


'===========================================================================
' * NAME     : fnStopServices
' * PURPOSE  : Stops the Mcafee Agent Running service 
' * RETURN   : True or False
'===========================================================================
Function fnStopServices()
	Dim colListOfServices, objService, strServiceName, intRC, objMaservice
	
	Set objMaservice = gbl_objWMIService.Get("Win32_Service.Name='Masvc'")
	SubLog " Initial Service State is : " & objMaService.state
	If DEBUG_ON = False Then 
		On Error Resume Next 
	End If
	strServiceName = "Masvc"
	Set gbl_objWMIService = GetObject("winmgmts:\\.\root\cimv2") 
	Set colListOfServices = gbl_objWMIService.ExecQuery("Select * from Win32_Service Where Name ='" & strServiceName & "'") 
	For Each objService in colListOfServices 
		intRC = gbl_wshShell.Run("Net Stop " &  strServiceName, 0, TRUE)
		Set objMaservice = gbl_objWMIService.Get("Win32_Service.Name='Masvc'")
		WScript.Sleep 10000
		SubLog " Final Service State is : " & objMaService.state
	Next
		
		If intRC = 2 then
			sublog "INFO:	Service already in stopped state: " & intRC
		end if
		
		if intRC <> 2 then
			if intRC <> 0 then
				sublog "ERROR:	Stopping the Service : " & intRC
				subendscript (intRC)
			else
				sublog "Info:	Stopping the Service is Successful"
			End if
		End If
		
End Function

'===========================================================================
' * NAME     : fnCheckBIOSCriteria
' * PURPOSE  : Verifies BIOS is OK to upgrade
' * RETURN   : n/a
'===========================================================================
Function fnCheckBIOSCriteria

	'*------------------------------------------------='
	'* Verify Client has not been previously upgraded ='
	'*------------------------------------------------='
	If not WScript.Arguments.Named.Exists("Rollback") Then
		 If fnWMIRegRead(HKEY_LOCAL_MACHINE, GBL_REG_UNINSTALL_KEY & ENH_NUM,"DisplayName") = ENH_BIOS Then
			subLog "INFO:	Registry key is already present - 'HKLM\" & GBL_REG_UNINSTALL_KEY & ENH_NUM & "'"
			subLog "INFO: Upgrade already run. No action needed"
		End If
	End If

	

'	If not WScript.Arguments.Named.Exists("Rollback") Then 
'		'*---------------------------------------------------='
'		'* Verify video driver is at the appropriate version ='
'		'*---------------------------------------------------='
'		if X64BitHost = True then
'			videoReturn = FnRunCmd(GBL_ScriptPath & "Drivers\x64\intelVideo-x64\DriverCheck.vbs", "/checkonly")
'			wscript.sleep 500
'		else
'			videoReturn = FnRunCmd(GBL_ScriptPath & "Drivers\x86\intelVideo-x86\DriverCheck.vbs", "/checkonly")
'			wscript.sleep 500
'		end if
'		if videoReturn = 199 or videoReturn = 200 then
'			sublog "video Driver meets minimum version requirement to upgrade BIOS. Continuing with BIOS upgrade."
'		else
'			sublog "Video driver not at required version. Cannot upgrade BIOS at this time. Hostname must be re-booted to rerun this package."
'			SubEndScript (197)
'		end if
'	End if
End Function


'*===============================================*'
' NAME     : RegDriverFail                       *'
' PURPOSE  : Create a key and add a string value *'
' INPUT    : N/A                                 *'
' OUTPUT   : N/A                                 *'
' RETURN   : N/A                                 *'
'*===============================================*'

Sub RegDriverFail
	If DEBUG_ON = False Then On Error Resume Next
	Err.Clear
		If WScript.Arguments.Named.Exists("DriverOnly") Then 
			ENH_DrvrFAILED = "CE16563-Driver-retry-failed"
		end if
	
	GBL_objRegistry.CreateKey HKEY_LOCAL_MACHINE,GBL_REG_UNINSTALL_KEY & REG_VALFAIL
	If Err.Number = 0 Then 
		GBL_objRegistry.SetStringValue HKEY_LOCAL_MACHINE,GBL_REG_UNINSTALL_KEY & REG_VALFAIL,"DisplayName",ENH_DrvrFAILED
		If Err.Number = 0 Then 
			SubLog "INFO:	Registry Value Set - 'HKLM\" & GBL_REG_UNINSTALL_KEY & REG_VALFAIL & "' - DisplayName " & "(Value: " & ENH_DrvrFAILED & ")'"
		Else
			SubLog "Error: Failed to set registry value - 'HKLM\" & GBL_REG_UNINSTALL_KEY & REG_VALFAIL & "' - DisplayName " & "(Value: " & ENH_DrvrFAILED & ")'"
			SubLog "Error: " & Err.Number & " - " & Err.Description
			SubEndScript(RC_REGISTRY_ERROR)
		End If
	Else
		SubLog "Error: Failed to create registry key - 'HKLM\" & GBL_REG_UNINSTALL_KEY & REG_VALFAIL
		SubLog "Error: " & Err.Number & " - " & Err.Description
		SubEndScript(RC_REGISTRY_ERROR)
	End If
End Sub


'*================================================*'
'* NAME     : AddRegDrvrSuccess                   *'
'* PURPOSE  : Create a key and add a string value *'
'* INPUT    : N/A                                 *'
'* OUTPUT   : N/A                                 *'
'* RETURN   : N/A                                 *'
'*================================================*'
Sub AddRegDrvrSuccess
	If DEBUG_ON = False Then On Error Resume Next
	Err.Clear
	GBL_objRegistry.CreateKey HKEY_LOCAL_MACHINE,GBL_REG_UNINSTALL_KEY & ENH_NUM
	If Err.Number = 0 Then 
		GBL_objRegistry.SetStringValue HKEY_LOCAL_MACHINE,GBL_REG_UNINSTALL_KEY & ENH_NUM,"DriverSuccess",ENH_DriverSuccess
		If Err.Number = 0 Then 
			subLog "INFO:	Registry Value Set - 'HKLM\" & GBL_REG_UNINSTALL_KEY & ENH_NUM & "' - DriverSuccess " & "(Value: " & ENH_DriverSuccess & ")'"
		Else
			subLog "Error: Failed to set registry value - 'HKLM\" & GBL_REG_UNINSTALL_KEY & ENH_NUM & "' - DriverSuccess " & "(Value: " & ENH_DriverSuccess & ")'"
			subLog "Error: " & Err.Number & " - " & Err.Description
			SubEndScript(RC_REGISTRY_ERROR)
		End If
	Else
		subLog "Error: Failed to create registry key - 'HKLM\" & GBL_REG_UNINSTALL_KEY & ENH_NUM
		subLog "Error: " & Err.Number & " - " & Err.Description
		SubEndScript(RC_REGISTRY_ERROR)
	End If
End Sub



'===========================================================================
' * NAME     : fnPowerAdapterCheck
' * PURPOSE  : checks for Power Adapter
' * RETURN   : n/a
'===========================================================================
Function fnPowerAdapterCheck
	
'*---------------------------------------------------------------------='
'* Initiates Communication Message box to check for A/C power adapter. ='
'*---------------------------------------------------------------------='				
	Set colItems = GBL_objWMIService.ExecQuery("Select * from Win32_Battery")
		For Each objItem in colItems
		strBattery = objItem.BatteryStatus
		Next

	If strBattery = 2 then
			subLog "INFO:	Adapter is Attached, Continuing with BIOS Upgrade"
	Else
		strPopUp = msgbox("An A/C power adapter was NOT detected. Please plug in your A/C power adapter and DO NOT unplug it until AFTER your system reboots and you are logged in.  Click OK to continue the update. If you cannot reach a power source please click 'Cancel' to exit this update." ,1, "System Update Notification")
		if strPopUp = 1 then
		end if

		if strPopUp = 2 then
			subLog "INFO:	Cancelled by User. Exiting script"
			SubEndScript (0)
		end if	

	End if
'*-------------------------------------------------='
'* Loop to check for adapter after user clicks OK. ='
'*-------------------------------------------------='	
	Do while strBattery <> 2
	
	Set colItems = GBL_objWMIService.ExecQuery("Select * from Win32_Battery")
		For Each objItem in colItems
		strBattery=objItem.BatteryStatus
		Next
	
	If strBattery = 2 then
		subLog "INFO:	Adapter is now attached, Continuing with BIOS Upgrade."
	Else
		subLog "INFO:	Adapter is not attached, running Retry Communication"
		strPopUp = msgbox("An A/C power adapter was NOT detected.  To continue, please plug in your A/C power adapter and DO NOT unplug it until AFTER your system reboots.  Click Retry to initiate the update." ,5, "System Update")
		if strPopUp = 4 then
		end if

		if strPopUp = 2 then
			subLog "INFO:	Cancelled by User. Exiting script"
			SubEndScript (0)
		end if

	End if

	Loop
	
end Function

'*============================================================*'
'* NAME     : verifyRegKeysSuccess                            *'
'* PURPOSE  : Verify full success through registry key checks *'
'* INPUT    : N/A                                             *'
'* OUTPUT   : N/A                                             *'
'* RETURN   : N/A                                             *'
'*============================================================*'
Sub verifyRegKeysSuccess

	if fnWMIRegRead(HKEY_LOCAL_MACHINE,GBL_REG_UNINSTALL_KEY & ENH_NUM,"DriverSuccess") = ENH_DriverSuccess then
		if fnWMIRegRead(HKEY_LOCAL_MACHINE,GBL_REG_UNINSTALL_KEY & ENH_NUM,"DisplayName") = ENH_BIOS then
			AddRegAllSuccess
			Sublog "Drivers Were Upgraded Successful, Will validate BIOS upgrade after reboot if needed, see ""failsafe.log""."
		else
			Sublog "Driver upgrades were Successful, BIOS did not upgrade or user cancelled, Retry will be attempted after an estimated 24 hours."
			DeleteUpgradeRegKey
		end if
	else
		DeleteUpgradeRegKey
		Sublog "Driver upgrades were not Successful, Retry will be attempted after an estimated 24 hours."
	end if
end sub


'*=========================================================*'
'* NAME     : fnWMIRegRead                                 *'
'* PURPOSE  : Uses WMI to read a value from the registry   *'
'* INPUT    : RegHive = Root Hive, e.g. HKEY_LOCAL_MACHINE *'
'*             These are defined at the top of the script  *'
'*             strKey = Key name                           *'
'*             strValueName = name of value to be read     *'
'* OUTPUT   : N/A                                          *'
'* RETURN   : The value in the key or an empty string      *'
'*=========================================================*'
Function fnWMIRegRead(RegHive, strKey, strValueName)
	DIM strValue
	DIM lReturn    
	If DEBUG_ON = False Then On Error Resume Next
	strValue = ""
	lReturn = GBL_objRegistry.GetStringValue(RegHive, strKey, strValueName, strValue)	
	If lReturn <> 0 Then
        	strValue = ""
	End If
    	fnWMIRegRead = strValue
End Function

'*================================================*'
'* NAME     : AddRegAllSuccess                    *'
'* PURPOSE  : Create a key and add a string value *'
'* INPUT    : N/A                                 *'
'* OUTPUT   : N/A                                 *'
'* RETURN   : N/A                                 *'
'*================================================*'
Sub AddRegAllSuccess
	If DEBUG_ON = False Then On Error Resume Next
	Err.Clear
	GBL_objRegistry.CreateKey HKEY_LOCAL_MACHINE,GBL_REG_UNINSTALL_KEY & ENH_NUM
	If Err.Number = 0 Then 
		GBL_objRegistry.SetStringValue HKEY_LOCAL_MACHINE,GBL_REG_UNINSTALL_KEY & ENH_NUM,"AllSuccess",ENH_allClear
		If Err.Number = 0 Then 
			subLog "INFO:	Registry Value Set - 'HKLM\" & GBL_REG_UNINSTALL_KEY & ENH_NUM & "' - AllSuccess " & "(Value: " & ENH_allClear & ")'"
		Else
			subLog "Error: Failed to set registry value - 'HKLM\" & GBL_REG_UNINSTALL_KEY & ENH_NUM & "' - AllSuccess " & "(Value: " & ENH_allClear & ")'"
			subLog "Error: " & Err.Number & " - " & Err.Description
			SubEndScript(RC_REGISTRY_ERROR)
		End If
	Else
		subLog "Error: Failed to create registry key - 'HKLM\" & GBL_REG_UNINSTALL_KEY & ENH_NUM
		subLog "Error: " & Err.Number & " - " & Err.Description
		SubEndScript(RC_REGISTRY_ERROR)
	End If
End Sub
	
'*==================================*'
' NAME     : DeleteUpgradeRegKey    *'
' PURPOSE  : Deletes a registry Key *'
' INPUT    : N/A                    *'
' OUTPUT   : N/A                    *'
' RETURN   : N/A                    *'
'*==================================*'
Sub DeleteUpgradeRegKey
	If DEBUG_ON = False Then On Error Resume Next
	Err.Clear
	GBL_objRegistry.DeleteKey HKEY_LOCAL_MACHINE,GBL_REG_UNINSTALL_KEY & ENH_NUM
	If Err.Number = 0 Then 
		SubLog "INFO:	Successfully deleted Registry Key - 'HKLM\" & GBL_REG_UNINSTALL_KEY & ENH_NUM & "'"
	Else
		SubLog "Error: Failed to delete registry key - 'HKLM\" & GBL_REG_UNINSTALL_KEY & ENH_NUM & "'"
		SubLog "Error: " & Err.Number & " - " & Err.Description
		subEndScript(RC_REGISTRY_ERROR)
	End If	
End sub
	
' *======================================================================
' * NAME     : subCreateActiveSetup
' * PURPOSE  : Confgures Active Setup	
' * INPUT    : N/A
' * OUTPUT   : N/A
' * RETURN   : N/A
' * Errors	 : N/A
' *======================================================================
Sub subCreateActiveSetup
	If Not DEBUG_ON Then 
	   On Error Resume Next
	   Err.Clear
	End If 
	
	subLog ""
	subLog "Creating Active Setup"
	GBL_WshShell.regwrite ACTIVE_SETUP_REG,"","REG_SZ" 
	'Update Version
	GBL_WshShell.regwrite ACTIVE_SETUP_REG & "Version",APP_VERSION,"REG_SZ"
	'Update Stub
	GBL_WshShell.regwrite ACTIVE_SETUP_REG & "StubPath",chr(34) & GBL_ScriptPath & "updateuserkey.vbs" & chr(34),"REG_SZ"
end sub

'=============================================================================
' NAME     : fnCloseApps
' PURPOSE  : Closes the Running Process
' INPUT    : N/A
' OUTPUT   : N/A
' RETURN   : N/A
'=============================================================================
Function fnCloseApps(ByVal Process)
	Dim colProcessList
	Dim objProcess
	On Error Resume Next
	fnCloseApps = False
	Set colProcessList = GBL_objWMIService.ExecQuery("Select * from Win32_Process WHERE Name = '" & Process & "'")
	For Each objProcess In colProcessList
		objProcess.Terminate()
		subLog "INFO        Application Terminated - '" & Process & "'"
		fnCloseApps = True
	Next
End Function

'=============================================================================
' NAME     : fnGetBIOSVersion
' PURPOSE  : Checks BIOS Version (WMI call).
' INPUT    : N/A
' OUTPUT   : GBL_intBIOSVersion = Right two numbers of BIOS version
' RETURN   : N/A
'=============================================================================
function fnGetBIOSVersion
		Set colBios = GBL_objWMIService.ExecQuery("Select * from Win32_BIOS")
        For Each objBios In colBios
            g_strCurrentBIOSLevel = objBios.SMBIOSBIOSVersion
		Next
	
	subLog "INFO:	BIOS Level is: " & g_strCurrentBIOSLevel
	GBL_IntBiosVersion = Right(g_strCurrentBIOSLevel, 2) '* Sets "GBL_intBIOSVersion" to equal the Right most two characters of the BIOS Version
	GBL_intBIOSVersion = CInt(GBL_intBIOSVersion)
	
end function



'=============================================================================
' NAME     : fnModelSelect
' PURPOSE  : Checks Computer Model and sets variables
' INPUT    : N/A
' OUTPUT   : N/A
' RETURN   : N/A
'=============================================================================
function fnModelSelect

'*------------------------------------------------------='
	'* Validates client model is a model listed for upgrade ='
	'*------------------------------------------------------='

		
			Set colItems = GBL_objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
			For Each objItem In colItems
				strModel = objItem.Model
				sublog "this is a " & strModel
			Next

	'*------------------------------------------------------------------------------='
	'* Add models to the case list below using all UPPERCASE letters                ='
	'* BiosVersionUpgradeCmd: the command used to run the BIOS.exe silent           ='
	'* BiosVersionRollbackCmd: the command used to run the BIOS used for Rollback   ='
	'* UpgradedBiosVersion: Enter the two right most Characters of the BIOS version ='
	'* being upgraded to. example "A10" = "10"                                      ='
	'*------------------------------------------------------------------------------='


	Select Case UCase(strModel)
		
		Case "HP ELITEDESK 800 G2 SFF"
			if X64BitHost = True then
				BiosVersionUpgradeCmd = "BIOS_exe\BIOS2.16\2_16-64.exe -s -a -b -r"
				BiosVersionRollbackCmd = "BIOS_exe\BIOS2.12\N\A -s -a -b -r"
			else
				BiosVersionUpgradeCmd = "BIOS_exe\BIOS2.16\2_16-32.exe -s -a -b -r"
				BiosVersionRollbackCmd = "BIOS_exe\BIOS2.12\N\A -s -a -b -r"
			end if
			UpgradedBiosVersion = 16 '* Enter the two right most Characters of the BIOS version. example "A10" = "10"
			sublog "selected ""HP EliteDesk 800 G2 SFF"" settings"
			strBiosCmd = BiosVersionUpgradeCmd
			If WScript.Arguments.Named.Exists("Rollback") Then strBiosCmd = BiosVersionRollbackCmd
			
		Case Else
			subLog "INFO:	Model is not in scope: " & strModel
			subLog "INFO:	Exiting script"
			SubEndScript (31)
	End Select
end function