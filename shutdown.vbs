On Error Resume Next
'WScript.sleep 300000 'Delay 300 seconds
'Set WshShell = CreateObject("WScript.Shell")
'WshShell.Run "file.bat", , TRUE 'run batch file, Use your absolute path here
Dim OpSysSet, OpSys 
Set OpSysSet = GetObject("winmgmts:{(Shutdown)}//" & strComputer & "/root/cimv2").ExecQuery("select * from Win32_OperatingSystem where Primary=true") 
        For each OpSys in OpSysSet   
        opSys.Shutdown()   
        Next 
WScript.Quit