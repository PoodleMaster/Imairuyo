' wscript.exe プロセスを終了する
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'wscript.exe'")
For Each objProcess in colProcesses
    objProcess.Terminate
Next
