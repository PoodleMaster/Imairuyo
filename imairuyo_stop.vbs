Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set fso = CreateObject("Scripting.FileSystemObject")
Dim logFilePath, targetProcessID

' ログファイルのパス
logFilePath = fso.GetSpecialFolder(2) & "\imairuyo_start.log"

' ログファイルからプロセスIDを取得
If fso.FileExists(logFilePath) Then
    Dim logFile, line
    Set logFile = fso.OpenTextFile(logFilePath, 1)
    line = logFile.ReadLine
    logFile.Close
    targetProcessID = Split(line, ": ")(1) ' プロセスIDを取得
Else
    WScript.Echo "imairuyo_start is not running."
    WScript.Quit
End If

' WMIを使って該当プロセスを終了
Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId = " & targetProcessID)
For Each objProcess In colProcesses
    objProcess.Terminate
Next
