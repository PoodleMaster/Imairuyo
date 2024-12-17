Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set objShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' ログファイルのパス
Dim logFilePath
logFilePath = fso.GetSpecialFolder(2) & "\imairuyo_start.log"

' ロックファイルのパス
Dim lockFilePath
lockFilePath = fso.GetSpecialFolder(2) & "\imairuyo_lock.lck"

' ロックファイルが存在する場合、すでに実行中とみなして終了
If fso.FileExists(lockFilePath) Then
    MsgBox "Another instance is already running.", vbExclamation, "Error"
    WScript.Quit
End If

' ロックファイルを作成して、他のインスタンスからの起動を防止
Set lockFile = fso.CreateTextFile(lockFilePath, True)
lockFile.Close

' 自身のプロセスIDを取得
Dim processID
processID = GetCurrentProcessID()

' 起動時にPIDをメッセージボックスに表示
MsgBox "imairuyo_start: PID = " & processID, vbInformation, "Process ID"

' ログファイルにプロセスIDを記録
Dim logFile
Set logFile = fso.CreateTextFile(logFilePath, True)
logFile.WriteLine "imairuyo_start: PID = " & processID
logFile.Close

' メインループ
Do
    objShell.SendKeys "^{ESC}"
    WScript.Sleep 10000 ' 10秒
Loop

' 現在のプロセスIDを取得する関数
Function GetCurrentProcessID()
    Dim colProcesses, objProcess
    Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'wscript.exe'")
    For Each objProcess In colProcesses
        ' 自身のスクリプト名で絞り込む
        If InStr(LCase(objProcess.CommandLine), LCase(WScript.ScriptFullName)) > 0 Then
            GetCurrentProcessID = objProcess.ProcessId
            Exit Function
        End If
    Next
End Function
