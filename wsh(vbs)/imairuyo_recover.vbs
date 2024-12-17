' wscript.exe プロセスを終了する 
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set fso = CreateObject("Scripting.FileSystemObject")

' ロックファイルのパス
Dim lockFilePath
lockFilePath = fso.GetSpecialFolder(2) & "\imairuyo_lock.lck"

' WScript.exe プロセスをすべて終了（自身のプロセスは除外）
MsgBox "(1) Terminated process start", vbInformation, "Check point"
Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'wscript.exe'")
For Each objProcess In colProcesses
    ' 自身のプロセスを除外
    If objProcess.ProcessId <> GetCurrentProcessID() Then
        objProcess.Terminate
        MsgBox "Terminated process with PID: " & objProcess.ProcessId, vbInformation, "Process Terminated"
    End If
Next

' プロセス終了待機
WScript.Sleep 2000 ' 2秒待機して完全終了を確認
MsgBox "(1) Terminated process end", vbInformation, "Check point"

' ロックファイルを削除
MsgBox "(2) Lock file deleted process start", vbInformation, "Check point"
If fso.FileExists(lockFilePath) Then
    fso.DeleteFile(lockFilePath)
    MsgBox "(2) Lock file deleted.", vbInformation, "Lock File Deleted"
Else
    MsgBox "(2) Lock file not found.", vbExclamation, "Lock File Missing"
End If

' 最後に自身のプロセスを終了
MsgBox "Recovery completed. Exiting the script.", vbInformation, "Recovery Complete"

' 自身のプロセスを終了
WScript.Quit


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
