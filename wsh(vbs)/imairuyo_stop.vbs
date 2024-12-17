Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set fso = CreateObject("Scripting.FileSystemObject")
Dim logFilePath, lockFilePath, targetProcessID

' ログファイルとロックファイルのパス
logFilePath = fso.GetSpecialFolder(2) & "\imairuyo_start.log"
lockFilePath = fso.GetSpecialFolder(2) & "\imairuyo_lock.lck"

' ログファイルからPIDを取得
If fso.FileExists(logFilePath) Then
    Dim logFile, line, tempPID
    Set logFile = fso.OpenTextFile(logFilePath, 1)
    line = Trim(logFile.ReadLine) ' 余分な空白・改行を削除
    logFile.Close

    ' PIDを抽出
    If InStr(line, "PID = ") > 0 Then
        tempPID = Trim(Split(line, "PID = ")(1)) ' PID部分を取り出す
        If IsNumeric(tempPID) Then
            targetProcessID = CLng(tempPID) ' 数値に変換
            MsgBox "Extracted PID: " & targetProcessID, vbInformation, "PID Retrieved"
        Else
            MsgBox "PID is not numeric: " & tempPID, vbCritical, "Error"
            WScript.Quit
        End If
    Else
        MsgBox "Log file format invalid.", vbCritical, "Error"
        WScript.Quit
    End If
Else
    MsgBox "Log file not found.", vbCritical, "Error"
    WScript.Quit
End If

' WMIを使って該当プロセスを終了
On Error Resume Next
Dim colProcesses, objProcess
Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId = " & targetProcessID)

Dim processFound
processFound = False

For Each objProcess In colProcesses
    objProcess.Terminate
    processFound = True
    MsgBox "Terminated process with PID: " & targetProcessID, vbInformation, "Success"
Next

If Not processFound Then
    MsgBox "Process with PID " & targetProcessID & " not found.", vbExclamation, "Not Found"
End If
On Error GoTo 0

' プロセス終了待機
WScript.Sleep 2000 ' 2秒待機して完全終了を確認

' ロックファイルを削除
If fso.FileExists(lockFilePath) Then
    fso.DeleteFile(lockFilePath)
    MsgBox "Lock file deleted.", vbInformation, "Clean Up"
End If
