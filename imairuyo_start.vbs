Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set objShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' ログファイルのパス
Dim logFilePath
logFilePath = fso.GetSpecialFolder(2) & "\imairuyo_start.log"

' 以下のフォルダにimairuyoのPIDを格納します。
' PIDの確認方法は、以下のコマンドで確認できます。
' cd %TEMP%
' type imairuyo_start.log

' 自身のプロセスIDを取得
Dim processID
processID = GetCurrentProcessID()

' ログファイルにプロセスIDを記録
Dim logFile
Set logFile = fso.CreateTextFile(logFilePath, True)
logFile.WriteLine "imairuyo_start: " & processID
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
