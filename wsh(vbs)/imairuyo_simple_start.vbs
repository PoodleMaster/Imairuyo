Set objShell = CreateObject("WScript.Shell")

Do
    objShell.SendKeys "^{ESC}"
    WScript.Sleep 10000 ' 10 seconds
Loop
