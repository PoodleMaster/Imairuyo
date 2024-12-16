Add-Type -AssemblyName System.Windows.Forms

[Console]::Write("Running...")

while ($true) {
    [System.Windows.Forms.SendKeys]::SendWait('^{ESC}')
    Start-Sleep -Seconds 10
}
