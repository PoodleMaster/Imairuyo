Add-Type -AssemblyName System.Windows.Forms

$prefix = "Running... :"
$textLength = [System.Text.Encoding]::UTF8.GetByteCount($prefix)

[Console]::Clear()
[Console]::ForegroundColor = [ConsoleColor]::Yellow
[Console]::Write($prefix)

while ($true) {
    [System.Windows.Forms.SendKeys]::SendWait('+')
    $datetime = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
    [Console]::ForegroundColor = [ConsoleColor]::White
    [Console]::Write($datetime)
    Start-Sleep -Seconds 1
    [System.Console]::SetCursorPosition($textLength, 0)
}
