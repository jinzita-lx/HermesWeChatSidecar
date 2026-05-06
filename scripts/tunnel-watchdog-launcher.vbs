' tunnel-watchdog-launcher.vbs
' Used by Task Scheduler to launch the SSH tunnel watchdog truly hidden
' (no PowerShell window flash on logon).

Option Explicit
Dim sh, ps1, cmd
Set sh = CreateObject("WScript.Shell")

ps1 = "C:\HermesWeChatSidecar\scripts\tunnel-watchdog.ps1"
cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & ps1 & """"
' 0 = hidden, False = don't wait
sh.Run cmd, 0, False
