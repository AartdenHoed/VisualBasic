
Set oShell = CreateObject("Shell.Application")
oShell.ShellExecute "powershell.exe", "-command C:\AdHC\PowerShell\DeployExec.ps1", , "runas", 0

