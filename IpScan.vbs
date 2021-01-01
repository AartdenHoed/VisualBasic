command = "powershell.exe -nologo -command C:\AdHC\PowerShell\IpScan.ps1"

set shell = CreateObject("WScript.Shell")

shell.Run command,0
