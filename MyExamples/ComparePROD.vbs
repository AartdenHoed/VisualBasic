command = "powershell.exe -nologo -command C:\AdHC\PowerShell\ComparePROD.ps1"

set shell = CreateObject("WScript.Shell")

shell.Run command,0
