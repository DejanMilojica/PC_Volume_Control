Set WshShell = CreateObject("WScript.Shell") 
WshShell.Run chr(34) & "volume_control.bat" & Chr(34), 0
Set WshShell = Nothing