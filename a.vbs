Set objShell = CreateObject("WScript.Shell")

Do
    MoveMouse 10, 10 ' Move the mouse cursor by 10 pixels horizontally and vertically
    WScript.Sleep 30000 ' Sleep for 30 seconds
Loop

Sub MoveMouse(xDelta, yDelta)
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "rundll32.exe user32.dll,SetCursorPos " & (objShell.RegRead("HKEY_CURRENT_USER\Control Panel\Mouse\MouseSpeed") + xDelta) & " " & (objShell.RegRead("HKEY_CURRENT_USER\Control Panel\Mouse\MouseSpeed") + yDelta)
End Sub
