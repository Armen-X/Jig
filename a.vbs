Set objShell = CreateObject("WScript.Shell")

Do
    ' Move the mouse cursor by a minimal amount to prevent Skype from going idle
    MoveMouse(1, 1) ' Move the mouse cursor by 1 pixel horizontally and vertically
    WScript.Sleep 30000 ' Sleep for 30 seconds
Loop

Sub MoveMouse(xDelta, yDelta)
    Set objShell = CreateObject("WScript.Shell")
    currentX = objShell.RegRead("HKEY_CURRENT_USER\Control Panel\Mouse\MouseSpeed")
    
    ' Calculate new mouse coordinates (keep it at the same position)
    newX = currentX
    newY = currentX
    
    objShell.Run "rundll32.exe user32.dll,SetCursorPos " & newX & " " & newY
End Sub
