Set objShell = CreateObject("WScript.Shell")
Randomize ' Initialize the random number generator

Do
    ' Generate random coordinates within a range (adjust as needed)
    xDelta = Int((20 * Rnd) + 1) ' Random value between 1 and 20
    yDelta = Int((20 * Rnd) + 1) ' Random value between 1 and 20

    MoveMouse xDelta, yDelta ' Move the mouse cursor by the random values
    WScript.Sleep 30000 ' Sleep for 30 seconds
Loop

Sub MoveMouse(xDelta, yDelta)
    Set objShell = CreateObject("WScript.Shell")
    currentX = objShell.RegRead("HKEY_CURRENT_USER\Control Panel\Mouse\MouseSpeed")
    
    ' Calculate new mouse coordinates
    newX = currentX + xDelta
    newY = currentX + yDelta
    
    objShell.Run "rundll32.exe user32.dll,SetCursorPos " & newX & " " & newY
End Sub
