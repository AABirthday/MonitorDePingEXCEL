Function Ping(strip)
Dim objshell, boolcode
Set objshell = CreateObject("Wscript.Shell")
boolcode = objshell.Run("ping -n 1 -w 1000 " & strip, 0, True)
If boolcode = 0 Then
    Ping = True
Else
    Ping = False
End If
End Function

Sub PingSystem()
Dim strip As String
Do Until Hoja1.Range("G9").Value = "STOP"
Hoja1.Range("G9").Value = "TESTING"
For introw = 2 To ActiveSheet.Cells(65536, 2).End(xlUp).Row
    strip = ActiveSheet.Cells(introw, 2).Value
    If Ping(strip) = True Then
        ActiveSheet.Cells(introw, 3).Value = "Online"
        ActiveSheet.Cells(introw, 3).Font.Color = RGB(0, 0, 0)
        Application.Wait (Now + TimeValue("0:00:01"))
        ActiveSheet.Cells(introw, 3).Font.Color = RGB(0, 200, 0)
    Else
        ActiveSheet.Cells(introw, 3).Value = "Offline"
        ActiveSheet.Cells(introw, 3).Interior.ColorIndex = 0
        ActiveSheet.Cells(introw, 3).Font.Color = RGB(200, 0, 0)
        Application.Wait (Now + TimeValue("0:00:01"))
        ActiveSheet.Cells(introw, 3).Interior.ColorIndex = 6
    End If
    If Hoja1.Range("G9").Value = "STOP" Then
        Exit For
    End If
Next
Loop
Hoja1.Range("G9").Value = "IDLE"
End Sub

Sub stop_ping()
    Hoja1.Range("G9").Value = "STOP"
End Sub
