Attribute VB_Name = "Module2"

Public Sub Delay(mmSec As Long)
On Error GoTo ShowError
    Dim start As Single
    start = Timer
    While (Timer - start) < (mmSec / 1000#)
        DoEvents

    If IsStop = True Then
        Exit Sub
    End If
    Wend
    Exit Sub

ShowError:
    MsgBox Err.Source & "------" & Err.Description
    Exit Sub
End Sub


Public Sub DelayMS(mmSec As Long)
On Error GoTo ShowError
    Dim start As Single
    start = Timer
    While (Timer - start) < (mmSec / 1000#)
        DoEvents
   
    If IsStop = True Then
        Exit Sub
    End If

    Wend
    Exit Sub

ShowError:
    MsgBox Err.Source & "------" & Err.Description
    Exit Sub
End Sub

Public Sub DelaySWithFlag(Sec As Long, flag As Boolean)
On Error GoTo ShowError
    Dim start As Single
    start = Timer
    While (Timer - start) < Sec
        DoEvents
   
        If flag = True Then
            Exit Sub
        End If
        
        If IsStop = True Then
            Exit Sub
        End If
    Wend
    Exit Sub

ShowError:
    MsgBox Err.Source & "------" & Err.Description
    Exit Sub
End Sub


