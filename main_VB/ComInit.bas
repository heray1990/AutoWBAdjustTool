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


Public Sub ComInit()
On Error GoTo ErrExit

    If Form1.MSComm1.PortOpen = True Then
        Form1.MSComm1.PortOpen = False
    End If
    
    With Form1
        .MSComm1.CommPort = setTVCurrentComID
        .MSComm1.Settings = setTVCurrentComBaud & ",N,8,1"
        .MSComm1.InputLen = 0
        
        .MSComm1.InBufferCount = 0
        .MSComm1.OutBufferCount = 0
        .MSComm1.InputMode = comInputModeBinary
        
        .MSComm1.NullDiscard = False
        .MSComm1.DTREnable = False
        .MSComm1.EOFEnable = False
        .MSComm1.RTSEnable = False
        .MSComm1.SThreshold = 1
        .MSComm1.RThreshold = 1
        .MSComm1.InBufferSize = 1024
        .MSComm1.OutBufferSize = 512
        
        .MSComm1.PortOpen = True
    End With
    
    Exit Sub

ErrExit:
        MsgBox Err.Description, vbCritical, Err.Source
End Sub


