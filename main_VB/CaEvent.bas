Attribute VB_Name = "Module1"
Public ObjCa210 As Ca200
Public ObjCa As Ca
Public ObjProbe As Probe
Public ObjMemory As Memory


Public Sub CONNECT_CA210()
    Dim strERR As String
    Dim iReturn As Integer

On Error GoTo ER
    Set ObjCa210 = New Ca200
    sel = MsgBox("Please Set the Probe at 0-CALL Position! Are you sure?", vbYesNo + vbInformation, "Calibration")

    Select Case sel
        Case vbYes
            frmInformation.Show
            frmInformation.Infbox.Caption = "Please Wait,Initiating..."
      
            ObjCa210.AutoConnect
            Set ObjCa = ObjCa210.SingleCa
            Set ObjProbe = ObjCa.SingleProbe
            Set ObjMemory = ObjCa.Memory
       
            ObjCa.CalZero
            ObjCa.SyncMode = 3
            ObjCa.AveragingMode = 2
            ObjCa.SetAnalogRange 2.5, 2.5
            ObjCa.DisplayMode = 0
            ObjMemory.ChannelNO = 1
      
            frmInformation.Infbox.Caption = ""
            frmInformation.Hide
            Form1.ZOrder (0)
            Delay 200
    
            MsgBox "Please Set the Probe at Measure Position", vbOKOnly + vbInformation, "Calibration OK"
            IsCa210ok = True
        Case vbNo
    End Select
        
    Exit Sub

ER:
    strERR = "Error from " + Err.Source + Chr$(10) + Chr$(13)
    strERR = strERR + Err.Description + Chr$(10) + Chr$(13)
    strERR = strERR + "HRESULT" + CStr(Err.Number - vbObjectError)
    iReturn = MsgBox(strERR, vbRetryCancel)
    
    Select Case iReturn
        Case vbRetry: Resume
        Case Else:
        ObjCa.RemoteMode = 0
    End Select
End Sub

Public Sub Set22Channel()
      Set ObjMemory = ObjCa.Memory
      ObjMemory.ChannelNO = 6
End Sub
Public Sub Set26Channel()
      Set ObjMemory = ObjCa.Memory
      ObjMemory.ChannelNO = 4
End Sub
Public Sub Set32Channel()
      Set ObjMemory = ObjCa.Memory
      ObjMemory.ChannelNO = 5
End Sub
Public Sub Set22LEDChannel()
      Set ObjMemory = ObjCa.Memory
      ObjMemory.ChannelNO = 7
End Sub
Public Sub Set32LEDChannel()
      Set ObjMemory = ObjCa.Memory
      ObjMemory.ChannelNO = 8
End Sub
Public Sub Set40LEDChannel()
      Set ObjMemory = ObjCa.Memory
      ObjMemory.ChannelNO = 9
End Sub
Public Function GetX() As Single
    ObjCa.Measure
    GetX = ObjProbe.sx
End Function
Public Function GetY() As Single
    ObjCa.Measure
    GetY = ObjProbe.sy
End Function
Public Function GetLv() As Single
    ObjCa.Measure
    GetLv = ObjProbe.lv
End Function
Public Function GetGainX() As Single
    ObjCa.Measure
    GetGainX = ObjProbe.sx
End Function
Public Function GetGainY() As Single
    ObjCa.Measure
    GetGainY = ObjProbe.sy
End Function

Public Sub subGet_CA210_Value()
 On Error GoTo ER

    ObjCa.Measure
    CTBuff.x = ObjProbe.sx
    ObjCa.Measure
    CTBuff.y = ObjProbe.sy
    ObjCa.Measure
    CTBuff.lv = ObjProbe.lv
    
Exit Sub
ER:
    strERR = "Error from " + Err.Source + Chr$(10) + Chr$(13)
    strERR = strERR + Err.Description + Chr$(10) + Chr$(13)
    strERR = strERR + "HRESULT" + CStr(Err.Number - vbObjectError)
    iReturn = MsgBox(strERR, vbRetryCancel)
    Select Case iReturn
        Case vbRetry: Resume
        Case Else:
        ObjCa.RemoteMode = 0
    End Select
End Sub



