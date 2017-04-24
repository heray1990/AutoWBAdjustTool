Attribute VB_Name = "CaEvent"
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
            FormCalZero.Show
            DoEvents

            ObjCa210.AutoConnect
            Set ObjCa = ObjCa210.SingleCa
            Set ObjProbe = ObjCa.SingleProbe
            Set ObjMemory = ObjCa.Memory
       
            ObjCa.CalZero
            ObjCa.SyncMode = 3
            ObjCa.AveragingMode = 2
            ObjCa.SetAnalogRange 2.5, 2.5
            ObjCa.DisplayMode = 0
            ObjMemory.ChannelNO = glngCaChannel
    
            MsgBox "Please Set the Probe at Measure Position", vbOKOnly + vbInformation, "Calibration OK"
            gblCaConnected = True
            FormCalZero.Hide
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



