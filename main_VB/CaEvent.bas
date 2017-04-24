Attribute VB_Name = "CaEvent"
Option Explicit

Public ObjCa210 As Ca200
Public ObjCa As Ca
Public ObjProbe As Probe
Public ObjMemory As Memory


Public Sub SubConnectCa()
    On Error GoTo ER
    Dim strERR As String
    Dim iReturn As Integer

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
