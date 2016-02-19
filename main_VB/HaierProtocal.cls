VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HaierProtocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************
' Class module to handle protocal for Haier and
' CIBN.
'**********************************************

Option Explicit
Implements CommunicationProtocal

Private mSendDataBuf(0 To 11) As Byte

Private Sub Class_Initialize()
    mSendDataBuf(0) = &H0
    mSendDataBuf(1) = &H0
    mSendDataBuf(2) = &H0
    mSendDataBuf(3) = &H0
    mSendDataBuf(4) = &H0
    mSendDataBuf(5) = &H0
    mSendDataBuf(6) = &H0
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0
    mSendDataBuf(10) = &H0
    mSendDataBuf(11) = &H0
End Sub

Private Sub SendCmd()
    If isUartMode Then
        Form1.MSComm1.Output = mSendDataBuf
    Else
        Form1.tcpClient.SendData mSendDataBuf
    End If
End Sub

Private Function CalChkSum(ByRef data() As Byte) As Byte
    'byte10 is the checksum byte.
    'byte10 = 0xFF - byte1 - byte2 - ... - byte9 + 1
    'If it hits 0xFF, ignore 0xFF and plus 1 instead.
    Dim i As Integer

    CalChkSum = &HFF

    For i = 1 To 9
        If data(i) = 255 Then
            CalChkSum = CalChkSum + 1
        Else
            If CalChkSum < data(i) Then
                CalChkSum = 256 - (data(i) - CalChkSum)
            Else
                CalChkSum = CalChkSum - data(i)
            End If
        End If
    Next i

    If CalChkSum = 255 Then
        CalChkSum = 0
    Else
        CalChkSum = CalChkSum + 1
    End If
End Function

Private Sub CommunicationProtocal_EnterFacMode()
    'Empty cmd for Haier
End Sub

Private Sub CommunicationProtocal_ExitFacMode()
    'Empty cmd for Haier
End Sub

Private Sub CommunicationProtocal_SwitchInputSource(strInputSrc As String, intSrcNum As Integer)
    '55 2E 01 XX 00 00 00 00 00 00 CHK FE
    'HDMI1: 55 2E 01 06 00 00 00 00 00 00 CB FE
    mSendDataBuf(0) = &H55
    mSendDataBuf(1) = &H2E
    mSendDataBuf(2) = &H1

    If strInputSrc = "HDMI" Then
        If intSrcNum = 1 Then
            mSendDataBuf(3) = &H6
        ElseIf intSrcNum = 2 Then
            mSendDataBuf(3) = &H7
        Else
            mSendDataBuf(3) = &H6
        End If
    ElseIf strInputSrc = "CVBS" Then
        If intSrcNum = 1 Then
            mSendDataBuf(3) = &H3
        ElseIf intSrcNum = 2 Then
            mSendDataBuf(3) = &H4
        Else
            mSendDataBuf(3) = &H3
        End If
    ElseIf strInputSrc = "YPbPr" Then
        mSendDataBuf(3) = &H5
    End If

    mSendDataBuf(4) = &H0
    mSendDataBuf(5) = &H0
    mSendDataBuf(6) = &H0
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)
    mSendDataBuf(11) = &HFE
    
    SendCmd
End Sub

Private Sub CommunicationProtocal_SetBrightness(intBrightness As Integer)
    '55 37 02 XX XX 00 00 00 00 00 CHK FE
    mSendDataBuf(0) = &H55
    mSendDataBuf(1) = &H37
    mSendDataBuf(2) = &H2
    mSendDataBuf(3) = CByte(intBrightness \ 256)
    mSendDataBuf(4) = CByte(intBrightness Mod 256)
    mSendDataBuf(5) = &H0
    mSendDataBuf(6) = &H0
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)
    mSendDataBuf(11) = &HFE
    
    SendCmd
End Sub

Private Sub CommunicationProtocal_SetContrast(intContrast As Integer)
    '55 39 02 XX XX 00 00 00 00 00 CHK FE
    mSendDataBuf(0) = &H55
    mSendDataBuf(1) = &H39
    mSendDataBuf(2) = &H2
    mSendDataBuf(3) = CByte(intContrast \ 256)
    mSendDataBuf(4) = CByte(intContrast Mod 256)
    mSendDataBuf(5) = &H0
    mSendDataBuf(6) = &H0
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)
    mSendDataBuf(11) = &HFE
    
    SendCmd
End Sub

Private Sub CommunicationProtocal_SetBacklight(intBacklight As Integer)
    '55 3B 02 XX XX 00 00 00 00 00 CHK FE
    mSendDataBuf(0) = &H55
    mSendDataBuf(1) = &H3B
    mSendDataBuf(2) = &H2
    mSendDataBuf(3) = CByte(intBacklight \ 256)
    mSendDataBuf(4) = CByte(intBacklight Mod 256)
    mSendDataBuf(5) = &H0
    mSendDataBuf(6) = &H0
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)
    mSendDataBuf(11) = &HFE
    
    SendCmd
End Sub

Private Sub CommunicationProtocal_SetColorTemp(lngColorT As Long, strInputSrc As String, intSrcNum As Integer)
    Dim value As Byte

    '55 02 01 XX 00 00 00 00 00 00 CHK FE
    Select Case lngColorT
        Case valColorTempCool1
            value = CByte(3)
        Case valColorTempNormal
            value = CByte(1)
        Case valColorTempWarm1
            value = CByte(2)
    End Select

    mSendDataBuf(0) = &H55
    mSendDataBuf(1) = &H2
    mSendDataBuf(2) = &H1
    mSendDataBuf(3) = value
    mSendDataBuf(4) = &H0
    mSendDataBuf(5) = &H0
    mSendDataBuf(6) = &H0
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)
    mSendDataBuf(11) = &HFE
    
    SendCmd
End Sub

Private Sub CommunicationProtocal_SetRGain(lngRGain As Long)
    '55 0A 02 XX XX 00 00 00 00 00 CHK FE
    mSendDataBuf(0) = &H55
    mSendDataBuf(1) = &HA
    mSendDataBuf(2) = &H2
    mSendDataBuf(3) = CByte(lngRGain \ 256)
    mSendDataBuf(4) = CByte(lngRGain Mod 256)
    mSendDataBuf(5) = &H0
    mSendDataBuf(6) = &H0
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)
    mSendDataBuf(11) = &HFE

    SendCmd
End Sub

Private Sub CommunicationProtocal_SetGGain(lngGGain As Long)
    '55 0B 02 XX XX 00 00 00 00 00 CHK FE
    mSendDataBuf(0) = &H55
    mSendDataBuf(1) = &HB
    mSendDataBuf(2) = &H2
    mSendDataBuf(3) = CByte(lngGGain \ 256)
    mSendDataBuf(4) = CByte(lngGGain Mod 256)
    mSendDataBuf(5) = &H0
    mSendDataBuf(6) = &H0
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)
    mSendDataBuf(11) = &HFE
    
    SendCmd
End Sub

Private Sub CommunicationProtocal_SetBGain(lngBGain As Long)
    '55 0C 02 XX XX 00 00 00 00 00 CHK FE
    mSendDataBuf(0) = &H55
    mSendDataBuf(1) = &HC
    mSendDataBuf(2) = &H2
    mSendDataBuf(3) = CByte(lngBGain \ 256)
    mSendDataBuf(4) = CByte(lngBGain Mod 256)
    mSendDataBuf(5) = &H0
    mSendDataBuf(6) = &H0
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)
    mSendDataBuf(11) = &HFE
    
    SendCmd
End Sub

Private Sub CommunicationProtocal_SetROffset(lngROffset As Long)
    '55 04 02 XX XX 00 00 00 00 00 CHK FE
    mSendDataBuf(0) = &H55
    mSendDataBuf(1) = &H4
    mSendDataBuf(2) = &H2
    mSendDataBuf(3) = CByte(lngROffset \ 256)
    mSendDataBuf(4) = CByte(lngROffset Mod 256)
    mSendDataBuf(5) = &H0
    mSendDataBuf(6) = &H0
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)
    mSendDataBuf(11) = &HFE
    
    SendCmd
End Sub

Private Sub CommunicationProtocal_SetGOffset(lngGOffset As Long)
    '55 05 02 XX XX 00 00 00 00 00 CHK FE
    mSendDataBuf(0) = &H55
    mSendDataBuf(1) = &H5
    mSendDataBuf(2) = &H2
    mSendDataBuf(3) = CByte(lngGOffset \ 256)
    mSendDataBuf(4) = CByte(lngGOffset Mod 256)
    mSendDataBuf(5) = &H0
    mSendDataBuf(6) = &H0
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)
    mSendDataBuf(11) = &HFE
    
    SendCmd
End Sub

Private Sub CommunicationProtocal_SetBOffset(lngBOffset As Long)
    '55 06 02 XX XX 00 00 00 00 00 CHK FE
    mSendDataBuf(0) = &H55
    mSendDataBuf(1) = &H6
    mSendDataBuf(2) = &H2
    mSendDataBuf(3) = CByte(lngBOffset \ 256)
    mSendDataBuf(4) = CByte(lngBOffset Mod 256)
    mSendDataBuf(5) = &H0
    mSendDataBuf(6) = &H0
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)
    mSendDataBuf(11) = &HFE
    
    SendCmd
End Sub

Private Sub CommunicationProtocal_SaveWBDataToAllSrc(strInputSrc As String, intSrcNum As Integer)
    '55 01 01 XX 00 00 00 00 00 00 CHK FE
    'HDMI1: 55 01 01 06 00 00 00 00 00 00 F8 FE
    mSendDataBuf(0) = &H55
    mSendDataBuf(1) = &H1
    mSendDataBuf(2) = &H1

    If strInputSrc = "HDMI" Then
        If intSrcNum = 1 Then
            mSendDataBuf(3) = &H6
        ElseIf intSrcNum = 2 Then
            mSendDataBuf(3) = &H7
        Else
            mSendDataBuf(3) = &H6
        End If
    ElseIf strInputSrc = "CVBS" Then
        If intSrcNum = 1 Then
            mSendDataBuf(3) = &H3
        ElseIf intSrcNum = 2 Then
            mSendDataBuf(3) = &H4
        Else
            mSendDataBuf(3) = &H3
        End If
    ElseIf strInputSrc = "YPbPr" Then
        mSendDataBuf(3) = &H5
    End If
    
    mSendDataBuf(4) = &H0
    mSendDataBuf(5) = &H0
    mSendDataBuf(6) = &H0
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)
    mSendDataBuf(11) = &HFE
    
    SendCmd
End Sub

