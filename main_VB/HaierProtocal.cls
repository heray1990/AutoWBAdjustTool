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
' Class module to handle protocal for Haier.
' When connecting T111, you need to turn "Uart bus" off in T111's factory menu.
'**********************************************

Option Explicit
Implements Protocal

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
    If gutdCommMode = modeUART Then
        FormMain.MSComm1.Output = mSendDataBuf
    ElseIf gutdCommMode = modeNetwork Then
        FormMain.tcpClient.SendData mSendDataBuf
    End If
    
    SubDelayMs glngDelayTime
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

Private Sub Protocal_EnterFacMode()
    'Empty cmd for Haier
End Sub

Private Sub Protocal_ExitFacMode()
    'Empty cmd for Haier
End Sub

Private Sub Protocal_SwitchInputSource(strInputSrc As String, intSrcNum As Integer)
    '55 2E 01 XX 00 00 00 00 00 00 CHK FE
    'HDMI1: 55 2E 01 07 00 00 00 00 00 00 CA FE
    mSendDataBuf(0) = &H55
    mSendDataBuf(1) = &H2E
    mSendDataBuf(2) = &H1

    If strInputSrc = "HDMI" Then
        If intSrcNum = 1 Then
            mSendDataBuf(3) = &H7
        ElseIf intSrcNum = 2 Then
            mSendDataBuf(3) = &H8
        ElseIf intSrcNum = 3 Then
            mSendDataBuf(3) = &H9
        Else
            mSendDataBuf(3) = &H7
        End If
    ElseIf strInputSrc = "AV" Then
        mSendDataBuf(3) = &H3
    ElseIf strInputSrc = "YPbPr" Then
        mSendDataBuf(3) = &H5
    Else
        mSendDataBuf(3) = &H7
    End If

    mSendDataBuf(4) = &H0
    mSendDataBuf(5) = &H0
    mSendDataBuf(6) = &H0
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)
    mSendDataBuf(11) = &HFE
    
    If Not (gstrChipSet = "T111") Then
        SendCmd
    End If
End Sub

'Set picture mode to standard.
Private Sub Protocal_ResetPicMode()
    '55 34 00 00 00 00 00 00 00 00 CC FE
    mSendDataBuf(0) = &H55
    mSendDataBuf(1) = &H34
    mSendDataBuf(2) = &H0
    mSendDataBuf(3) = &H0
    mSendDataBuf(4) = &H0
    mSendDataBuf(5) = &H0
    mSendDataBuf(6) = &H0
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0
    mSendDataBuf(10) = &HCC
    mSendDataBuf(11) = &HFE
    
    If Not (gstrChipSet = "T111") Then
        SendCmd
    End If
End Sub

Private Sub Protocal_SetBrightness(intBrightness As Integer)
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
    
    If Not (gstrChipSet = "T111") Then
        SendCmd
    End If
End Sub

Private Sub Protocal_SetContrast(intContrast As Integer)
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
    
    If Not (gstrChipSet = "T111") Then
        SendCmd
    End If
End Sub

Private Sub Protocal_SetBacklight(intBacklight As Integer)
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
    
    If Not (gstrChipSet = "T111") Then
        SendCmd
    End If
End Sub

Private Sub Protocal_SelColorTemp(strColorT As String, strInputSrc As String, intSrcNum As Integer)
    '55 02 01 XX 00 00 00 00 00 00 CHK FE

    mSendDataBuf(0) = &H55
    mSendDataBuf(1) = &H2
    mSendDataBuf(2) = &H1

    If (gstrChipSet = "T111") Then
        Select Case strColorT
            Case COLORTEMP_COOL1
                mSendDataBuf(3) = &H3
            Case COLORTEMP_STANDARD
                mSendDataBuf(3) = &H1
            Case COLORTEMP_WARM1
                mSendDataBuf(3) = &H2
        End Select
    Else
        Select Case strColorT
            Case COLORTEMP_COOL1
                mSendDataBuf(3) = &H0
            Case COLORTEMP_STANDARD
                mSendDataBuf(3) = &H1
            Case COLORTEMP_WARM1
                mSendDataBuf(3) = &H2
        End Select
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

Private Sub Protocal_SetRGBGain(lngRGain As Long, lngGGain As Long, lngBGain As Long)
    '55 0A 06 XX XX XX XX XX XX 00 CHK FE
    mSendDataBuf(0) = &H55
    mSendDataBuf(1) = &HA
    mSendDataBuf(2) = &H6
    mSendDataBuf(3) = CByte(lngRGain \ 256)
    mSendDataBuf(4) = CByte(lngRGain Mod 256)
    mSendDataBuf(5) = CByte(lngGGain \ 256)
    mSendDataBuf(6) = CByte(lngGGain Mod 256)
    mSendDataBuf(7) = CByte(lngBGain \ 256)
    mSendDataBuf(8) = CByte(lngBGain Mod 256)
    mSendDataBuf(9) = &H0
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)
    mSendDataBuf(11) = &HFE

    SendCmd
End Sub

Private Sub Protocal_SetRGBOffset(lngROffset As Long, lngGOffset As Long, lngBOffset As Long)
    '55 04 06 XX XX XX XX XX XX 00 CHK FE
    mSendDataBuf(0) = &H55
    mSendDataBuf(1) = &H4
    mSendDataBuf(2) = &H6
    If (gstrChipSet = "T111") Then
        mSendDataBuf(3) = CByte((lngROffset) \ 256)
        mSendDataBuf(4) = CByte((lngROffset) Mod 256)
        mSendDataBuf(5) = CByte((lngGOffset) \ 256)
        mSendDataBuf(6) = CByte((lngGOffset) Mod 256)
        mSendDataBuf(7) = CByte((lngBOffset) \ 256)
        mSendDataBuf(8) = CByte((lngBOffset) Mod 256)
    Else
        mSendDataBuf(3) = CByte((lngROffset + 1024) \ 256)
        mSendDataBuf(4) = CByte((lngROffset + 1024) Mod 256)
        mSendDataBuf(5) = CByte((lngGOffset + 1024) \ 256)
        mSendDataBuf(6) = CByte((lngGOffset + 1024) Mod 256)
        mSendDataBuf(7) = CByte((lngBOffset + 1024) \ 256)
        mSendDataBuf(8) = CByte((lngBOffset + 1024) Mod 256)
    End If
    mSendDataBuf(9) = &H0
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)
    mSendDataBuf(11) = &HFE

    SendCmd
End Sub

Private Sub Protocal_SaveWBDataToAllSrc(strInputSrc As String, intSrcNum As Integer)
    '55 01 01 06 00 00 00 00 00 00 F8 FE
    mSendDataBuf(0) = &H55
    mSendDataBuf(1) = &H1
    mSendDataBuf(2) = &H1
    mSendDataBuf(3) = &H6
    mSendDataBuf(4) = &H0
    mSendDataBuf(5) = &H0
    mSendDataBuf(6) = &H0
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)
    mSendDataBuf(11) = &HFE
    
    If Not (gstrChipSet = "T111") Then
        SendCmd
    End If
End Sub

Public Sub Protocal_ChannelPreset()

End Sub

