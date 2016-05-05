VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LetvCurvedProtocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************
' Only for Letv Max4-65-Curved.
'**********************************************

Option Explicit
Implements Protocal

Private mSendDataBuf(0 To 6) As Byte

Private Sub Class_Initialize()
    mSendDataBuf(0) = &H0
    mSendDataBuf(1) = &H0
    mSendDataBuf(2) = &H0
    mSendDataBuf(3) = &H0
    mSendDataBuf(4) = &H0
    mSendDataBuf(5) = &H0
    mSendDataBuf(6) = &H0
End Sub

Private Function i2clib_write(ByVal slave_addr_sft As Integer, ByRef cmd() As Byte, ByVal cmd_len As Long, ByRef dat() As Byte, ByVal data_len As Long) As Boolean
    Dim ack As Integer
    If (cmd_len = 0 And data_len = 0) Then
        I2c_Start
        ack = I2c_Tx(slave_addr_sft)
        I2c_Stop
            
        If ack = 0 Then
            i2clib_write = True
        Else
            i2clib_write = False
        End If
        GoTo Err
    End If
    i2clib_write = I2cWriteData(slave_addr_sft, cmd(0), cmd_len, dat(0), data_len)
Err:
End Function


Private Function SendCmd(ByRef data() As Byte) As Boolean
    Dim dummy(0) As Byte

    If i2clib_write(&H6E, dummy, 0, data(), 7) = False Then
        SendCmd = False
    Else
        SendCmd = True
    End If
    DelayMS delayTime
End Function

Private Function CalChkSum(ByRef data() As Byte) As Byte
    Dim i As Integer

    CalChkSum = &H0

    For i = 0 To 5
        CalChkSum = CalChkSum Xor data(i)
    Next i
End Function

Private Sub Protocal_EnterFacMode()
    '6E 51 84 03 FF 00 01 46
    mSendDataBuf(0) = &H51
    mSendDataBuf(1) = &H84
    mSendDataBuf(2) = &H3
    mSendDataBuf(3) = &HFF
    mSendDataBuf(4) = &H0
    mSendDataBuf(5) = &H1
    mSendDataBuf(6) = &H46

    If Not SendCmd(mSendDataBuf) Then
        SaveLogInFile "EnterFacMode Fail!"
    End If
End Sub

Private Sub Protocal_ExitFacMode()
End Sub

Private Sub Protocal_SwitchInputSource(strInputSrc As String, intSrcNum As Integer)
    'HDMI1: 6E 51 84 03 60 00 03 DB
    mSendDataBuf(0) = &H51
    mSendDataBuf(1) = &H84
    mSendDataBuf(2) = &H3
    mSendDataBuf(3) = &H60
    mSendDataBuf(4) = &H0

    If intSrcNum = 1 Then
        mSendDataBuf(5) = &H3
    ElseIf intSrcNum = 2 Then
        mSendDataBuf(5) = &H2
    ElseIf intSrcNum = 3 Then
        mSendDataBuf(5) = &H1
    Else
        mSendDataBuf(5) = &H3
    End If
    
    mSendDataBuf(6) = CalChkSum(mSendDataBuf) Xor &H6E
    
    If Not SendCmd(mSendDataBuf) Then
        SaveLogInFile "SwitchInputSource Fail!"
    End If
End Sub

'Set picture mode to standard.
Private Sub Protocal_ResetPicMode()
    Protocal_SetBrightness 50
    Protocal_SetContrast 50
    Protocal_SetBacklight 100
End Sub

Private Sub Protocal_SetBrightness(intBrightness As Integer)
    '6E 51 84 03 10 00 XX CHK
    mSendDataBuf(0) = &H51
    mSendDataBuf(1) = &H84
    mSendDataBuf(2) = &H3
    mSendDataBuf(3) = &H10
    mSendDataBuf(4) = &H0
    mSendDataBuf(5) = CByte(intBrightness Mod 256)
    mSendDataBuf(6) = CalChkSum(mSendDataBuf) Xor &H6E
    
    If Not SendCmd(mSendDataBuf) Then
        SaveLogInFile "SetBrightness" & CStr(intBrightness) & "Fail!"
    End If
End Sub

Private Sub Protocal_SetContrast(intContrast As Integer)
    '6E 51 84 03 12 00 XX CHK
    mSendDataBuf(0) = &H51
    mSendDataBuf(1) = &H84
    mSendDataBuf(2) = &H3
    mSendDataBuf(3) = &H12
    mSendDataBuf(4) = &H0
    mSendDataBuf(5) = CByte(intContrast Mod 256)
    mSendDataBuf(6) = CalChkSum(mSendDataBuf) Xor &H6E
    
    If Not SendCmd(mSendDataBuf) Then
        SaveLogInFile "SetContrast" & CStr(intContrast) & "Fail!"
    End If
End Sub

Private Sub Protocal_SetBacklight(intBacklight As Integer)
    '6E 51 84 03 13 00 XX CHK
    mSendDataBuf(0) = &H51
    mSendDataBuf(1) = &H84
    mSendDataBuf(2) = &H3
    mSendDataBuf(3) = &H13
    mSendDataBuf(4) = &H0
    mSendDataBuf(5) = CByte(intBacklight Mod 256)
    mSendDataBuf(6) = CalChkSum(mSendDataBuf) Xor &H6E
    
    If Not SendCmd(mSendDataBuf) Then
        SaveLogInFile "SetBacklight" & CStr(intBacklight) & "Fail!"
    End If
End Sub

Private Sub Protocal_SelColorTemp(strColorT As String, strInputSrc As String, intSrcNum As Integer)
    Select Case strColorT
      Case cstrColorTempCool1
         Call SetColorTempCool1
      Case cstrColorTempNormal
         Call SetColorTempNormal
      Case cstrColorTempWarm1
         Call SetColorTempWarm1
    End Select
End Sub

Private Sub SetColorTempCool1()
    'Cool
    '6E 51 84 03 14 00 02 AE
    mSendDataBuf(0) = &H51
    mSendDataBuf(1) = &H84
    mSendDataBuf(2) = &H3
    mSendDataBuf(3) = &H14
    mSendDataBuf(4) = &H0
    mSendDataBuf(5) = &H2
    mSendDataBuf(6) = &HAE
    
    If Not SendCmd(mSendDataBuf) Then
        SaveLogInFile "SetColorTempCool1 Fail!"
    End If
End Sub

Private Sub SetColorTempNormal()
    'normal
    '6E 51 84 03 14 00 01 AD
    mSendDataBuf(0) = &H51
    mSendDataBuf(1) = &H84
    mSendDataBuf(2) = &H3
    mSendDataBuf(3) = &H14
    mSendDataBuf(4) = &H0
    mSendDataBuf(5) = &H1
    mSendDataBuf(6) = &HAD
    
    If Not SendCmd(mSendDataBuf) Then
        SaveLogInFile "SetColorTempNormal Fail!"
    End If
End Sub

Private Sub SetColorTempWarm1()
    'warm
    '6E 51 84 03 14 00 00 AC
    mSendDataBuf(0) = &H51
    mSendDataBuf(1) = &H84
    mSendDataBuf(2) = &H3
    mSendDataBuf(3) = &H14
    mSendDataBuf(4) = &H0
    mSendDataBuf(5) = &H0
    mSendDataBuf(6) = &HAC
    
    If Not SendCmd(mSendDataBuf) Then
        SaveLogInFile "SetColorTempWarm1 Fail!"
    End If
End Sub

Private Sub Protocal_SetRGBGain(lngRGain As Long, lngGGain As Long, lngBGain As Long)
    Call SetRGain(lngRGain)
    Call SetGGain(lngGGain)
    Call SetBGain(lngBGain)
End Sub

Private Sub SetRGain(lngRGain As Long)
    '6E 51 84 03 16 00 XX CHK
    mSendDataBuf(0) = &H51
    mSendDataBuf(1) = &H84
    mSendDataBuf(2) = &H3
    mSendDataBuf(3) = &H16
    mSendDataBuf(4) = &H0
    mSendDataBuf(5) = CByte(lngRGain Mod 256)
    mSendDataBuf(6) = CalChkSum(mSendDataBuf) Xor &H6E
    
    If Not SendCmd(mSendDataBuf) Then
        SaveLogInFile "SetRGain" & CStr(lngRGain) & "Fail!"
    End If
End Sub

Private Sub SetGGain(lngGGain As Long)
    '6E 51 84 03 18 00 XX CHK
    mSendDataBuf(0) = &H51
    mSendDataBuf(1) = &H84
    mSendDataBuf(2) = &H3
    mSendDataBuf(3) = &H18
    mSendDataBuf(4) = &H0
    mSendDataBuf(5) = CByte(lngGGain Mod 256)
    mSendDataBuf(6) = CalChkSum(mSendDataBuf) Xor &H6E
    
    If Not SendCmd(mSendDataBuf) Then
        SaveLogInFile "SetGGain" & CStr(lngGGain) & "Fail!"
    End If
End Sub

Private Sub SetBGain(lngBGain As Long)
    '6E 51 84 03 1A 00 XX CHK
    mSendDataBuf(0) = &H51
    mSendDataBuf(1) = &H84
    mSendDataBuf(2) = &H3
    mSendDataBuf(3) = &H1A
    mSendDataBuf(4) = &H0
    mSendDataBuf(5) = CByte(lngBGain Mod 256)
    mSendDataBuf(6) = CalChkSum(mSendDataBuf) Xor &H6E
    
    If Not SendCmd(mSendDataBuf) Then
        SaveLogInFile "SetBGain" & CStr(lngBGain) & "Fail!"
    End If
End Sub

Private Sub Protocal_SetRGBOffset(lngROffset As Long, lngGOffset As Long, lngBOffset As Long)
    Call SetROffset(lngROffset)
    Call SetGOffset(lngGOffset)
    Call SetBOffset(lngBOffset)
End Sub

Private Sub SetROffset(lngROffset As Long)
    '6E 51 84 03 6C 00 XX CHK
    mSendDataBuf(0) = &H51
    mSendDataBuf(1) = &H84
    mSendDataBuf(2) = &H3
    mSendDataBuf(3) = &H6C
    mSendDataBuf(4) = &H0
    mSendDataBuf(5) = CByte(lngROffset Mod 256)
    mSendDataBuf(6) = CalChkSum(mSendDataBuf) Xor &H6E
    
    If Not SendCmd(mSendDataBuf) Then
        SaveLogInFile "SetROffset" & CStr(lngROffset) & "Fail!"
    End If
End Sub

Private Sub SetGOffset(lngGOffset As Long)
    '6E 51 84 03 6E 00 XX CHK
    mSendDataBuf(0) = &H51
    mSendDataBuf(1) = &H84
    mSendDataBuf(2) = &H3
    mSendDataBuf(3) = &H6E
    mSendDataBuf(4) = &H0
    mSendDataBuf(5) = CByte(lngGOffset Mod 256)
    mSendDataBuf(6) = CalChkSum(mSendDataBuf) Xor &H6E
    
    If Not SendCmd(mSendDataBuf) Then
        SaveLogInFile "SetGOffset" & CStr(lngGOffset) & "Fail!"
    End If
End Sub

Private Sub SetBOffset(lngBOffset As Long)
    '6E 51 84 03 70 00 XX CHK
    mSendDataBuf(0) = &H51
    mSendDataBuf(1) = &H84
    mSendDataBuf(2) = &H3
    mSendDataBuf(3) = &H70
    mSendDataBuf(4) = &H0
    mSendDataBuf(5) = CByte(lngBOffset Mod 256)
    mSendDataBuf(6) = CalChkSum(mSendDataBuf) Xor &H6E
    
    If Not SendCmd(mSendDataBuf) Then
        SaveLogInFile "SetBOffset" & CStr(lngBOffset) & "Fail!"
    End If
End Sub

Private Sub Protocal_SaveWBDataToAllSrc(strInputSrc As String, intSrcNum As Integer)
End Sub
