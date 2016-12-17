VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LetvMST6M60"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************
' Class module to handle protocal for MST6M60
' of Letv.
'**********************************************

Option Explicit
Implements Protocal

Private mSendDataBuf(0 To 10) As Byte
Private mDDCDataWithoutChksum(0 To 5) As Byte

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
    mDDCDataWithoutChksum(0) = &H0
    mDDCDataWithoutChksum(1) = &H0
    mDDCDataWithoutChksum(2) = &H0
    mDDCDataWithoutChksum(3) = &H0
    mDDCDataWithoutChksum(4) = &H0
    mDDCDataWithoutChksum(5) = &H0
End Sub

Private Sub SendCmd()
    Form1.MSComm1.Output = mSendDataBuf
    
    DelayMS delayTime
End Sub

Private Function CalDDCChkSum(ByRef data() As Byte) As Byte
    Dim i As Integer
    Dim tmp As Integer

    tmp = 0
    CalDDCChkSum = &H0

    For i = 0 To 5
        tmp = tmp + data(i)
    Next i
    
    CalDDCChkSum = tmp And &HF
End Function

Private Function CalChkSum(ByRef data() As Byte) As Byte
    Dim i As Integer
    Dim tmp As Integer

    tmp = 0
    CalChkSum = &H0

    For i = 0 To 9
        tmp = tmp + data(i)
    Next i
    
    CalChkSum = &HFF - tmp And &HFF
End Function

Private Sub DataToDDC()
    For i = 0 To 5
        mDDCDataWithoutChksum(i) = mSendDataBuf(i + 4)
    Next i
End Sub

Private Sub Protocal_EnterFacMode()
    'Open PQ
    'E0 0B 40 53 24 00 01 00 00 00 5C
    mSendDataBuf(0) = &HE0
    mSendDataBuf(1) = &HB
    mSendDataBuf(2) = &H40
    mSendDataBuf(4) = &H24
    mSendDataBuf(5) = &H0
    mSendDataBuf(6) = &H1
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0

    DataToDDC
    mSendDataBuf(3) = CalDDCChkSum(mDDCDataWithoutChksum) * 16 + &H3
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)

    SendCmd
End Sub

Private Sub Protocal_ExitFacMode()
    'The last command should delay 2s to save the setting values to TV.
    DelayMS 2000
End Sub

Private Sub Protocal_SwitchInputSource(strInputSrc As String, intSrcNum As Integer)
    'Change resolution to 1080P
    'E0 0B 40 81 05 07 80 04 38 00 8B
    mSendDataBuf(0) = &HE0
    mSendDataBuf(1) = &HB
    mSendDataBuf(2) = &H40
    mSendDataBuf(4) = &H5
    mSendDataBuf(5) = &H7
    mSendDataBuf(6) = &H80
    mSendDataBuf(7) = &H4
    mSendDataBuf(8) = &H38
    mSendDataBuf(9) = &H0

    DataToDDC
    mSendDataBuf(3) = CalDDCChkSum(mDDCDataWithoutChksum) * 16 + &H1
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)

    SendCmd
End Sub

'Set picture mode to standard.
Private Sub Protocal_ResetPicMode()
End Sub

Private Sub Protocal_SetBrightness(intBrightness As Integer)
End Sub

Private Sub Protocal_SetContrast(intContrast As Integer)
End Sub

Private Sub Protocal_SetBacklight(intBacklight As Integer)
End Sub

Private Sub Protocal_SelColorTemp(strColorT As String, strInputSrc As String, intSrcNum As Integer)
    'Cool:   E0 0B 40 6D 06 00 00 00 00 00 61
    'Normal: E0 0B 40 7D 06 01 00 00 00 00 50
    'Warm:   E0 0B 40 8D 06 02 00 00 00 00 3F
    mSendDataBuf(0) = &HE0
    mSendDataBuf(1) = &HB
    mSendDataBuf(2) = &H40
    mSendDataBuf(4) = &H6
    mSendDataBuf(6) = &H0
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0

    Select Case strColorT
      Case cstrColorTempCool1
         mSendDataBuf(5) = &H0
      Case cstrColorTempNormal
         mSendDataBuf(5) = &H1
      Case cstrColorTempWarm1
         mSendDataBuf(5) = &H2
    End Select
    
    DataToDDC
    mSendDataBuf(3) = CalDDCChkSum(mDDCDataWithoutChksum) * 16 + &HD
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)
    
    SendCmd
End Sub

Private Sub Protocal_SetRGBGain(lngRGain As Long, lngGGain As Long, lngBGain As Long)
    Call SetRGain(lngRGain)
    Call SetGGain(lngGGain)
    Call SetBGain(lngBGain)
End Sub

Private Sub SetRGain(lngRGain As Long)
    'E0 0B 40 X8 06 00 XX XX 00 00 CHK
    mSendDataBuf(0) = &HE0
    mSendDataBuf(1) = &HB
    mSendDataBuf(2) = &H40
    mSendDataBuf(4) = &H6
    mSendDataBuf(5) = &H0
    mSendDataBuf(6) = CByte(lngRGain Mod 256)
    mSendDataBuf(7) = CByte(lngRGain \ 256)
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0

    DataToDDC
    mSendDataBuf(3) = CalDDCChkSum(mDDCDataWithoutChksum) * 16 + &H8
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)

    SendCmd
End Sub

Private Sub SetGGain(lngGGain As Long)
    'E0 0B 40 X8 06 02 XX XX 00 00 CHK
    mSendDataBuf(0) = &HE0
    mSendDataBuf(1) = &HB
    mSendDataBuf(2) = &H40
    mSendDataBuf(4) = &H6
    mSendDataBuf(5) = &H2
    mSendDataBuf(6) = CByte(lngGGain Mod 256)
    mSendDataBuf(7) = CByte(lngGGain \ 256)
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0

    DataToDDC
    mSendDataBuf(3) = CalDDCChkSum(mDDCDataWithoutChksum) * 16 + &H8
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)

    SendCmd
End Sub

Private Sub SetBGain(lngBGain As Long)
    'E0 0B 40 X8 06 01 XX XX 00 00 CHK
    mSendDataBuf(0) = &HE0
    mSendDataBuf(1) = &HB
    mSendDataBuf(2) = &H40
    mSendDataBuf(4) = &H6
    mSendDataBuf(5) = &H1
    mSendDataBuf(6) = CByte(lngBGain Mod 256)
    mSendDataBuf(7) = CByte(lngBGain \ 256)
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0

    DataToDDC
    mSendDataBuf(3) = CalDDCChkSum(mDDCDataWithoutChksum) * 16 + &H8
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)

    SendCmd
End Sub

Private Sub Protocal_SetRGBOffset(lngROffset As Long, lngGOffset As Long, lngBOffset As Long)
    Call SetROffset(lngROffset)
    Call SetGOffset(lngGOffset)
    Call SetBOffset(lngBOffset)
End Sub

Private Sub SetROffset(lngROffset As Long)
    'E0 0B 40 X8 06 03 XX XX 00 00 CHK
    mSendDataBuf(0) = &HE0
    mSendDataBuf(1) = &HB
    mSendDataBuf(2) = &H40
    mSendDataBuf(4) = &H6
    mSendDataBuf(5) = &H3
    mSendDataBuf(6) = CByte(lngROffset Mod 256)
    mSendDataBuf(7) = CByte(lngROffset \ 256)
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0

    DataToDDC
    mSendDataBuf(3) = CalDDCChkSum(mDDCDataWithoutChksum) * 16 + &H8
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)

    SendCmd
End Sub

Private Sub SetGOffset(lngGOffset As Long)
    'E0 0B 40 X8 06 05 XX XX 00 00 CHK
    mSendDataBuf(0) = &HE0
    mSendDataBuf(1) = &HB
    mSendDataBuf(2) = &H40
    mSendDataBuf(4) = &H6
    mSendDataBuf(5) = &H5
    mSendDataBuf(6) = CByte(lngGOffset Mod 256)
    mSendDataBuf(7) = CByte(lngGOffset \ 256)
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0

    DataToDDC
    mSendDataBuf(3) = CalDDCChkSum(mDDCDataWithoutChksum) * 16 + &H8
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)

    SendCmd
End Sub

Private Sub SetBOffset(lngBOffset As Long)
    'E0 0B 40 X8 06 04 XX XX 00 00 CHK
    mSendDataBuf(0) = &HE0
    mSendDataBuf(1) = &HB
    mSendDataBuf(2) = &H40
    mSendDataBuf(4) = &H6
    mSendDataBuf(5) = &H4
    mSendDataBuf(6) = CByte(lngBOffset Mod 256)
    mSendDataBuf(7) = CByte(lngBOffset \ 256)
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0

    DataToDDC
    mSendDataBuf(3) = CalDDCChkSum(mDDCDataWithoutChksum) * 16 + &H8
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)

    SendCmd
End Sub

Private Sub Protocal_SaveWBDataToAllSrc(strInputSrc As String, intSrcNum As Integer)
    'E0 0B 40 X8 07 00 00 00 00 00 CHK
    mSendDataBuf(0) = &HE0
    mSendDataBuf(1) = &HB
    mSendDataBuf(2) = &H40
    mSendDataBuf(4) = &H7
    mSendDataBuf(5) = &H0
    mSendDataBuf(6) = &H0
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0

    DataToDDC
    mSendDataBuf(3) = CalDDCChkSum(mDDCDataWithoutChksum) * 16 + &H8
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)

    SendCmd
End Sub

Public Sub Protocal_ChannelPreset()
End Sub


