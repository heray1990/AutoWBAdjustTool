VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KONKAProtocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************
' Class module to handle protocal for KONKA.
'**********************************************

Option Explicit
Implements Protocal

Private mSendDataBuf As String

Private Sub Class_Initialize()
    mSendDataBuf = ""
End Sub

Private Sub SendCmd()
    FormMain.tcpServer.SendData mSendDataBuf
    
    SubDelayMs glngDelayTime
End Sub

Private Sub Protocal_EnterFacMode()
    'PASS
End Sub

Private Sub Protocal_ExitFacMode()
    mSendDataBuf = "{Method:[{""name"" : ""finishWBSettings""}]}"
    SendCmd
End Sub

Private Sub Protocal_SwitchInputSource(strInputSrc As String, intSrcNum As Integer)
    If strInputSrc = "HDMI" Then
        If intSrcNum = 1 Then
            mSendDataBuf = "{Method:[{""name"" : ""setInputSource"", ""param_value"" : ""HDMI""}]}"
        ElseIf intSrcNum = 2 Then
            mSendDataBuf = "{Method:[{""name"" : ""setInputSource"", ""param_value"" : ""HDMI2""}]}"
        Else
            mSendDataBuf = "{Method:[{""name"" : ""setInputSource"", ""param_value"" : ""HDMI""}]}"
        End If
    ElseIf strInputSrc = "AV" Then
        mSendDataBuf = "{Method:[{""name"" : ""setInputSource"", ""param_value"" : ""AV""}]}"
    Else
        mSendDataBuf = "{Method:[{""name"" : ""setInputSource"", ""param_value"" : ""HDMI""}]}"
    End If

    SendCmd
End Sub

'Set picture mode to standard.
Private Sub Protocal_ResetPicMode()
    mSendDataBuf = "{Method:[{""name"" : ""setBrightness"", ""param_value"" : ""50""}]}"
    SendCmd
    
    mSendDataBuf = "{Method:[{""name"" : ""setContrast"", ""param_value"" : ""50""}]}"
    SendCmd
    
    mSendDataBuf = "{Method:[{""name"" : ""setBacklight"", ""param_value"" : ""100""}]}"
    SendCmd
End Sub

Private Sub Protocal_SetBrightness(intBrightness As Integer)
    mSendDataBuf = "{Method:[{""name"" : ""setBrightness"", ""param_value"" : """ + CStr(intBrightness) + """}]}"
    SendCmd
End Sub

Private Sub Protocal_SetContrast(intContrast As Integer)
    mSendDataBuf = "{Method:[{""name"" : ""setContrast"", ""param_value"" : """ + CStr(intContrast) + """}]}"
    SendCmd
End Sub

Private Sub Protocal_SetBacklight(intBacklight As Integer)
    mSendDataBuf = "{Method:[{""name"" : ""setBacklight"", ""param_value"" : """ + CStr(intBacklight) + """}]}"
    SendCmd
End Sub

Private Sub Protocal_SelColorTemp(strColorT As String, strInputSrc As String, intSrcNum As Integer)
    Select Case strColorT
      Case COLORTEMP_COOL1
         mSendDataBuf = "{Method:[{""name"" : ""setWBMode"", ""param_value"" : ""cool""}]}"
      Case COLORTEMP_STANDARD
         mSendDataBuf = "{Method:[{""name"" : ""setWBMode"", ""param_value"" : ""normal""}]}"
      Case COLORTEMP_WARM1
         mSendDataBuf = "{Method:[{""name"" : ""setWBMode"", ""param_value"" : ""warm""}]}"
    End Select
    
    SendCmd
End Sub

Private Sub Protocal_SetRGBGain(lngRGain As Long, lngGGain As Long, lngBGain As Long)
    Call SetRGain(lngRGain)
    Call SetGGain(lngGGain)
    Call SetBGain(lngBGain)
End Sub

Private Sub SetRGain(lngRGain As Long)
    mSendDataBuf = "{Method:[{""name"" : ""setWBRedGain"", ""param_value"" : """ + CStr(Hex(lngRGain)) + """}]}"
    SendCmd
End Sub

Private Sub SetGGain(lngGGain As Long)
    mSendDataBuf = "{Method:[{""name"" : ""setWBGreenGain"", ""param_value"" : """ + CStr(Hex(lngGGain)) + """}]}"
    SendCmd
End Sub

Private Sub SetBGain(lngBGain As Long)
    mSendDataBuf = "{Method:[{""name"" : ""setWBBlueGain"", ""param_value"" : """ + CStr(Hex(lngBGain)) + """}]}"
    SendCmd
End Sub

Private Sub Protocal_SetRGBOffset(lngROffset As Long, lngGOffset As Long, lngBOffset As Long)
    Call SetROffset(lngROffset)
    Call SetGOffset(lngGOffset)
    Call SetBOffset(lngBOffset)
End Sub

Private Sub SetROffset(lngROffset As Long)
    mSendDataBuf = "{Method:[{""name"" : ""setWBRedOffset"", ""param_value"" : """ + CStr(Hex(lngROffset)) + """}]}"
    SendCmd
End Sub

Private Sub SetGOffset(lngGOffset As Long)
    mSendDataBuf = "{Method:[{""name"" : ""setWBGreenOffset"", ""param_value"" : """ + CStr(Hex(lngGOffset)) + """}]}"
    SendCmd
End Sub

Private Sub SetBOffset(lngBOffset As Long)
    mSendDataBuf = "{Method:[{""name"" : ""setWBBlueOffset"", ""param_value"" : """ + CStr(Hex(lngBOffset)) + """}]}"
    SendCmd
End Sub

Private Sub Protocal_SaveWBDataToAllSrc(strInputSrc As String, intSrcNum As Integer)
    mSendDataBuf = "{Method:[{""name"" : ""setCurrentWBToAllSources""}]}"
    SendCmd
End Sub

Public Sub Protocal_ChannelPreset()
    'PASS
End Sub


