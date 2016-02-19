VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommunicationProtocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************
' Interface for handling protocal.
'**********************************************

Option Explicit

Public Sub EnterFacMode()

End Sub

Public Sub ExitFacMode()

End Sub

Public Sub SwitchInputSource(strInputSrc As String, intSrcNum As Integer)

End Sub

Public Sub SetBrightness(intBrightness As Integer)

End Sub

Public Sub SetContrast(intContrast As Integer)

End Sub

Public Sub SetBacklight(intBacklight As Integer)

End Sub

Public Sub SetColorTemp(lngColorT As Long, strInputSrc As String, intSrcNum As Integer)

End Sub

Public Sub SetRGain(lngRGain As Long)

End Sub

Public Sub SetGGain(lngGGain As Long)

End Sub

Public Sub SetBGain(lngBGain As Long)

End Sub

Public Sub SetROffset(lngROffset As Long)

End Sub

Public Sub SetGOffset(lngGOffset As Long)

End Sub

Public Sub SetBOffset(lngBOffset As Long)

End Sub

Public Sub SaveWBDataToAllSrc(strInputSrc As String, intSrcNum As Integer)

End Sub
