Attribute VB_Name = "Module5"
Option Explicit
'====================================================
'====================================================


Public Declare Function initColorTemp Lib "ColorT.dll" (ByRef Timming As Long, ByRef Pattern As Long, ByRef pMaxLv As Long, ByRef pMinLv As Long, ByRef Calibration As Long, ByRef MinBri As Long, ByVal ModelFile As String, ByVal pCurDir As String) As Long
Public Declare Function DeinitColorTemp Lib "ColorT.dll" (ByVal ModelFile As String) As Long
Public Declare Function setColorTemp Lib "ColorT.dll" (ByVal colorT As Long, ByRef pCOLORs As COLORTEMPSPEC, ByVal refHighLowMode As Long) As Long
Public Declare Function checkColorTemp Lib "ColorT.dll" (ByRef getC As REALCOLOR, ByVal colorT As Long) As Long
Public Declare Function checkColorTempTest Lib "ColorT.dll" (ByRef getC As REALCOLOR, ByVal colorT As Long) As Long
Public Declare Function adjustColorTemp Lib "ColorT.dll" (ByVal FixValue As Long, ByVal xyAdjMode As Long, ByVal step As Long, ByRef pREALRGB As REALRGB) As Long
