Attribute VB_Name = "Module5"
Option Explicit
'====================================================
'====================================================


Public Declare Function initColorTemp Lib "ColorT.dll" (ByRef Calibration As Long, ByRef MinBri As Long, ByVal ModelFile As String, ByVal pCurDir As String) As Long
Public Declare Function DeinitColorTemp Lib "ColorT.dll" (ByVal ModelFile As String) As Long
Public Declare Function setColorTemp Lib "ColorT.dll" (ByVal colorT As Long, ByRef pCOLORs As COLORTEMPSPEC, ByVal refHighLowMode As Long) As Long
Public Declare Function checkColorTemp Lib "ColorT.dll" (ByRef getC As REALCOLOR, ByVal colorT As Long) As Long
Public Declare Function checkColorTempOffset Lib "ColorT.dll" (ByRef getC As REALCOLOR, ByVal colorT As Long) As Long
Public Declare Function adjustColorTemp Lib "ColorT.dll" (ByVal FixValue As Long, ByVal xyAdjMode As Long, ByVal step As Long, ByRef pREALRGB As REALRGB, ByRef resultCode As Long) As Long
Public Declare Function adjustColorTempOffset Lib "ColorT.dll" (ByRef pREALRGB As REALRGB) As Long
Public Declare Function adjustColorTempForCIBN Lib "ColorT.dll" (ByRef pREALRGB As REALRGB) As Long

