Attribute VB_Name = "ColorTDLL"
Option Explicit
'====================================================
'====================================================


Public Declare Function ColorTInit Lib "ColorT.dll" (ByRef pudtConfigData As udtSpecData) As Long
Public Declare Function ColorTDeInit Lib "ColorT.dll" (ByRef pudtConfigData As udtSpecData) As Long
Public Declare Function ColorTSetSpec Lib "ColorT.dll" (ByVal colorT As String, ByRef pCOLORs As COLORTEMPSPEC, ByVal refHighLowMode As Long) As Long
Public Declare Function ColorTChk Lib "ColorT.dll" (ByRef getC As REALCOLOR, ByVal colorT As String) As Long
Public Declare Function ColorTAdjRGBGainLetv Lib "ColorT.dll" (ByVal FixValue As Long, ByRef pREALRGB As REALRGB, ByRef resultCode As Long) As Long
Public Declare Function ColorTAdjRGBOffset Lib "ColorT.dll" (ByRef pREALRGB As REALRGB) As Long
Public Declare Function ColorTAdjRGBGain Lib "ColorT.dll" (ByRef pREALRGB As REALRGB) As Long

