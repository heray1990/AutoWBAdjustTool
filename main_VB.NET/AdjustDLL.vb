Option Strict Off
Option Explicit On
Module AdjustDLL
	'====================================================
	'====================================================
	
	
	Public Declare Function initColorTemp Lib "ColorT.dll" (ByRef Calibration As Integer, ByRef MinBri As Integer, ByVal ModelFile As String, ByVal pCurDir As String) As Integer
	Public Declare Function DeinitColorTemp Lib "ColorT.dll" (ByVal ModelFile As String) As Integer
	'UPGRADE_WARNING: Structure COLORTEMPSPEC may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function setColorTemp Lib "ColorT.dll" (ByVal colorT As String, ByRef pCOLORs As COLORTEMPSPEC, ByVal refHighLowMode As Integer) As Integer
	'UPGRADE_WARNING: Structure REALCOLOR may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function checkColorTemp Lib "ColorT.dll" (ByRef getC As REALCOLOR, ByVal colorT As String) As Integer
	'UPGRADE_WARNING: Structure REALRGB may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function adjustColorTemp Lib "ColorT.dll" (ByVal FixValue As Integer, ByRef pREALRGB As REALRGB, ByRef resultCode As Integer) As Integer
	'UPGRADE_WARNING: Structure REALRGB may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function adjustColorTempOffset Lib "ColorT.dll" (ByRef pREALRGB As REALRGB) As Integer
	'UPGRADE_WARNING: Structure REALRGB may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function adjustColorTempForCIBN Lib "ColorT.dll" (ByRef pREALRGB As REALRGB) As Integer
End Module