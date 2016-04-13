Option Strict Off
Option Explicit On
Friend Class LetvProtocal
	Implements _CommunicationProtocal
	'**********************************************
	' Class module to handle protocal for Letv.
	'**********************************************
	
	
	Private mSendDataBuf(9) As Byte
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
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
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Private Sub SendCmd()
		If isUartMode Then
			Form1.MSComm1.Output = VB6.CopyArray(mSendDataBuf)
		Else
			Form1.tcpClient.SendData(mSendDataBuf)
		End If
		
		DelayMS(delayTime)
	End Sub
	
	Private Function CalChkSum(ByRef data() As Byte) As Byte
		Dim i As Short
		
		CalChkSum = &H0
		
		For i = 0 To 8
			CalChkSum = CalChkSum Xor data(i)
		Next i
	End Function
	
	Private Sub CommunicationProtocal_EnterFacMode() Implements _CommunicationProtocal.EnterFacMode
		'6E 51 86 03 FE E1 A0 00 01 04
		mSendDataBuf(0) = &H6E
		mSendDataBuf(1) = &H51
		mSendDataBuf(2) = &H86
		mSendDataBuf(3) = &H3
		mSendDataBuf(4) = &HFE
		mSendDataBuf(5) = &HE1
		mSendDataBuf(6) = &HA0
		mSendDataBuf(7) = &H0
		mSendDataBuf(8) = &H1
		mSendDataBuf(9) = &H4
		
		SendCmd()
	End Sub
	
	Private Sub CommunicationProtocal_ExitFacMode() Implements _CommunicationProtocal.ExitFacMode
		'6E 51 86 03 FE E1 A0 00 00 05
		mSendDataBuf(0) = &H6E
		mSendDataBuf(1) = &H51
		mSendDataBuf(2) = &H86
		mSendDataBuf(3) = &H3
		mSendDataBuf(4) = &HFE
		mSendDataBuf(5) = &HE1
		mSendDataBuf(6) = &HA0
		mSendDataBuf(7) = &H0
		mSendDataBuf(8) = &H0
		mSendDataBuf(9) = &H5
		
		SendCmd()
	End Sub
	
	Private Sub CommunicationProtocal_SwitchInputSource(ByRef strInputSrc As String, ByRef intSrcNum As Short) Implements _CommunicationProtocal.SwitchInputSource
		'HDMI1: 6E 51 86 03 FE 60 00 23 02 05
		mSendDataBuf(0) = &H6E
		mSendDataBuf(1) = &H51
		mSendDataBuf(2) = &H86
		mSendDataBuf(3) = &H3
		mSendDataBuf(4) = &HFE
		mSendDataBuf(5) = &H60
		mSendDataBuf(6) = &H0
		
		If strInputSrc = "HDMI" Then
			If intSrcNum = 1 Then
				mSendDataBuf(7) = &H23
			ElseIf intSrcNum = 2 Then 
				mSendDataBuf(7) = &H43
			ElseIf intSrcNum = 3 Then 
				mSendDataBuf(7) = &H63
			Else
				mSendDataBuf(7) = &H23
			End If
		ElseIf strInputSrc = "AV" Then 
			If intSrcNum = 1 Then
				mSendDataBuf(7) = &H25
			ElseIf intSrcNum = 2 Then 
				mSendDataBuf(7) = &H45
			ElseIf intSrcNum = 3 Then 
				mSendDataBuf(7) = &H65
			Else
				mSendDataBuf(7) = &H25
			End If
		ElseIf strInputSrc = "YPbPr" Then 
			If intSrcNum = 1 Then
				mSendDataBuf(7) = &H27
			ElseIf intSrcNum = 2 Then 
				mSendDataBuf(7) = &H47
			ElseIf intSrcNum = 3 Then 
				mSendDataBuf(7) = &H67
			Else
				mSendDataBuf(7) = &H27
			End If
		Else
			mSendDataBuf(7) = &H23
		End If
		
		mSendDataBuf(8) = &H2
		mSendDataBuf(9) = CalChkSum(mSendDataBuf)
		
		SendCmd()
	End Sub
	
	'Set picture mode to standard.
	Private Sub CommunicationProtocal_ResetPicMode() Implements _CommunicationProtocal.ResetPicMode
		'6E 51 86 03 FE E1 A7 05 01 CHK
		mSendDataBuf(0) = &H6E
		mSendDataBuf(1) = &H51
		mSendDataBuf(2) = &H86
		mSendDataBuf(3) = &H3
		mSendDataBuf(4) = &HFE
		mSendDataBuf(5) = &HE1
		mSendDataBuf(6) = &HA7
		mSendDataBuf(7) = &H5
		mSendDataBuf(8) = &H1
		mSendDataBuf(9) = CalChkSum(mSendDataBuf)
		
		SendCmd()
	End Sub
	
	Private Sub CommunicationProtocal_SetBrightness(ByRef intBrightness As Short) Implements _CommunicationProtocal.SetBrightness
		'6E 51 86 03 FE 10 00 XX XX CHK
		mSendDataBuf(0) = &H6E
		mSendDataBuf(1) = &H51
		mSendDataBuf(2) = &H86
		mSendDataBuf(3) = &H3
		mSendDataBuf(4) = &HFE
		mSendDataBuf(5) = &H10
		mSendDataBuf(6) = &H0
		mSendDataBuf(7) = CByte(intBrightness \ 256)
		mSendDataBuf(8) = CByte(intBrightness Mod 256)
		mSendDataBuf(9) = CalChkSum(mSendDataBuf)
		
		SendCmd()
	End Sub
	
	Private Sub CommunicationProtocal_SetContrast(ByRef intContrast As Short) Implements _CommunicationProtocal.SetContrast
		'6E 51 86 03 FE 12 00 XX XX CHK
		mSendDataBuf(0) = &H6E
		mSendDataBuf(1) = &H51
		mSendDataBuf(2) = &H86
		mSendDataBuf(3) = &H3
		mSendDataBuf(4) = &HFE
		mSendDataBuf(5) = &H12
		mSendDataBuf(6) = &H0
		mSendDataBuf(7) = CByte(intContrast \ 256)
		mSendDataBuf(8) = CByte(intContrast Mod 256)
		mSendDataBuf(9) = CalChkSum(mSendDataBuf)
		
		SendCmd()
	End Sub
	
	Private Sub CommunicationProtocal_SetBacklight(ByRef intBacklight As Short) Implements _CommunicationProtocal.SetBacklight
		'6E 51 86 03 FE 13 00 XX XX CHK
		mSendDataBuf(0) = &H6E
		mSendDataBuf(1) = &H51
		mSendDataBuf(2) = &H86
		mSendDataBuf(3) = &H3
		mSendDataBuf(4) = &HFE
		mSendDataBuf(5) = &H13
		mSendDataBuf(6) = &H0
		mSendDataBuf(7) = CByte(intBacklight \ 256)
		mSendDataBuf(8) = CByte(intBacklight Mod 256)
		mSendDataBuf(9) = CalChkSum(mSendDataBuf)
		
		SendCmd()
	End Sub
	
	Private Sub CommunicationProtocal_SelColorTemp(ByRef strColorT As String, ByRef strInputSrc As String, ByRef intSrcNum As Short) Implements _CommunicationProtocal.SelColorTemp
		Select Case strColorT
			Case cstrColorTempCool1
				Call SetColorTempCool1(strInputSrc, intSrcNum)
			Case cstrColorTempNormal
				Call SetColorTempNormal(strInputSrc, intSrcNum)
			Case cstrColorTempWarm1
				Call SetColorTempWarm1(strInputSrc, intSrcNum)
		End Select
	End Sub
	
	Private Sub SetColorTempCool1(ByRef strInputSrc As String, ByRef intSrcNum As Short)
		'HDMI Cool
		'6E 51 86 03 FE 14 0A 23 01 78
		mSendDataBuf(0) = &H6E
		mSendDataBuf(1) = &H51
		mSendDataBuf(2) = &H86
		mSendDataBuf(3) = &H3
		mSendDataBuf(4) = &HFE
		mSendDataBuf(5) = &H14
		mSendDataBuf(6) = &HA
		
		If strInputSrc = "HDMI" Then
			If intSrcNum = 1 Then
				mSendDataBuf(7) = &H23
			ElseIf intSrcNum = 2 Then 
				mSendDataBuf(7) = &H43
			ElseIf intSrcNum = 3 Then 
				mSendDataBuf(7) = &H63
			Else
				mSendDataBuf(7) = &H23
			End If
		ElseIf strInputSrc = "AV" Then 
			If intSrcNum = 1 Then
				mSendDataBuf(7) = &H25
			ElseIf intSrcNum = 2 Then 
				mSendDataBuf(7) = &H45
			ElseIf intSrcNum = 3 Then 
				mSendDataBuf(7) = &H65
			Else
				mSendDataBuf(7) = &H25
			End If
		ElseIf strInputSrc = "YPbPr" Then 
			If intSrcNum = 1 Then
				mSendDataBuf(7) = &H27
			ElseIf intSrcNum = 2 Then 
				mSendDataBuf(7) = &H47
			ElseIf intSrcNum = 3 Then 
				mSendDataBuf(7) = &H67
			Else
				mSendDataBuf(7) = &H27
			End If
		Else
			mSendDataBuf(7) = &H23
		End If
		
		mSendDataBuf(8) = &H1
		mSendDataBuf(9) = CalChkSum(mSendDataBuf)
		
		SendCmd()
	End Sub
	
	Private Sub SetColorTempNormal(ByRef strInputSrc As String, ByRef intSrcNum As Short)
		'HDMI normal
		'6E 51 86 03 FE 14 06 23 01 74
		mSendDataBuf(0) = &H6E
		mSendDataBuf(1) = &H51
		mSendDataBuf(2) = &H86
		mSendDataBuf(3) = &H3
		mSendDataBuf(4) = &HFE
		mSendDataBuf(5) = &H14
		mSendDataBuf(6) = &H6
		
		If strInputSrc = "HDMI" Then
			If intSrcNum = 1 Then
				mSendDataBuf(7) = &H23
			ElseIf intSrcNum = 2 Then 
				mSendDataBuf(7) = &H43
			ElseIf intSrcNum = 3 Then 
				mSendDataBuf(7) = &H63
			Else
				mSendDataBuf(7) = &H23
			End If
		ElseIf strInputSrc = "AV" Then 
			If intSrcNum = 1 Then
				mSendDataBuf(7) = &H25
			ElseIf intSrcNum = 2 Then 
				mSendDataBuf(7) = &H45
			ElseIf intSrcNum = 3 Then 
				mSendDataBuf(7) = &H65
			Else
				mSendDataBuf(7) = &H25
			End If
		ElseIf strInputSrc = "YPbPr" Then 
			If intSrcNum = 1 Then
				mSendDataBuf(7) = &H27
			ElseIf intSrcNum = 2 Then 
				mSendDataBuf(7) = &H47
			ElseIf intSrcNum = 3 Then 
				mSendDataBuf(7) = &H67
			Else
				mSendDataBuf(7) = &H27
			End If
		Else
			mSendDataBuf(7) = &H23
		End If
		
		mSendDataBuf(8) = &H1
		mSendDataBuf(9) = CalChkSum(mSendDataBuf)
		
		SendCmd()
	End Sub
	
	Private Sub SetColorTempWarm1(ByRef strInputSrc As String, ByRef intSrcNum As Short)
		'HDMI warm
		'6E 51 86 03 FE 14 05 23 01 77
		mSendDataBuf(0) = &H6E
		mSendDataBuf(1) = &H51
		mSendDataBuf(2) = &H86
		mSendDataBuf(3) = &H3
		mSendDataBuf(4) = &HFE
		mSendDataBuf(5) = &H14
		mSendDataBuf(6) = &H5
		
		If strInputSrc = "HDMI" Then
			If intSrcNum = 1 Then
				mSendDataBuf(7) = &H23
			ElseIf intSrcNum = 2 Then 
				mSendDataBuf(7) = &H43
			ElseIf intSrcNum = 3 Then 
				mSendDataBuf(7) = &H63
			Else
				mSendDataBuf(7) = &H23
			End If
		ElseIf strInputSrc = "AV" Then 
			If intSrcNum = 1 Then
				mSendDataBuf(7) = &H25
			ElseIf intSrcNum = 2 Then 
				mSendDataBuf(7) = &H45
			ElseIf intSrcNum = 3 Then 
				mSendDataBuf(7) = &H65
			Else
				mSendDataBuf(7) = &H25
			End If
		ElseIf strInputSrc = "YPbPr" Then 
			If intSrcNum = 1 Then
				mSendDataBuf(7) = &H27
			ElseIf intSrcNum = 2 Then 
				mSendDataBuf(7) = &H47
			ElseIf intSrcNum = 3 Then 
				mSendDataBuf(7) = &H67
			Else
				mSendDataBuf(7) = &H27
			End If
		Else
			mSendDataBuf(7) = &H23
		End If
		
		mSendDataBuf(8) = &H1
		mSendDataBuf(9) = CalChkSum(mSendDataBuf)
		
		SendCmd()
	End Sub
	
	Private Sub CommunicationProtocal_SetRGBGain(ByRef lngRGain As Integer, ByRef lngGGain As Integer, ByRef lngBGain As Integer) Implements _CommunicationProtocal.SetRGBGain
		Call SetRGain(lngRGain)
		Call SetGGain(lngGGain)
		Call SetBGain(lngBGain)
	End Sub
	
	Private Sub SetRGain(ByRef lngRGain As Integer)
		'6E 51 86 03 FE 16 00 XX XX CHK
		mSendDataBuf(0) = &H6E
		mSendDataBuf(1) = &H51
		mSendDataBuf(2) = &H86
		mSendDataBuf(3) = &H3
		mSendDataBuf(4) = &HFE
		mSendDataBuf(5) = &H16
		mSendDataBuf(6) = &H0
		mSendDataBuf(7) = CByte(lngRGain \ 256)
		mSendDataBuf(8) = CByte(lngRGain Mod 256)
		mSendDataBuf(9) = CalChkSum(mSendDataBuf)
		
		SendCmd()
	End Sub
	
	Private Sub SetGGain(ByRef lngGGain As Integer)
		'6E 51 86 03 FE 18 00 XX XX CHK
		mSendDataBuf(0) = &H6E
		mSendDataBuf(1) = &H51
		mSendDataBuf(2) = &H86
		mSendDataBuf(3) = &H3
		mSendDataBuf(4) = &HFE
		mSendDataBuf(5) = &H18
		mSendDataBuf(6) = &H0
		mSendDataBuf(7) = CByte(lngGGain \ 256)
		mSendDataBuf(8) = CByte(lngGGain Mod 256)
		mSendDataBuf(9) = CalChkSum(mSendDataBuf)
		
		SendCmd()
	End Sub
	
	Private Sub SetBGain(ByRef lngBGain As Integer)
		'6E 51 86 03 FE 1A 00 XX XX CHK
		mSendDataBuf(0) = &H6E
		mSendDataBuf(1) = &H51
		mSendDataBuf(2) = &H86
		mSendDataBuf(3) = &H3
		mSendDataBuf(4) = &HFE
		mSendDataBuf(5) = &H1A
		mSendDataBuf(6) = &H0
		mSendDataBuf(7) = CByte(lngBGain \ 256)
		mSendDataBuf(8) = CByte(lngBGain Mod 256)
		mSendDataBuf(9) = CalChkSum(mSendDataBuf)
		
		SendCmd()
	End Sub
	
	Private Sub CommunicationProtocal_SetRGBOffset(ByRef lngROffset As Integer, ByRef lngGOffset As Integer, ByRef lngBOffset As Integer) Implements _CommunicationProtocal.SetRGBOffset
		Call SetROffset(lngROffset)
		Call SetGOffset(lngGOffset)
		Call SetBOffset(lngBOffset)
	End Sub
	
	Private Sub SetROffset(ByRef lngROffset As Integer)
		'6E 51 86 03 FE 6C 00 XX XX CHK
		mSendDataBuf(0) = &H6E
		mSendDataBuf(1) = &H51
		mSendDataBuf(2) = &H86
		mSendDataBuf(3) = &H3
		mSendDataBuf(4) = &HFE
		mSendDataBuf(5) = &H6C
		mSendDataBuf(6) = &H0
		mSendDataBuf(7) = CByte(lngROffset \ 256)
		mSendDataBuf(8) = CByte(lngROffset Mod 256)
		mSendDataBuf(9) = CalChkSum(mSendDataBuf)
		
		SendCmd()
	End Sub
	
	Private Sub SetGOffset(ByRef lngGOffset As Integer)
		'6E 51 86 03 FE 6E 00 XX XX CHK
		mSendDataBuf(0) = &H6E
		mSendDataBuf(1) = &H51
		mSendDataBuf(2) = &H86
		mSendDataBuf(3) = &H3
		mSendDataBuf(4) = &HFE
		mSendDataBuf(5) = &H6E
		mSendDataBuf(6) = &H0
		mSendDataBuf(7) = CByte(lngGOffset \ 256)
		mSendDataBuf(8) = CByte(lngGOffset Mod 256)
		mSendDataBuf(9) = CalChkSum(mSendDataBuf)
		
		SendCmd()
	End Sub
	
	Private Sub SetBOffset(ByRef lngBOffset As Integer)
		'6E 51 86 03 FE 70 00 XX XX CHK
		mSendDataBuf(0) = &H6E
		mSendDataBuf(1) = &H51
		mSendDataBuf(2) = &H86
		mSendDataBuf(3) = &H3
		mSendDataBuf(4) = &HFE
		mSendDataBuf(5) = &H70
		mSendDataBuf(6) = &H0
		mSendDataBuf(7) = CByte(lngBOffset \ 256)
		mSendDataBuf(8) = CByte(lngBOffset Mod 256)
		mSendDataBuf(9) = CalChkSum(mSendDataBuf)
		
		SendCmd()
	End Sub
	
	Private Sub CommunicationProtocal_SaveWBDataToAllSrc(ByRef strInputSrc As String, ByRef intSrcNum As Short) Implements _CommunicationProtocal.SaveWBDataToAllSrc
		'6E 51 86 03 FE 14 05 23 00 76
		mSendDataBuf(0) = &H6E
		mSendDataBuf(1) = &H51
		mSendDataBuf(2) = &H86
		mSendDataBuf(3) = &H3
		mSendDataBuf(4) = &HFE
		mSendDataBuf(5) = &H14
		mSendDataBuf(6) = &H5
		
		If strInputSrc = "HDMI" Then
			If intSrcNum = 1 Then
				mSendDataBuf(7) = &H23
			ElseIf intSrcNum = 2 Then 
				mSendDataBuf(7) = &H43
			ElseIf intSrcNum = 3 Then 
				mSendDataBuf(7) = &H63
			Else
				mSendDataBuf(7) = &H23
			End If
		ElseIf strInputSrc = "AV" Then 
			If intSrcNum = 1 Then
				mSendDataBuf(7) = &H25
			ElseIf intSrcNum = 2 Then 
				mSendDataBuf(7) = &H45
			ElseIf intSrcNum = 3 Then 
				mSendDataBuf(7) = &H65
			Else
				mSendDataBuf(7) = &H25
			End If
		ElseIf strInputSrc = "YPbPr" Then 
			If intSrcNum = 1 Then
				mSendDataBuf(7) = &H27
			ElseIf intSrcNum = 2 Then 
				mSendDataBuf(7) = &H47
			ElseIf intSrcNum = 3 Then 
				mSendDataBuf(7) = &H67
			Else
				mSendDataBuf(7) = &H27
			End If
		Else
			mSendDataBuf(7) = &H23
		End If
		
		mSendDataBuf(8) = &H0
		mSendDataBuf(9) = CalChkSum(mSendDataBuf)
		
		SendCmd()
	End Sub
End Class