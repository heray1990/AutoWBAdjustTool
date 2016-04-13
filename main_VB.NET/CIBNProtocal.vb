Option Strict Off
Option Explicit On
Friend Class CIBNProtocal
	Implements _CommunicationProtocal
	'**********************************************
	' Class module to handle protocal for CIBN and
	' Haier.
	'**********************************************
	
	
	Private mSendDataBuf(11) As Byte
	
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
		mSendDataBuf(10) = &H0
		mSendDataBuf(11) = &H0
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
		'byte10 is the checksum byte.
		'byte10 = 0xFF - byte1 - byte2 - ... - byte9 + 1
		'If it hits 0xFF, ignore 0xFF and plus 1 instead.
		Dim i As Short
		
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
	
	Private Sub CommunicationProtocal_EnterFacMode() Implements _CommunicationProtocal.EnterFacMode
		'Empty cmd for CIBN
	End Sub
	
	Private Sub CommunicationProtocal_ExitFacMode() Implements _CommunicationProtocal.ExitFacMode
		'Empty cmd for CIBN
	End Sub
	
	Private Sub CommunicationProtocal_SwitchInputSource(ByRef strInputSrc As String, ByRef intSrcNum As Short) Implements _CommunicationProtocal.SwitchInputSource
		'55 2E 01 XX 00 00 00 00 00 00 CHK FE
		'HDMI1: 55 2E 01 04 00 00 00 00 00 00 CD FE
		mSendDataBuf(0) = &H55
		mSendDataBuf(1) = &H2E
		mSendDataBuf(2) = &H1
		
		If strInputSrc = "HDMI" Then
			If intSrcNum = 1 Then
				mSendDataBuf(3) = &H4
			ElseIf intSrcNum = 2 Then 
				mSendDataBuf(3) = &H5
			ElseIf intSrcNum = 3 Then 
				mSendDataBuf(3) = &H6
			Else
				mSendDataBuf(3) = &H4
			End If
		ElseIf strInputSrc = "AV" Then 
			mSendDataBuf(3) = &H3
		ElseIf strInputSrc = "YPbPr" Then 
			mSendDataBuf(3) = &H7
		Else
			mSendDataBuf(3) = &H4
		End If
		
		mSendDataBuf(4) = &H0
		mSendDataBuf(5) = &H0
		mSendDataBuf(6) = &H0
		mSendDataBuf(7) = &H0
		mSendDataBuf(8) = &H0
		mSendDataBuf(9) = &H0
		mSendDataBuf(10) = CalChkSum(mSendDataBuf)
		mSendDataBuf(11) = &HFE
		
		SendCmd()
	End Sub
	
	'Set picture mode to standard.
	Private Sub CommunicationProtocal_ResetPicMode() Implements _CommunicationProtocal.ResetPicMode
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
		
		SendCmd()
	End Sub
	
	Private Sub CommunicationProtocal_SetBrightness(ByRef intBrightness As Short) Implements _CommunicationProtocal.SetBrightness
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
		
		SendCmd()
	End Sub
	
	Private Sub CommunicationProtocal_SetContrast(ByRef intContrast As Short) Implements _CommunicationProtocal.SetContrast
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
		
		SendCmd()
	End Sub
	
	Private Sub CommunicationProtocal_SetBacklight(ByRef intBacklight As Short) Implements _CommunicationProtocal.SetBacklight
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
		
		SendCmd()
	End Sub
	
	Private Sub CommunicationProtocal_SelColorTemp(ByRef strColorT As String, ByRef strInputSrc As String, ByRef intSrcNum As Short) Implements _CommunicationProtocal.SelColorTemp
		'55 02 01 XX 00 00 00 00 00 00 CHK FE
		
		mSendDataBuf(0) = &H55
		mSendDataBuf(1) = &H2
		mSendDataBuf(2) = &H1
		
		Select Case strColorT
			Case cstrColorTempCool1
				mSendDataBuf(3) = &H0
			Case cstrColorTempNormal
				mSendDataBuf(3) = &H1
			Case cstrColorTempWarm1
				mSendDataBuf(3) = &H2
		End Select
		
		mSendDataBuf(4) = &H0
		mSendDataBuf(5) = &H0
		mSendDataBuf(6) = &H0
		mSendDataBuf(7) = &H0
		mSendDataBuf(8) = &H0
		mSendDataBuf(9) = &H0
		mSendDataBuf(10) = CalChkSum(mSendDataBuf)
		mSendDataBuf(11) = &HFE
		
		SendCmd()
	End Sub
	
	Private Sub CommunicationProtocal_SetRGBGain(ByRef lngRGain As Integer, ByRef lngGGain As Integer, ByRef lngBGain As Integer) Implements _CommunicationProtocal.SetRGBGain
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
		
		SendCmd()
	End Sub
	
	Private Sub CommunicationProtocal_SetRGBOffset(ByRef lngROffset As Integer, ByRef lngGOffset As Integer, ByRef lngBOffset As Integer) Implements _CommunicationProtocal.SetRGBOffset
		'55 04 06 XX XX XX XX XX XX 00 CHK FE
		mSendDataBuf(0) = &H55
		mSendDataBuf(1) = &H4
		mSendDataBuf(2) = &H6
		mSendDataBuf(3) = CByte((lngROffset) \ 256)
		mSendDataBuf(4) = CByte((lngROffset) Mod 256)
		mSendDataBuf(5) = CByte((lngGOffset) \ 256)
		mSendDataBuf(6) = CByte((lngGOffset) Mod 256)
		mSendDataBuf(7) = CByte((lngBOffset) \ 256)
		mSendDataBuf(8) = CByte((lngBOffset) Mod 256)
		mSendDataBuf(9) = &H0
		mSendDataBuf(10) = CalChkSum(mSendDataBuf)
		mSendDataBuf(11) = &HFE
		
		SendCmd()
	End Sub
	
	Private Sub CommunicationProtocal_SaveWBDataToAllSrc(ByRef strInputSrc As String, ByRef intSrcNum As Short) Implements _CommunicationProtocal.SaveWBDataToAllSrc
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
		mSendDataBuf(10) = &HF8
		mSendDataBuf(11) = &HFE
		
		SendCmd()
	End Sub
End Class