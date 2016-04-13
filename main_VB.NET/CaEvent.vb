Option Strict Off
Option Explicit On
Module CaEvent
	Public ObjCa210 As CA200SRVRLib.Ca200
	Public ObjCa As CA200SRVRLib.Ca
	Public ObjProbe As CA200SRVRLib.Probe
	Public ObjMemory As CA200SRVRLib.Memory
	
	
	Public Sub CONNECT_CA210()
		Dim sel As Object
		Dim strERR As String
		Dim iReturn As Short
		
		On Error GoTo ER
		ObjCa210 = New CA200SRVRLib.Ca200
		'UPGRADE_WARNING: Couldn't resolve default property of object sel. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sel = MsgBox("Please Set the Probe at 0-CALL Position! Are you sure?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Calibration")
		
		Select Case sel
			Case MsgBoxResult.Yes
				ObjCa210.AutoConnect()
				ObjCa = ObjCa210.SingleCa
				ObjProbe = ObjCa.SingleProbe
				ObjMemory = ObjCa.Memory
				
				ObjCa.CalZero()
				ObjCa.SyncMode = 3
				ObjCa.AveragingMode = 2
				ObjCa.SetAnalogRange(2.5, 2.5)
				ObjCa.DisplayMode = 0
				ObjMemory.ChannelNO = Ca210ChannelNO
				
				'UPGRADE_WARNING: Form method Form1.ZOrder has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				Form1.BringToFront()
				Delay(200)
				
				MsgBox("Please Set the Probe at Measure Position", MsgBoxStyle.OKOnly + MsgBoxStyle.Information, "Calibration OK")
				IsCa210ok = True
			Case MsgBoxResult.No
		End Select
		
		Exit Sub
		
ER: 
		strERR = "Error from " & Err.Source & Chr(10) & Chr(13)
		strERR = strERR & Err.Description & Chr(10) & Chr(13)
		strERR = strERR & "HRESULT" & CStr(Err.Number - vbObjectError)
		iReturn = MsgBox(strERR, MsgBoxStyle.RetryCancel)
		
		Select Case iReturn
			Case MsgBoxResult.Retry : Resume 
			Case Else
				ObjCa.RemoteMode = 0
		End Select
	End Sub
	
	
	Public Function GetX() As Single
		ObjCa.Measure()
		GetX = ObjProbe.sx
	End Function
	Public Function GetY() As Single
		ObjCa.Measure()
		GetY = ObjProbe.sy
	End Function
	Public Function GetLv() As Single
		ObjCa.Measure()
		GetLv = ObjProbe.lv
	End Function
	Public Function GetGainX() As Single
		ObjCa.Measure()
		GetGainX = ObjProbe.sx
	End Function
	Public Function GetGainY() As Single
		ObjCa.Measure()
		GetGainY = ObjProbe.sy
	End Function
	
	Public Sub subGet_CA210_Value()
		Dim iReturn As Object
		Dim strERR As Object
		Dim CTBuff As Object
		On Error GoTo ER
		
		ObjCa.Measure()
		'UPGRADE_WARNING: Couldn't resolve default property of object CTBuff.x. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CTBuff.x = ObjProbe.sx
		ObjCa.Measure()
		'UPGRADE_WARNING: Couldn't resolve default property of object CTBuff.y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CTBuff.y = ObjProbe.sy
		ObjCa.Measure()
		'UPGRADE_WARNING: Couldn't resolve default property of object CTBuff.lv. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CTBuff.lv = ObjProbe.lv
		
		Exit Sub
ER: 
		'UPGRADE_WARNING: Couldn't resolve default property of object strERR. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strERR = "Error from " & Err.Source & Chr(10) & Chr(13)
		'UPGRADE_WARNING: Couldn't resolve default property of object strERR. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strERR = strERR + Err.Description + Chr(10) + Chr(13)
		'UPGRADE_WARNING: Couldn't resolve default property of object strERR. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strERR = strERR + "HRESULT" + CStr(Err.Number - vbObjectError)
		'UPGRADE_WARNING: Couldn't resolve default property of object iReturn. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		iReturn = MsgBox(strERR, MsgBoxStyle.RetryCancel)
		Select Case iReturn
			Case MsgBoxResult.Retry : Resume 
			Case Else
				ObjCa.RemoteMode = 0
		End Select
	End Sub
End Module