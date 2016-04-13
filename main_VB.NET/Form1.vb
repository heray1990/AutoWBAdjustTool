Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class Form1
	Inherits System.Windows.Forms.Form
	
	Dim RES As Integer
	Dim Result As Boolean
	Dim presetData As COLORTEMPSPEC
	Dim cCOOL1 As COLORTEMPSPEC
	Dim cNORMAL As COLORTEMPSPEC
	Dim cWARM1 As COLORTEMPSPEC
	Dim cFFCOOL1 As COLORTEMPSPEC
	Dim cFFNORMAL As COLORTEMPSPEC
	Dim cFFWARM1 As COLORTEMPSPEC
	Dim rColor As REALCOLOR
	Dim rColorLastChk As REALCOLOR
	Dim Calibrate As Object
	Dim MinBrightness As Integer
	Dim resCodeForAdjustColorTemp As Integer
	Dim cmdMark As String
	Dim clsCommProtocal As _CommunicationProtocal
	Dim clsCIBNProtocal As CIBNProtocal
	Dim clsLetvProtocal As LetvProtocal
	
	Dim ivpg As VPGCtrl.IVPGCtrl
	Private mTitle As String
	Dim m_Title As String
	
	Private WithEvents Obj As VPGCtrl.VPGCtrl
	
	Private Sub subMainProcesser()
		Dim i As Object
		Dim j As Short
		
		On Error GoTo ErrExit
		subInitBeforeRunning()
		
		If IsStop = True Then
			Exit Sub
		End If
		
		If IsCa210ok = False Then
			MsgBox("CA210 disconnected,Please click'Connect'->'Connect CA210'to do operation!", MsgBoxStyle.OKOnly + MsgBoxStyle.Information, "warning")
			subInitAfterRunning()
			
			Exit Sub
		End If
		
		checkResult.BackColor = System.Drawing.ColorTranslator.FromOle(&H80FFFF)
		IsStop = False
		checkResult.Text = "RUN..."
		checkResult.ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0)
		CheckStep.Text = ""
		CheckStep.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
		CheckStep.ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
		
		lbAdjustCOOL_1.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
		lbAdjustCOOL_2.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
		lbAdjustNormal.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
		lbAdjustWARM_1.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
		lbAdjustWARM_2.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
		
		'UPGRADE_ISSUE: PictureBox method Picture1.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		Picture1.Cls()
		lbColorTempWrong.Visible = False
		
		ObjMemory = ObjCa.Memory
		ObjMemory.ChannelNO = Ca210ChannelNO
		
		strBuff = ""
		
		Log_Info("###INITIAL USER###")
		Log_Info("###ADJUST COLORTEMP###")
		
		clsCommProtocal.EnterFacMode()
		
		Call clsCommProtocal.SwitchInputSource(setTVInputSource, setTVInputSourcePortNum)
		
		Call clsCommProtocal.ResetPicMode()
		
		Call ChangePattern("103")
		DelayMS(delayTime)
		
		Label6.Text = "WHITE"
		
ADJUST_GAIN_AGAIN_COOL1: 
		If isAdjustCool1 Then
			lbAdjustCOOL_1.BackColor = System.Drawing.ColorTranslator.FromOle(&H80FFFF)
			Result = autoAdjustColorTemperature_Gain(cstrColorTempCool1, adjustMode3, HighBri)
			
			If Result = False Then
				ShowError_Sys((1))
				GoTo FAIL
			Else
				Call clsCommProtocal.SaveWBDataToAllSrc(setTVInputSource, setTVInputSourcePortNum)
			End If
			
			lbAdjustCOOL_1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0FFC0)
			
			If adjustGainAgainCool1Flag > 0 Then
				GoTo CHECK_COOL1
			End If
		End If
		
ADJUST_GAIN_AGAIN_NORMAL: 
		If isAdjustNormal Then
			lbAdjustNormal.BackColor = System.Drawing.ColorTranslator.FromOle(&H80FFFF)
			Result = autoAdjustColorTemperature_Gain(cstrColorTempNormal, adjustMode3, HighBri)
			
			If Result = False Then
				ShowError_Sys((3))
				GoTo FAIL
			Else
				Call clsCommProtocal.SaveWBDataToAllSrc(setTVInputSource, setTVInputSourcePortNum)
			End If
			
			lbAdjustNormal.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0FFC0)
			
			If adjustGainAgainNormalFlag > 0 Then
				GoTo CHECK_NORMAL
			End If
		End If
		
ADJUST_GAIN_AGAIN_WARM1: 
		If isAdjustWarm1 Then
			lbAdjustWARM_1.BackColor = System.Drawing.ColorTranslator.FromOle(&H80FFFF)
			Result = autoAdjustColorTemperature_Gain(cstrColorTempWarm1, adjustMode3, HighBri)
			
			If Result = False Then
				ShowError_Sys((4))
				GoTo FAIL
			Else
				Call clsCommProtocal.SaveWBDataToAllSrc(setTVInputSource, setTVInputSourcePortNum)
			End If
			
			lbAdjustWARM_1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0FFC0)
			
			If adjustGainAgainWarm1Flag > 0 Then
				GoTo CHECK_WARM1
			End If
		End If
		
		If isAdjustOffset Then
			Label6.Text = "GREY"
			
			Call ChangePattern("109")
			DelayMS(delayTime)
			
			If isAdjustCool1 Then
				lbAdjustCOOL_1.BackColor = System.Drawing.ColorTranslator.FromOle(&H80FFFF)
				Result = autoAdjustColorTemperature_Offset(cstrColorTempCool1, FixG, LowBri)
				
				If Result = False Then
					ShowError_Sys((11))
					GoTo FAIL
				Else
					Call clsCommProtocal.SaveWBDataToAllSrc(setTVInputSource, setTVInputSourcePortNum)
				End If
				
				lbAdjustCOOL_1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0FFC0)
			End If
			
			If isAdjustNormal Then
				lbAdjustNormal.BackColor = System.Drawing.ColorTranslator.FromOle(&H80FFFF)
				Result = autoAdjustColorTemperature_Offset(cstrColorTempNormal, FixG, LowBri)
				
				If Result = False Then
					ShowError_Sys((13))
					GoTo FAIL
				Else
					Call clsCommProtocal.SaveWBDataToAllSrc(setTVInputSource, setTVInputSourcePortNum)
				End If
				
				lbAdjustNormal.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0FFC0)
			End If
			
			If isAdjustWarm1 Then
				lbAdjustWARM_1.BackColor = System.Drawing.ColorTranslator.FromOle(&H80FFFF)
				Result = autoAdjustColorTemperature_Offset(cstrColorTempWarm1, FixG, LowBri)
				
				If Result = False Then
					ShowError_Sys((14))
					GoTo FAIL
				Else
					Call clsCommProtocal.SaveWBDataToAllSrc(setTVInputSource, setTVInputSourcePortNum)
				End If
				
				lbAdjustWARM_1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0FFC0)
			End If
		End If
		
		If isCheckColorTemp Then
			If isAdjustOffset Then
				Call ChangePattern("103")
				DelayMS(delayTime)
			End If
CHECK_COOL1: 
			If isAdjustCool1 Then
				Label6.Text = "CHECK"
				lbAdjustCOOL_1.BackColor = System.Drawing.ColorTranslator.FromOle(&H80FFFF)
				Result = checkColorAgain(cstrColorTempCool1, HighBri)
				
				If Result = False Then
					ShowError_Sys((1))
					
					If adjustGainAgainCool1Flag > 0 Then
						GoTo FAIL
					End If
					
					adjustGainAgainCool1Flag = adjustGainAgainCool1Flag + 1
					
					GoTo ADJUST_GAIN_AGAIN_COOL1
				End If
				
				lbAdjustCOOL_1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0FFC0)
			End If
			
CHECK_NORMAL: 
			If isAdjustNormal Then
				Label6.Text = "CHECK"
				lbAdjustNormal.BackColor = System.Drawing.ColorTranslator.FromOle(&H80FFFF)
				Result = checkColorAgain(cstrColorTempNormal, HighBri)
				
				If Result = False Then
					ShowError_Sys((3))
					
					If adjustGainAgainNormalFlag > 0 Then
						GoTo FAIL
					End If
					
					adjustGainAgainNormalFlag = adjustGainAgainNormalFlag + 1
					
					GoTo ADJUST_GAIN_AGAIN_NORMAL
				End If
				
				lbAdjustNormal.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0FFC0)
			End If
			
CHECK_WARM1: 
			If isAdjustWarm1 Then
				Label6.Text = "CHECK"
				lbAdjustWARM_1.BackColor = System.Drawing.ColorTranslator.FromOle(&H80FFFF)
				Result = checkColorAgain(cstrColorTempWarm1, HighBri)
				
				If Result = False Then
					ShowError_Sys((4))
					
					If adjustGainAgainWarm1Flag > 0 Then
						GoTo FAIL
					End If
					
					adjustGainAgainWarm1Flag = adjustGainAgainWarm1Flag + 1
					
					GoTo ADJUST_GAIN_AGAIN_WARM1
				End If
				
				lbAdjustWARM_1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0FFC0)
			End If
		End If
		
		'Last check:
		'Cool, 100% white pattern, brightness = 100, contrast = 100
		'Check Lv and save x, y, lv
		Call ChangePattern("101")
		DelayMS(delayTime)
		
		Call clsCommProtocal.SetBrightness(100)
		Log_Info("Set brightness to 100")
		
		Call clsCommProtocal.SetContrast(100)
		Log_Info("Set contrast to 100")
		
		Call clsCommProtocal.SetBacklight(100)
		Log_Info("Set backlight to 100")
		
		Call clsCommProtocal.SelColorTemp(cstrColorTempCool1, setTVInputSource, setTVInputSourcePortNum)
		Log_Info("Set color temp to cool1")
		
		DelayMS(delayTime)
		ObjCa.Measure()
		rColorLastChk.xx = CInt(ObjProbe.sx * 10000)
		rColorLastChk.yy = CInt(ObjProbe.sy * 10000)
		rColorLastChk.lv = CInt(ObjProbe.lv)
		
		Log_Info("x = " & Str(rColorLastChk.xx) & ", y = " & Str(rColorLastChk.yy) & ", lv = " & Str(rColorLastChk.lv))
		
		showData((lastChkShwDataStep))
		
		Call clsCommProtocal.SetBrightness(50)
		Log_Info("Set brightness to 50")
		
		Call clsCommProtocal.SetContrast(50)
		Log_Info("Set contrast to 50")
		
		DelayMS(delayTime)
		
		clsCommProtocal.ResetPicMode()
		DelayMS(delayTime)
		
		If rColorLastChk.lv < maxBrightnessSpec Then
			ShowError_Sys((30))
			GoTo FAIL
		End If
		
PASS: 
		clsCommProtocal.ExitFacMode()
		
		cmdMark = "PASS"
		Call saveALLcData()
		
		CheckStep.ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
		CheckStep.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0FFC0)
		CheckStep.Text = CheckStep.Text & "TEST ALL PASS"
		CheckStep.SelectionStart = Len(CheckStep.Text)
		CheckStep.Focus()
		checkResult.ForeColor = System.Drawing.ColorTranslator.FromOle(&HC000)
		checkResult.Text = "PASS"
		DelayMS(delayTime)
		checkResult.BackColor = System.Drawing.ColorTranslator.FromOle(&HFF00)
		checkResult.ForeColor = System.Drawing.ColorTranslator.FromOle(&HC00000)
		
		Label6.Text = "PASS"
		
		Call subInitAfterRunning()
		
		Exit Sub
		
FAIL: 
		clsCommProtocal.ExitFacMode()
		
		cmdMark = "FAIL"
		Call saveALLcData()
		
		CheckStep.SelectionStart = Len(CheckStep.Text)
		CheckStep.Focus()
		CheckStep.BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF)
		checkResult.BackColor = System.Drawing.ColorTranslator.FromOle(&HFF)
		checkResult.ForeColor = System.Drawing.ColorTranslator.FromOle(&H808080)
		checkResult.Text = "FAIL"
		DelayMS(delayTime)
		checkResult.ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
		checkResult.ForeColor = System.Drawing.ColorTranslator.FromOle(&HFFFF)
		
		Label6.Text = "FAIL"
		
		Call subInitAfterRunning()
		
		Exit Sub
		
ErrExit: 
		MsgBox(Err.Description, MsgBoxStyle.Critical, Err.Source)
	End Sub
	
	Private Sub subInitBeforeRunning()
		countTime = 0
		lbTimer.Text = "0s"
		Timer1.Enabled = True
		
		txtInput.ReadOnly = True
		'gstrBarCode = ""
		adjustGainAgainCool1Flag = 0
		adjustGainAgainNormalFlag = 0
		adjustGainAgainWarm1Flag = 0
	End Sub
	
	Private Sub subInitAfterRunning()
		Timer1.Enabled = False
		
		adjustGainAgainCool1Flag = 0
		adjustGainAgainNormalFlag = 0
		adjustGainAgainWarm1Flag = 0
		
		txtInput.Text = ""
		txtInput.Focus()
		txtInput.ReadOnly = False
		
		If isUartMode = False Then
			isNetworkConnected = False
			tcpClient.Close()
		End If
	End Sub
	
	Sub ShowError_Sys(ByRef t As Short)
		Dim s As String
		
		s = "Unknown"
		
		Select Case t
			Case 1
				s = "ColorTemp_COOL_1 is Wrong, Please Check Again."
			Case 2
				s = "ColorTemp_COOL_2 is Wrong, Please Check Again."
			Case 3
				s = "ColorTemp_NORMAL is Wrong, Please Check Again."
			Case 4
				s = "ColorTemp_WARM_1 is Wrong, Please Check Again."
			Case 5
				s = "ColorTemp_WARM_2 is Wrong, Please Check Again."
			Case 6
				s = "LAB_SN:" & gstrBarCode & "(End)  Len:" & Str(gintBarCodeLen) & vbCrLf & "条形码长度不对，请确认！"
			Case 7
				s = "Can not Write DVI EDID."
			Case 8
				s = "Calibrate FAIL.(AUTO LEVEL)"
			Case 9
				s = "RS232 Connector Error"
			Case 10
				s = "Read DSUB EDID FAIL"
			Case 11
				s = "OFFSET_Color_COOL_1 is Wrong, Please Check Again."
			Case 12
				s = "OFFSET_Color_COOL_2 is Wrong, Please Check Again."
			Case 13
				s = "OFFSET_Color_NORMAL is Wrong, Please Check Again."
			Case 14
				s = "OFFSET_Color_WARM_1 is Wrong, Please Check Again."
			Case 15
				s = "OFFSET_Color_WARM_2 is Wrong, Please Check Again."
			Case 16
				s = "HDMI2 CheckSum is Wrong"
			Case 17
				s = "Can not Write HDMI-2 EDID."
			Case 18
				s = "min_Brightness is over SPEC."
			Case 19
				s = "FW Version is Wrong."
			Case 20
				s = "Can not Write OSD-SN."
			Case 21
				s = "max_Brightness is over SPEC."
			Case 22
				s = "ColorTemp_COOL_1 is Wrong, Please Check Again."
			Case 23
				s = "ColorTemp_COOL_2 is Wrong, Please Check Again."
			Case 24
				s = "ColorTemp_NORMAL is Wrong, Please Check Again."
			Case 25
				s = "ColorTemp_WARM_1 is Wrong, Please Check Again."
			Case 26
				s = "ColorTemp_WARM_2 is Wrong, Please Check Again."
			Case 27
				s = "ColorTemp_5000 is Wrong, Please Check Again."
			Case 28
				s = "ColorTemp_3000 is Wrong, Please Check Again."
			Case 29
				s = "LightSensor Data is Wrong, Please Check Again."
			Case 30
				s = "亮度不在规格！"
			Case 31
				s = ""
		End Select
		
		CheckStep.ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
		CheckStep.Text = CheckStep.Text & "Error Code:" & Str(t) & vbCrLf & s & vbCrLf
		CheckStep.SelectionStart = Len(CheckStep.Text)
		CheckStep.Focus()
	End Sub
	
	Private Function autoAdjustColorTemperature_Gain(ByRef strColorTemp As String, ByRef adjustVal As Integer, ByRef HighLowMode As Integer) As Boolean
		Dim i, j As Object
		Dim k As Short
		
		Call clsCommProtocal.SelColorTemp(strColorTemp, setTVInputSource, setTVInputSourcePortNum)
		DelayMS(delayTime)
		
		' Set Offset first
		If adjustGainAgainCool1Flag = 0 Then
			Call setColorTemp(strColorTemp, presetData, 0)
			DelayMS(delayTime)
			
			rRGB.cRR = presetData.nColorRR
			rRGB.cGG = presetData.nColorGG
			rRGB.cBB = presetData.nColorBB
			
			Call saveData(strColorTemp, 0)
		End If
		
		Call LoadData(strColorTemp, 0)
		Call clsCommProtocal.SetRGBOffset(rRGB1.cRR, rRGB1.cGG, rRGB1.cBB)
		
		Log_Info("========Adjust " & strColorTemp & "========")
		
		For j = 1 To 2
			Call setColorTemp(strColorTemp, presetData, HighLowMode)
			DelayMS(delayTime)
			
			Log_Info("Init current colorTemp. RES:" & Str(RES))
			rRGB.cRR = presetData.nColorRR
			rRGB.cGG = presetData.nColorGG
			rRGB.cBB = presetData.nColorBB
			
			Label1.Text = Str(presetData.xx)
			Label3.Text = Str(presetData.yy)
			
			Call clsCommProtocal.SetRGBGain(rRGB.cRR, rRGB.cGG, rRGB.cBB)
			
			showData((1))
			
			resCodeForAdjustColorTemp = 0
			
			For k = 1 To 50
				If IsStop = True Then GoTo Cancel
				
				RES = checkColorTemp(rColor, strColorTemp)
				Log_Info("Check colorTemp. RES:" & Str(RES))
				
				If RES Then Exit For
				
				If RES = False Then
					If UCase(gstrBrand) = "CIBN" Or UCase(gstrBrand) = "CAN" Or UCase(gstrBrand) = "CANTV" Or UCase(gstrBrand) = "HAIER" Then
						Call adjustColorTempForCIBN(rRGB)
					Else ' Letv
						If resCodeForAdjustColorTemp = 0 Then
							Call adjustColorTemp(adjustMode3, rRGB, resCodeForAdjustColorTemp)
						ElseIf resCodeForAdjustColorTemp = 1 Then 
							Call adjustColorTemp(adjustMode1, rRGB, resCodeForAdjustColorTemp)
						ElseIf resCodeForAdjustColorTemp = 2 Then 
							Call adjustColorTemp(adjustMode2, rRGB, resCodeForAdjustColorTemp)
						ElseIf resCodeForAdjustColorTemp = 3 Then 
							Call adjustColorTemp(adjustMode3, rRGB, resCodeForAdjustColorTemp)
						ElseIf resCodeForAdjustColorTemp = 4 Then 
							Call adjustColorTemp(adjustMode4, rRGB, resCodeForAdjustColorTemp)
						End If
					End If
					
					Log_Info("SET_RGB_GAN: R = " & CStr(rRGB.cRR) & ", G = " & CStr(rRGB.cGG) & ", B = " & CStr(rRGB.cBB) & ", resultcode = " & CStr(resCodeForAdjustColorTemp))
					
					Call clsCommProtocal.SetRGBGain(rRGB.cRR, rRGB.cGG, rRGB.cBB)
					
					showData((2))
				End If
			Next k
			
			If RES Then Exit For
			
		Next j
		
Cancel: 
		If RES Then
			Call saveData(strColorTemp, HighLowMode)
			Log_Info("Save current data of " & strColorTemp & ".")
			autoAdjustColorTemperature_Gain = True
		Else
			autoAdjustColorTemperature_Gain = False
		End If
		
	End Function
	
	Private Function autoAdjustColorTemperature_Offset(ByRef strColorTemp As String, ByRef FixValue As Integer, ByRef HighLowMode As Integer) As Boolean
		Dim i, j As Object
		Dim k As Short
		
		Call clsCommProtocal.SelColorTemp(strColorTemp, setTVInputSource, setTVInputSourcePortNum)
		DelayMS(delayTime)
		
		Log_Info("========Adjust " & strColorTemp & "========")
		
		For j = 1 To 2
			Call setColorTemp(strColorTemp, presetData, HighLowMode)
			DelayMS(delayTime)
			Log_Info("Init current colorTemp. RES:" & Str(RES))
			rRGB.cRR = presetData.nColorRR
			rRGB.cGG = presetData.nColorGG
			rRGB.cBB = presetData.nColorBB
			
			'Label1 = Str$(presetData.xx)
			'Label3 = Str$(presetData.yy)
			
			Call clsCommProtocal.SetRGBOffset(rRGB.cRR, rRGB.cGG, rRGB.cBB)
			
			showData((3))
			
			For k = 1 To 50
				If IsStop = True Then GoTo Cancel
				
				RES = checkColorTemp(rColor, strColorTemp)
				Log_Info("Check colorTemp. RES:" & Str(RES))
				
				If RES Then Exit For
				If RES = False Then
					Call adjustColorTempOffset(rRGB)
					
					Log_Info("SET_RGB_OFFSET: R = " & CStr(rRGB.cRR) & ", G = " & CStr(rRGB.cGG) & ", B = " & CStr(rRGB.cBB))
					
					Call clsCommProtocal.SetRGBOffset(rRGB.cRR, rRGB.cGG, rRGB.cBB)
					
					showData((4))
				End If
				
				DelayMS(200)
			Next k
			
			If RES Then Exit For
			
			DelayMS(delayTime)
		Next j
		
Cancel: 
		If RES Then
			Call saveData(strColorTemp, HighLowMode)
			Log_Info("Save current data of " & strColorTemp & ".")
			autoAdjustColorTemperature_Offset = True
		Else
			autoAdjustColorTemperature_Offset = False
		End If
		
	End Function
	
	Private Function checkColorAgain(ByRef strColorTemp As String, ByRef HighLowMode As Integer) As Boolean
		Dim i, j As Object
		Dim k As Short
		
		Call clsCommProtocal.SelColorTemp(strColorTemp, setTVInputSource, setTVInputSourcePortNum)
		DelayMS(delayTime)
		
		Log_Info("========Check " & strColorTemp & "========")
		
		For j = 1 To 2
			Call setColorTemp(strColorTemp, presetData, HighLowMode)
			DelayMS(delayTime)
			Log_Info("Init current colorTemp. RES:" & Str(RES))
			
			Label1.Text = Str(presetData.xx)
			Label3.Text = Str(presetData.yy)
			
			showData((5))
			
			If IsStop = True Then GoTo Cancel
			
			RES = checkColorTemp(rColor, strColorTemp)
			Log_Info("Check colorTemp. RES:" & Str(RES))
			
			If RES Then Exit For
			
			DelayMS(delayTime)
		Next j
		
Cancel: 
		If RES Then
			checkColorAgain = True
		Else
			checkColorAgain = False
		End If
		
	End Function
	
	
	'step = lastChkShwDataStep: Check max brightness of TV with brightness 100 and contrast 100 in 100% white pattern.
	'UPGRADE_NOTE: step was upgraded to step_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub showData(ByRef step_Renamed As Short)
		On Error Resume Next
		Dim xPos, yPos As Object
		Dim vPos As Integer
		
		DelayMS(delayTime)
		ObjCa.Measure()
		rColor.xx = CInt(ObjProbe.sx * 10000)
		rColor.yy = CInt(ObjProbe.sy * 10000)
		rColor.lv = CInt(ObjProbe.lv)
		
		'UPGRADE_ISSUE: PictureBox method Picture1.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		Picture1.Cls()
		
		'The values here are about 15 times bigger than the actual pixel.
		'(1515,1275) is the origin of dx-dy axis.
		'In lv axis, 1660 is the distance from the bottom edge of blue rectangle to the top of Picture1.
		'In dx, 365 is half a side of blue rectangle.
		'UPGRADE_WARNING: Couldn't resolve default property of object xPos. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xPos = 1515 + (rColor.xx - presetData.xx) * 365 / presetData.xt
		'UPGRADE_WARNING: Couldn't resolve default property of object yPos. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		yPos = 1275 - (rColor.yy - presetData.yy) * 385 / presetData.yt
		If step_Renamed = lastChkShwDataStep Then
			vPos = 1660 - (rColor.lv - maxBrightnessSpec) * 385 / 50
		Else
			vPos = 1660 - (rColor.lv - presetData.lv) * 385 / 50
		End If
		
		'In dx-dy axis, 360 is the distance from left edge of white rectangle to the left of Picture1.
		'In dx-dy axis, 2660 is the distance from right edge of white rectangle to the left of Picture1.
		'In dx-dy axis, 80 is the distance from top edge of white rectangle to the top of Picture1.
		'In dx-dy axis, 2660 is the distance from bottom edge of white rectangle to the top of Picture1.
		'UPGRADE_WARNING: Couldn't resolve default property of object xPos. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If xPos < 360 Then xPos = 360
		'UPGRADE_WARNING: Couldn't resolve default property of object xPos. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If xPos > 2660 Then xPos = 2660
		'UPGRADE_WARNING: Couldn't resolve default property of object yPos. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If yPos < 80 Then yPos = 80
		'UPGRADE_WARNING: Couldn't resolve default property of object yPos. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If yPos > 2480 Then yPos = 2480
		
		If step_Renamed <> lastChkShwDataStep Then
			If System.Math.Abs(rColor.xx - presetData.xx) <= presetData.xt And System.Math.Abs(rColor.yy - presetData.yy) <= presetData.yt Then
				lbColorTempWrong.Visible = False
				'UPGRADE_ISSUE: PictureBox method Picture1.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				Picture1.Circle (xPos, yPos), 23, &H30FF30
			Else
				lbColorTempWrong.Visible = True
				'UPGRADE_ISSUE: PictureBox method Picture1.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				Picture1.Circle (xPos, yPos), 23, &HFF
				
				If rColor.xx < 5 Then
					IsStop = True
					ObjCa.RemoteMode = 2
					MsgBox("Please check the CA210 Probe is OK or not.")
					RES = False
				End If
			End If
		End If
		
		'In lv axis, 3060 is the distance from left edge of white rectangle to the left of Picture1.
		'In lv axis, 3390 is the distance from right edge of white rectangle to the left of Picture1.
		If step_Renamed = lastChkShwDataStep Then
			If rColor.lv > maxBrightnessSpec Then
				'UPGRADE_ISSUE: PictureBox method Picture1.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				Picture1.Line (3060, vPos) - (3390, vPos), &H30FF30
			Else
				'UPGRADE_ISSUE: PictureBox method Picture1.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				Picture1.Line (3060, vPos) - (3390, vPos), &HFF
			End If
		Else
			If rColor.lv > presetData.lv Then
				'UPGRADE_ISSUE: PictureBox method Picture1.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				Picture1.Line (3060, vPos) - (3390, vPos), &H30FF30
			Else
				'UPGRADE_ISSUE: PictureBox method Picture1.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				Picture1.Line (3060, vPos) - (3390, vPos), &HFF
			End If
		End If
		
		Log_Info("_x/y/Lv: " & Str(rColor.xx) & " / " & Str(rColor.yy) & " / " & Str(rColor.lv))
		
		If Label6.Text <> "CHECK" Then Log_Info("_R/G/B:  " & Str(rRGB.cRR) & " / " & Str(rRGB.cGG) & " / " & Str(rRGB.cBB))
		
		Label_x.Text = Str(rColor.xx)
		Label_y.Text = Str(rColor.yy)
		Label_Lv.Text = Str(rColor.lv)
		DelayMS(30)
	End Sub
	
	Public Sub tbDisConnectastro_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tbDisConnectastro.Click
		ObjCa.RemoteMode = 0
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		countTime = countTime + 1
		lbTimer.Text = CStr(countTime) & "s"
	End Sub
	
	Public Sub vbSetSPEC_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles vbSetSPEC.Click
		frmSetData.Show()
	End Sub
	
	Public Sub vbAbout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles vbAbout.Click
		frmAbout.Show()
	End Sub
	
	Public Sub vbConCA310_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles vbConCA310.Click
		If IsCa210ok = True Then
			ObjCa.RemoteMode = 1
			Exit Sub
		Else
			CONNECT_CA210()
		End If
	End Sub
	
	
	Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		i = 0
		IsStop = False
		txtInput.ReadOnly = False
		
		gstrBrand = Split(gstrCurProjName, gstrDelimiterForProjName)(0)
		
		If UCase(gstrBrand) = "CIBN" Or UCase(gstrBrand) = "CAN" Or UCase(gstrBrand) = "CANTV" Or UCase(gstrBrand) = "HAIER" Then
			clsCIBNProtocal = New CIBNProtocal
			clsCommProtocal = clsCIBNProtocal
		Else
			clsLetvProtocal = New LetvProtocal
			clsCommProtocal = clsLetvProtocal
		End If
		
		mTitle = Me.Text
		subInitInterface()
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Calibrate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		RES = initColorTemp(Calibrate, MinBrightness, gstrCurProjName, My.Application.Info.DirectoryPath)
	End Sub
	
	Public Sub subInitInterface()
		Dim clsConfigData As ProjectConfig
		
		clsConfigData = New ProjectConfig
		clsConfigData.LoadConfigData()
		
		setTVCurrentComBaud = CInt(clsConfigData.ComBaud)
		setTVCurrentComID = clsConfigData.ComID
		setTVInputSource = clsConfigData.inputSource
		setTVInputSourcePortNum = CShort(VB.Right(setTVInputSource, 1))
		setTVInputSource = VB.Left(setTVInputSource, Len(setTVInputSource) - 1)
		delayTime = clsConfigData.DelayMS
		Ca210ChannelNO = clsConfigData.ChannelNum
		gintBarCodeLen = clsConfigData.BarCodeLen
		maxBrightnessSpec = clsConfigData.LvSpec
		gstrVPGModel = clsConfigData.VPGModel
		isAdjustCool2 = clsConfigData.EnableCool2
		isAdjustCool1 = clsConfigData.EnableCool1
		isAdjustNormal = clsConfigData.EnableNormal
		isAdjustWarm1 = clsConfigData.EnableWarm1
		isAdjustWarm2 = clsConfigData.EnableWarm2
		isCheckColorTemp = clsConfigData.EnableChkColor
		isAdjustOffset = clsConfigData.EnableAdjOffset
		
		If clsConfigData.CommMode = DeclareVariables.CommunicationMode.modeUART Then
			isUartMode = True
			lbCommMode.Text = "UART"
			subInitComPort()
		Else
			isUartMode = False
			lbCommMode.Text = "Network"
			subInitNetwork()
		End If
		
		'UPGRADE_NOTE: Object clsConfigData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		clsConfigData = Nothing
		
		txtInput.Text = ""
		lbModelName.Text = gstrCurProjName
		
		If isAdjustCool1 = True Then lbAdjustCOOL_1.ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
		If isAdjustCool2 = True Then lbAdjustCOOL_2.ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
		If isAdjustNormal = True Then lbAdjustNormal.ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
		If isAdjustWarm1 = True Then lbAdjustWARM_1.ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
		If isAdjustWarm2 = True Then lbAdjustWARM_2.ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
		
		If isAdjustCool1 = False Then lbAdjustCOOL_1.ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
		If isAdjustCool2 = False Then lbAdjustCOOL_2.ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
		If isAdjustNormal = False Then lbAdjustNormal.ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
		If isAdjustWarm1 = False Then lbAdjustWARM_1.ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
		If isAdjustWarm2 = False Then lbAdjustWARM_2.ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
		
		InitVPGDevice()
		DelayMS(delayTime)
		
		If setTVInputSource = "HDMI" Then
			'Timing 69: HDMI-720P60
			'Timing 74: HDMI-1080P60
			Call ChangeTiming("69")
		ElseIf setTVInputSource = "AV" Then 
			'Timing 38: PAL-BDGHI
			Call ChangeTiming("38")
		End If
	End Sub
	
	Private Sub subInitComPort()
		If MSComm1.PortOpen = True Then
			MSComm1.PortOpen = False
		End If
		
		MSComm1.CommPort = setTVCurrentComID
		MSComm1.Settings = setTVCurrentComBaud & ",N,8,1"
		MSComm1.InputLen = 0
		
		MSComm1.InBufferCount = 0
		MSComm1.OutBufferCount = 0
		MSComm1.InputMode = MSCommLib.InputModeConstants.comInputModeBinary
		
		MSComm1.NullDiscard = False
		MSComm1.DTREnable = False
		MSComm1.EOFEnable = False
		MSComm1.RTSEnable = False
		MSComm1.SThreshold = 1
		MSComm1.RThreshold = 1
		MSComm1.InBufferSize = 1024
		MSComm1.OutBufferSize = 512
	End Sub
	
	Private Sub subInitNetwork()
		isNetworkConnected = False
		With tcpClient
			.Protocol = MSWinsockLib.ProtocolConstants.sckTCPProtocol
			' IMPORTANT: be sure to change the RemoteHost
			' value to the name of your computer.
			.RemoteHost = strRemoteHost
			.RemotePort = lngRemotePort
		End With
	End Sub
	
	Private Sub txtInput_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtInput.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		On Error GoTo ErrExit
		
		If KeyAscii = 13 Then
			IsStop = False
			
			If txtInput.ReadOnly = False Then
				If txtInput.Text = "" Or Len(txtInput.Text) <> gintBarCodeLen Then
					MsgBox("条形码不对，请确认！(要求长度为：" & CStr(gintBarCodeLen) & ")")
					txtInput.Text = ""
					GoTo EventExitSub
				Else
					gstrBarCode = txtInput.Text
				End If
				
				If isUartMode = True Then
					If MSComm1.PortOpen = False Then
						MSComm1.PortOpen = True
					End If
					subMainProcesser()
				Else
					isNetworkConnected = False
					Do 
						'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
						If tcpClient.CtlState = MSWinsockLib.StateConstants.sckClosed Then
							Log_Info("TCP Connect")
							tcpClient.Connect()
							txtInput.ReadOnly = True
						End If
						Call DelaySWithFlag(cmdReceiveWaitS * 2, isNetworkConnected)
						
						'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
						If tcpClient.CtlState = MSWinsockLib.StateConstants.sckConnected Then
							subMainProcesser()
							Exit Do
						Else
							'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
							If tcpClient.CtlState <> MSWinsockLib.StateConstants.sckClosed Then
								tcpClient.Close()
							End If
							i = i + 1
						End If
						Log_Info("Re-connect to TV.")
					Loop While i <= 5
					txtInput.ReadOnly = False
				End If
			End If
			
			If IsStop = True Then
				GoTo EventExitSub
			End If
		End If
		GoTo EventExitSub
		
ErrExit: 
		txtInput.Text = ""
		'Invalid Port Number
		If Err.Number = 8002 Then
			MsgBox(Err.Description, MsgBoxStyle.Critical, Err.Source)
		End If
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub Form1_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		On Error GoTo ErrExit
		
		If UCase(gstrBrand) = "CIBN" Or UCase(gstrBrand) = "CAN" Or UCase(gstrBrand) = "CANTV" Or UCase(gstrBrand) = "HAIER" Then
			If Not (clsCIBNProtocal Is Nothing) Then
				'UPGRADE_NOTE: Object clsCIBNProtocal may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				clsCIBNProtocal = Nothing
			End If
		Else
			If Not (clsLetvProtocal Is Nothing) Then
				'UPGRADE_NOTE: Object clsLetvProtocal may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				clsLetvProtocal = Nothing
			End If
		End If
		
		If Not (clsCommProtocal Is Nothing) Then
			'UPGRADE_NOTE: Object clsCommProtocal may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			clsCommProtocal = Nothing
		End If
		
		
		IsStop = True
		If (IsCa210ok = True) Then
			ObjCa.RemoteMode = 0
		End If
		
		If MSComm1.PortOpen = True Then
			MSComm1.PortOpen = False
		End If
		
		Call DeinitColorTemp(gstrCurProjName)
		End
		Exit Sub
		
ErrExit: 
		MsgBox(Err.Description, MsgBoxStyle.Critical, Err.Source)
	End Sub
	
	
	Private Sub saveData(ByRef strColorTemp As String, ByRef HL As Integer)
		
		Select Case strColorTemp
			Case cstrColorTempCool1
				If HL Then
					cCOOL1.xx = rColor.xx
					cCOOL1.yy = rColor.yy
					cCOOL1.lv = rColor.lv
					cCOOL1.nColorRR = rRGB.cRR
					cCOOL1.nColorGG = rRGB.cGG
					cCOOL1.nColorBB = rRGB.cBB
				Else
					cFFCOOL1.xx = rColor.xx
					cFFCOOL1.yy = rColor.yy
					cFFCOOL1.lv = rColor.lv
					cFFCOOL1.nColorRR = rRGB.cRR
					cFFCOOL1.nColorGG = rRGB.cGG
					cFFCOOL1.nColorBB = rRGB.cBB
				End If
				
			Case cstrColorTempNormal
				If HL Then
					cNORMAL.xx = rColor.xx
					cNORMAL.yy = rColor.yy
					cNORMAL.lv = rColor.lv
					cNORMAL.nColorRR = rRGB.cRR
					cNORMAL.nColorGG = rRGB.cGG
					cNORMAL.nColorBB = rRGB.cBB
				Else
					cFFNORMAL.xx = rColor.xx
					cFFNORMAL.yy = rColor.yy
					cFFNORMAL.lv = rColor.lv
					cFFNORMAL.nColorRR = rRGB.cRR
					cFFNORMAL.nColorGG = rRGB.cGG
					cFFNORMAL.nColorBB = rRGB.cBB
				End If
				
			Case cstrColorTempWarm1
				If HL Then
					cWARM1.xx = rColor.xx
					cWARM1.yy = rColor.yy
					cWARM1.lv = rColor.lv
					cWARM1.nColorRR = rRGB.cRR
					cWARM1.nColorGG = rRGB.cGG
					cWARM1.nColorBB = rRGB.cBB
				Else
					cFFWARM1.xx = rColor.xx
					cFFWARM1.yy = rColor.yy
					cFFWARM1.lv = rColor.lv
					cFFWARM1.nColorRR = rRGB.cRR
					cFFWARM1.nColorGG = rRGB.cGG
					cFFWARM1.nColorBB = rRGB.cBB
				End If
		End Select
		
	End Sub
	
	Private Sub LoadData(ByRef strColorTemp As String, ByRef isGain As Boolean)
		Select Case strColorTemp
			Case cstrColorTempCool1
				If isGain Then
					rRGB1.cRR = cCOOL1.nColorRR
					rRGB1.cGG = cCOOL1.nColorGG
					rRGB1.cBB = cCOOL1.nColorBB
				Else
					rRGB1.cRR = cFFCOOL1.nColorRR
					rRGB1.cGG = cFFCOOL1.nColorGG
					rRGB1.cBB = cFFCOOL1.nColorBB
				End If
				
			Case cstrColorTempNormal
				If isGain Then
					rRGB1.cRR = cNORMAL.nColorRR
					rRGB1.cGG = cNORMAL.nColorGG
					rRGB1.cBB = cNORMAL.nColorBB
				Else
					rRGB1.cRR = cFFNORMAL.nColorRR
					rRGB1.cGG = cFFNORMAL.nColorGG
					rRGB1.cBB = cFFNORMAL.nColorBB
				End If
				
			Case cstrColorTempWarm1
				If isGain Then
					rRGB1.cRR = cWARM1.nColorRR
					rRGB1.cGG = cWARM1.nColorGG
					rRGB1.cBB = cWARM1.nColorBB
				Else
					rRGB1.cRR = cFFWARM1.nColorRR
					rRGB1.cGG = cFFWARM1.nColorGG
					rRGB1.cBB = cFFWARM1.nColorBB
				End If
		End Select
	End Sub
	
	Private Sub saveALLcData()
		If gstrBarCode = "" Then
			Exit Sub
		Else
			sqlstring = "select * from DataRecord"
			Executesql(sqlstring)
			rs.AddNew()
			
			rs.Fields(0).Value = gstrCurProjName
			rs.Fields(1).Value = gstrBarCode
			
			rs.Fields(2).Value = cCOOL1.xx
			rs.Fields(3).Value = cCOOL1.yy
			rs.Fields(4).Value = cCOOL1.nColorRR
			rs.Fields(5).Value = cCOOL1.nColorGG
			rs.Fields(6).Value = cCOOL1.nColorBB
			rs.Fields(7).Value = cNORMAL.xx
			rs.Fields(8).Value = cNORMAL.yy
			rs.Fields(9).Value = cNORMAL.nColorRR
			rs.Fields(10).Value = cNORMAL.nColorGG
			rs.Fields(11).Value = cNORMAL.nColorBB
			rs.Fields(12).Value = cWARM1.xx
			rs.Fields(13).Value = cWARM1.yy
			rs.Fields(14).Value = cWARM1.nColorRR
			rs.Fields(15).Value = cWARM1.nColorGG
			rs.Fields(16).Value = cWARM1.nColorBB
			
			rs.Fields(17).Value = cFFCOOL1.nColorRR
			rs.Fields(18).Value = cFFCOOL1.nColorGG
			rs.Fields(19).Value = cFFCOOL1.nColorBB
			rs.Fields(20).Value = cFFNORMAL.nColorRR
			rs.Fields(21).Value = cFFNORMAL.nColorGG
			rs.Fields(22).Value = cFFNORMAL.nColorBB
			rs.Fields(23).Value = cFFWARM1.nColorRR
			rs.Fields(24).Value = cFFWARM1.nColorGG
			rs.Fields(25).Value = cFFWARM1.nColorBB
			
			rs.Fields(26).Value = rColorLastChk.lv
			rs.Fields(27).Value = maxBrightnessSpec
			
			rs.Fields(28).Value = cmdMark
			rs.Fields(29).Value = Today
			rs.Fields(30).Value = TimeOfDay
			
			rs.Update()
			
			'UPGRADE_NOTE: Object cn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			cn = Nothing
			'UPGRADE_NOTE: Object rs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rs = Nothing
			sqlstring = ""
		End If
	End Sub
	
	Private Sub tcpClient_ConnectEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tcpClient.ConnectEvent
		'Success to connect the TV.
		isNetworkConnected = True
	End Sub
	
	Private Sub InitVPGDevice()
		
		Select Case gstrVPGModel
			Case "2401"
				ivpg = New VPGCtrl.VPGCtrl_24xx
				Obj = ivpg
				ivpg.InitDevice(VPGCtrl.VPG_MODEL.VPG_MODEL_VPG2401)
			Case "2402"
				ivpg = New VPGCtrl.VPGCtrl_24xx
				Obj = ivpg
				ivpg.InitDevice(VPGCtrl.VPG_MODEL.VPG_MODEL_VPG2402)
			Case "2333_B"
				ivpg = New VPGCtrl.VPGCtrl_24xx
				Obj = ivpg
				ivpg.InitDevice(VPGCtrl.VPG_MODEL.VPG_MODEL_VPG2333_B)
			Case "23293_B"
				ivpg = New VPGCtrl.VPGCtrl_24xx
				Obj = ivpg
				ivpg.InitDevice(VPGCtrl.VPG_MODEL.VPG_MODEL_VPG23293_B)
			Case "23294"
				ivpg = New VPGCtrl.VPGCtrl_24xx
				Obj = ivpg
				ivpg.InitDevice(VPGCtrl.VPG_MODEL.VPG_MODEL_VPG23294)
			Case "22293"
				ivpg = New VPGCtrl.VPGCtrl_22xx
				Obj = ivpg
				ivpg.InitDevice(VPGCtrl.VPG_MODEL.VPG_MODEL_VPG22293)
			Case "22293_A"
				ivpg = New VPGCtrl.VPGCtrl_22xx
				Obj = ivpg
				ivpg.InitDevice(VPGCtrl.VPG_MODEL.VPG_MODEL_VPG22293_A)
			Case "22293_B"
				ivpg = New VPGCtrl.VPGCtrl_22xx
				Obj = ivpg
				ivpg.InitDevice(VPGCtrl.VPG_MODEL.VPG_MODEL_VPG22293_B)
			Case "2233"
				ivpg = New VPGCtrl.VPGCtrl_22xx
				Obj = ivpg
				ivpg.InitDevice(VPGCtrl.VPG_MODEL.VPG_MODEL_VPG2233)
			Case "2233_A"
				ivpg = New VPGCtrl.VPGCtrl_22xx
				Obj = ivpg
				ivpg.InitDevice(VPGCtrl.VPG_MODEL.VPG_MODEL_VPG2233_A)
			Case "2233_B"
				ivpg = New VPGCtrl.VPGCtrl_22xx
				Obj = ivpg
				ivpg.InitDevice(VPGCtrl.VPG_MODEL.VPG_MODEL_VPG2233_B)
			Case "2234"
				ivpg = New VPGCtrl.VPGCtrl_22xx
				Obj = ivpg
				ivpg.InitDevice(VPGCtrl.VPG_MODEL.VPG_MODEL_VPG2234)
			Case "22294"
				ivpg = New VPGCtrl.VPGCtrl_22xx
				Obj = ivpg
				ivpg.InitDevice(VPGCtrl.VPG_MODEL.VPG_MODEL_VPG22294)
			Case "22294_A"
				ivpg = New VPGCtrl.VPGCtrl_22xx
				Obj = ivpg
				ivpg.InitDevice(VPGCtrl.VPG_MODEL.VPG_MODEL_VPG22294_A)
		End Select
		
	End Sub
	
	Private Sub Obj_OnChangedConnectState(ByVal bIsConnected As Boolean)
		If bIsConnected = False Then
			Me.Text = mTitle & " [Chroma " & gstrVPGModel & " Disconnected]"
		End If
	End Sub
	
	Private Sub ChangeTiming(ByRef Tim As String)
		Dim bNo(1) As Byte
		
		bNo(0) = (CShort(Tim) And &HFF00) \ 256
		bNo(1) = CShort(Tim) And &HFF
		
		ivpg.ExecuteCmd(VPGCtrl.VPG_CMD.VPG_CMD_CM_DOWNLOAD, VPGCtrl.VPG_SCMD.VPG_SCMD_SCM_CTL_RUNTIM, bNo, False)
	End Sub
	
	Private Sub ChangePattern(ByRef Ptn As String)
		Dim bNo(1) As Byte
		
		bNo(0) = (CShort(Ptn) And &HFF00) \ 256
		bNo(1) = CShort(Ptn) And &HFF
		
		ivpg.RunKey(VPGCtrl.VPG_KEY.VPG_KEY_CKEY_OUT)
		ivpg.ExecuteCmd(VPGCtrl.VPG_CMD.VPG_CMD_CM_DOWNLOAD, VPGCtrl.VPG_SCMD.VPG_SCMD_SCM_CTL_RUNPTN, bNo, False)
	End Sub
End Class