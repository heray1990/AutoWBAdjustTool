Option Strict Off
Option Explicit On
Friend Class frmSetData
	Inherits System.Windows.Forms.Form
	
	
	Private Sub frmSetData_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Label1.Text = gstrCurProjName
		
		txtChannel.Text = CStr(Ca210ChannelNO)
		txtSNLen.Text = CStr(gintBarCodeLen)
		txtLvSpec.Text = CStr(maxBrightnessSpec)
		cmbInputSource.Text = setTVInputSource & CStr(setTVInputSourcePortNum)
		txtDelay.Text = CStr(delayTime)
		
		cmbComBaud.Text = CStr(setTVCurrentComBaud)
		cmbComID.Text = "COM" & CStr(setTVCurrentComID)
		For i = 1 To 20
			cmbComID.Items.Add("COM" & i)
		Next i
		
		cmbComBaud.Items.Add("9600")
		cmbComBaud.Items.Add("19200")
		cmbComBaud.Items.Add("38400")
		cmbComBaud.Items.Add("57600")
		cmbComBaud.Items.Add("115200")
		
		cmbChromaModel.Text = gstrVPGModel
		
		If isUartMode Then
			optUart.Checked = True
			optNetwork.Checked = False
			cmbComBaud.Enabled = True
			cmbComID.Enabled = True
		Else
			optUart.Checked = False
			optNetwork.Checked = True
			cmbComBaud.Enabled = False
			cmbComID.Enabled = False
		End If
		
		If isAdjustCool2 Then
			Check1.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			Check1.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
		If isAdjustCool1 Then
			Check2.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			Check2.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
		If isAdjustNormal Then
			Check3.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			Check3.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
		If isAdjustWarm1 Then
			Check4.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			Check4.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
		If isAdjustWarm2 Then
			Check5.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			Check5.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
		If isCheckColorTemp Then
			Check6.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			Check6.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
		If isAdjustOffset Then
			Check7.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			Check7.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
	End Sub
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim clsSaveConfigData As ProjectConfig
		
		clsSaveConfigData = New ProjectConfig
		
		If Check1.CheckState = 1 Then clsSaveConfigData.EnableCool2 = True
		If Check1.CheckState = 0 Then clsSaveConfigData.EnableCool2 = False
		If Check2.CheckState = 1 Then clsSaveConfigData.EnableCool1 = True
		If Check2.CheckState = 0 Then clsSaveConfigData.EnableCool1 = False
		If Check3.CheckState = 1 Then clsSaveConfigData.EnableNormal = True
		If Check3.CheckState = 0 Then clsSaveConfigData.EnableNormal = False
		If Check4.CheckState = 1 Then clsSaveConfigData.EnableWarm1 = True
		If Check4.CheckState = 0 Then clsSaveConfigData.EnableWarm1 = False
		If Check5.CheckState = 1 Then clsSaveConfigData.EnableWarm2 = True
		If Check5.CheckState = 0 Then clsSaveConfigData.EnableWarm2 = False
		If Check6.CheckState = 1 Then clsSaveConfigData.EnableChkColor = True
		If Check6.CheckState = 0 Then clsSaveConfigData.EnableChkColor = False
		If Check7.CheckState = 1 Then clsSaveConfigData.EnableAdjOffset = True
		If Check7.CheckState = 0 Then clsSaveConfigData.EnableAdjOffset = False
		
		clsSaveConfigData.LvSpec = Val(txtLvSpec.Text)
		clsSaveConfigData.BarCodeLen = Val(txtSNLen.Text)
		
		If optUart.Checked = True Then
			clsSaveConfigData.CommMode = DeclareVariables.CommunicationMode.modeUART
		ElseIf optNetwork.Checked = True Then 
			clsSaveConfigData.CommMode = DeclareVariables.CommunicationMode.modeNetwork
		Else
			clsSaveConfigData.CommMode = DeclareVariables.CommunicationMode.modeUART
		End If
		
		clsSaveConfigData.ComBaud = cmbComBaud.Text
		clsSaveConfigData.ComID = Val(Replace(cmbComID.Text, "COM", ""))
		clsSaveConfigData.ChannelNum = Val(txtChannel.Text)
		clsSaveConfigData.DelayMS = Val(txtDelay.Text)
		clsSaveConfigData.inputSource = cmbInputSource.Text
		clsSaveConfigData.VPGModel = cmbChromaModel.Text
		
		clsSaveConfigData.SaveConfigData()
		
		'UPGRADE_NOTE: Object clsSaveConfigData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		clsSaveConfigData = Nothing
		
		Me.Close()
		
		Form1.subInitInterface()
		Form1.Show()
	End Sub
	
	'UPGRADE_WARNING: Event optNetwork.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optNetwork_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optNetwork.CheckedChanged
		If eventSender.Checked Then
			cmbComBaud.Enabled = False
			cmbComID.Enabled = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optUart.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optUart_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optUart.CheckedChanged
		If eventSender.Checked Then
			cmbComBaud.Enabled = True
			cmbComID.Enabled = True
		End If
	End Sub
End Class