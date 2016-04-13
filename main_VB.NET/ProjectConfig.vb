Option Strict Off
Option Explicit On
Friend Class ProjectConfig
	'**********************************************
	' Class module for handling config.xml of the
	' application.
	'**********************************************
	
	
	Private Declare Function PathFileExists Lib "shlwapi.dll"  Alias "PathFileExistsA"(ByVal pszPath As String) As Integer
	
	Private Structure udtConfigData
		Dim CommMode As DeclareVariables.CommunicationMode
		Dim strComBaud As String
		Dim intComID As Short
		Dim strInputSource As String
		Dim lngDelayMs As Integer
		Dim intChannelNum As Short
		Dim intBarCodeLen As Short
		Dim intLvSpec As Short
		Dim strVPGModel As String
		Dim bolEnableCool2 As Boolean
		Dim bolEnableCool1 As Boolean
		Dim bolEnableNormal As Boolean
		Dim bolEnableWarm1 As Boolean
		Dim bolEnableWarm2 As Boolean
		Dim bolEnableChkColor As Boolean
		Dim bolEnableAdjOffset As Boolean
	End Structure
	
	Private mConfigData As udtConfigData
	Private mstrConfigFilePath As String
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		mstrConfigFilePath = My.Application.Info.DirectoryPath & "\" & gstrCurProjName & "\config.xml"
		
		mConfigData.CommMode = DeclareVariables.CommunicationMode.modeUART
		mConfigData.strComBaud = "115200"
		mConfigData.intComID = 1
		mConfigData.strInputSource = "HDMI1"
		mConfigData.lngDelayMs = 500
		mConfigData.intChannelNum = 1
		mConfigData.intBarCodeLen = 1
		mConfigData.intLvSpec = 280
		mConfigData.strVPGModel = "22294"
		mConfigData.bolEnableCool2 = False
		mConfigData.bolEnableCool1 = True
		mConfigData.bolEnableNormal = True
		mConfigData.bolEnableWarm1 = True
		mConfigData.bolEnableWarm2 = False
		mConfigData.bolEnableChkColor = True
		mConfigData.bolEnableAdjOffset = False
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Public Sub LoadConfigData()
		Dim xmlDoc As New MSXML2.DOMDocument
		Dim success As Boolean
		
		If Not CBool(PathFileExists(mstrConfigFilePath)) Then
			MsgBox("Cannot open " & mstrConfigFilePath & " file.")
			End
		End If
		
		success = xmlDoc.Load(mstrConfigFilePath)
		
		If success = False Then
			MsgBox(xmlDoc.parseError.reason)
		Else
			If xmlDoc.selectSingleNode("/config/communication").selectSingleNode("@mode").Text = "UART" Then
				mConfigData.CommMode = DeclareVariables.CommunicationMode.modeUART
			Else
				mConfigData.CommMode = DeclareVariables.CommunicationMode.modeNetwork
			End If
			
			mConfigData.strComBaud = xmlDoc.selectSingleNode("/config/communication/common").selectSingleNode("@baud").Text
			mConfigData.intComID = Val(xmlDoc.selectSingleNode("/config/communication/common").selectSingleNode("@id").Text)
			mConfigData.strInputSource = xmlDoc.selectSingleNode("/config/input_source").Text
			mConfigData.lngDelayMs = Val(xmlDoc.selectSingleNode("/config/delayms").Text)
			mConfigData.intChannelNum = Val(xmlDoc.selectSingleNode("/config/channel_number").Text)
			mConfigData.intBarCodeLen = Val(xmlDoc.selectSingleNode("/config/length_bar_code").Text)
			mConfigData.intLvSpec = Val(xmlDoc.selectSingleNode("/config/Lv_spec").Text)
			mConfigData.strVPGModel = xmlDoc.selectSingleNode("/config/VPG_model").Text
		End If
		
		If xmlDoc.selectSingleNode("/config/cool_2").Text = "True" Then
			mConfigData.bolEnableCool2 = True
		Else
			mConfigData.bolEnableCool2 = False
		End If
		
		If xmlDoc.selectSingleNode("/config/cool_1").Text = "True" Then
			mConfigData.bolEnableCool1 = True
		Else
			mConfigData.bolEnableCool1 = False
		End If
		
		If xmlDoc.selectSingleNode("/config/normal").Text = "True" Then
			mConfigData.bolEnableNormal = True
		Else
			mConfigData.bolEnableNormal = False
		End If
		
		If xmlDoc.selectSingleNode("/config/warm_1").Text = "True" Then
			mConfigData.bolEnableWarm1 = True
		Else
			mConfigData.bolEnableWarm1 = False
		End If
		
		If xmlDoc.selectSingleNode("/config/warm_2").Text = "True" Then
			mConfigData.bolEnableWarm2 = True
		Else
			mConfigData.bolEnableWarm2 = False
		End If
		
		If xmlDoc.selectSingleNode("/config/check_color").Text = "True" Then
			mConfigData.bolEnableChkColor = True
		Else
			mConfigData.bolEnableChkColor = False
		End If
		
		If xmlDoc.selectSingleNode("/config/adjust_offset").Text = "True" Then
			mConfigData.bolEnableAdjOffset = True
		Else
			mConfigData.bolEnableAdjOffset = False
		End If
	End Sub
	
	Public Sub SaveConfigData()
		Dim xmlDoc As New MSXML2.DOMDocument
		Dim success As Boolean
		
		success = xmlDoc.Load(mstrConfigFilePath)
		
		If success = False Then
			MsgBox(xmlDoc.parseError.reason)
		Else
			If mConfigData.CommMode = DeclareVariables.CommunicationMode.modeUART Then
				xmlDoc.selectSingleNode("/config/communication").selectSingleNode("@mode").Text = "UART"
			Else
				xmlDoc.selectSingleNode("/config/communication").selectSingleNode("@mode").Text = "Network"
			End If
			
			xmlDoc.selectSingleNode("/config/communication/common").selectSingleNode("@baud").Text = mConfigData.strComBaud
			xmlDoc.selectSingleNode("/config/communication/common").selectSingleNode("@id").Text = CStr(mConfigData.intComID)
			xmlDoc.selectSingleNode("/config/input_source").Text = mConfigData.strInputSource
			xmlDoc.selectSingleNode("/config/delayms").Text = CStr(mConfigData.lngDelayMs)
			xmlDoc.selectSingleNode("/config/channel_number").Text = CStr(mConfigData.intChannelNum)
			xmlDoc.selectSingleNode("/config/length_bar_code").Text = CStr(mConfigData.intBarCodeLen)
			xmlDoc.selectSingleNode("/config/Lv_spec").Text = CStr(mConfigData.intLvSpec)
			xmlDoc.selectSingleNode("/config/VPG_model").Text = mConfigData.strVPGModel
			
			If mConfigData.bolEnableCool2 Then
				xmlDoc.selectSingleNode("/config/cool_2").Text = "True"
			Else
				xmlDoc.selectSingleNode("/config/cool_2").Text = "False"
			End If
			
			If mConfigData.bolEnableCool1 Then
				xmlDoc.selectSingleNode("/config/cool_1").Text = "True"
			Else
				xmlDoc.selectSingleNode("/config/cool_1").Text = "False"
			End If
			
			If mConfigData.bolEnableNormal Then
				xmlDoc.selectSingleNode("/config/normal").Text = "True"
			Else
				xmlDoc.selectSingleNode("/config/normal").Text = "False"
			End If
			
			If mConfigData.bolEnableWarm1 Then
				xmlDoc.selectSingleNode("/config/warm_1").Text = "True"
			Else
				xmlDoc.selectSingleNode("/config/warm_1").Text = "False"
			End If
			
			If mConfigData.bolEnableWarm2 Then
				xmlDoc.selectSingleNode("/config/warm_2").Text = "True"
			Else
				xmlDoc.selectSingleNode("/config/warm_2").Text = "False"
			End If
			
			If mConfigData.bolEnableChkColor Then
				xmlDoc.selectSingleNode("/config/check_color").Text = "True"
			Else
				xmlDoc.selectSingleNode("/config/check_color").Text = "False"
			End If
			
			If mConfigData.bolEnableAdjOffset Then
				xmlDoc.selectSingleNode("/config/adjust_offset").Text = "True"
			Else
				xmlDoc.selectSingleNode("/config/adjust_offset").Text = "False"
			End If
			
			xmlDoc.Save(mstrConfigFilePath)
		End If
	End Sub
	
	
	Public Property CommMode() As DeclareVariables.CommunicationMode
		Get
			CommMode = mConfigData.CommMode
		End Get
		Set(ByVal Value As DeclareVariables.CommunicationMode)
			mConfigData.CommMode = Value
		End Set
	End Property
	
	
	Public Property ComBaud() As String
		Get
			ComBaud = mConfigData.strComBaud
		End Get
		Set(ByVal Value As String)
			mConfigData.strComBaud = Value
		End Set
	End Property
	
	
	Public Property ComID() As Short
		Get
			ComID = mConfigData.intComID
		End Get
		Set(ByVal Value As Short)
			mConfigData.intComID = Value
		End Set
	End Property
	
	
	Public Property inputSource() As String
		Get
			inputSource = mConfigData.strInputSource
		End Get
		Set(ByVal Value As String)
			mConfigData.strInputSource = Value
		End Set
	End Property
	
	
	Public Property DelayMS() As Integer
		Get
			DelayMS = mConfigData.lngDelayMs
		End Get
		Set(ByVal Value As Integer)
			mConfigData.lngDelayMs = Value
		End Set
	End Property
	
	
	Public Property ChannelNum() As Short
		Get
			ChannelNum = mConfigData.intChannelNum
		End Get
		Set(ByVal Value As Short)
			mConfigData.intChannelNum = Value
		End Set
	End Property
	
	
	Public Property BarCodeLen() As Short
		Get
			BarCodeLen = mConfigData.intBarCodeLen
		End Get
		Set(ByVal Value As Short)
			mConfigData.intBarCodeLen = Value
		End Set
	End Property
	
	
	Public Property LvSpec() As Short
		Get
			LvSpec = mConfigData.intLvSpec
		End Get
		Set(ByVal Value As Short)
			mConfigData.intLvSpec = Value
		End Set
	End Property
	
	
	Public Property VPGModel() As String
		Get
			VPGModel = mConfigData.strVPGModel
		End Get
		Set(ByVal Value As String)
			mConfigData.strVPGModel = Value
		End Set
	End Property
	
	
	Public Property EnableCool2() As Boolean
		Get
			EnableCool2 = mConfigData.bolEnableCool2
		End Get
		Set(ByVal Value As Boolean)
			mConfigData.bolEnableCool2 = Value
		End Set
	End Property
	
	
	Public Property EnableCool1() As Boolean
		Get
			EnableCool1 = mConfigData.bolEnableCool1
		End Get
		Set(ByVal Value As Boolean)
			mConfigData.bolEnableCool1 = Value
		End Set
	End Property
	
	
	Public Property EnableNormal() As Boolean
		Get
			EnableNormal = mConfigData.bolEnableNormal
		End Get
		Set(ByVal Value As Boolean)
			mConfigData.bolEnableNormal = Value
		End Set
	End Property
	
	
	Public Property EnableWarm1() As Boolean
		Get
			EnableWarm1 = mConfigData.bolEnableWarm1
		End Get
		Set(ByVal Value As Boolean)
			mConfigData.bolEnableWarm1 = Value
		End Set
	End Property
	
	
	Public Property EnableWarm2() As Boolean
		Get
			EnableWarm2 = mConfigData.bolEnableWarm2
		End Get
		Set(ByVal Value As Boolean)
			mConfigData.bolEnableWarm2 = Value
		End Set
	End Property
	
	
	Public Property EnableChkColor() As Boolean
		Get
			EnableChkColor = mConfigData.bolEnableChkColor
		End Get
		Set(ByVal Value As Boolean)
			mConfigData.bolEnableChkColor = Value
		End Set
	End Property
	
	
	Public Property EnableAdjOffset() As Boolean
		Get
			EnableAdjOffset = mConfigData.bolEnableAdjOffset
		End Get
		Set(ByVal Value As Boolean)
			mConfigData.bolEnableAdjOffset = Value
		End Set
	End Property
End Class