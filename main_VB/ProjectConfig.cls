VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProjectConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************
' Class module for handling config.xml of the
' application.
'**********************************************

Option Explicit

Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Private Type udtConfigData
    CommMode As CommunicationMode
    strComBaud As String
    intComID As Integer
    strInputSource As String
    lngDelayMs As Long
    intChannelNum As Integer
    intBarCodeLen As Integer
    intLvSpec As Integer
    strVPGModel As String
    strVPGTiming As String
    strVPG100IRE As String
    strVPG80IRE As String
    strVPG20IRE As String
    lngI2cClockRate As Long
    bolEnableCool2 As Boolean
    bolEnableCool1 As Boolean
    bolEnableNormal As Boolean
    bolEnableWarm1 As Boolean
    bolEnableWarm2 As Boolean
    bolEnableChkColor As Boolean
    bolEnableAdjOffset As Boolean
End Type

Private mConfigData As udtConfigData
Private mstrConfigFilePath As String

Private Sub Class_Initialize()
    mstrConfigFilePath = App.Path & "\" & gstrCurProjName & "\config.xml"

    mConfigData.CommMode = modeUART
    mConfigData.strComBaud = "115200"
    mConfigData.intComID = 1
    mConfigData.strInputSource = "HDMI1"
    mConfigData.lngDelayMs = 500
    mConfigData.intChannelNum = 1
    mConfigData.intBarCodeLen = 1
    mConfigData.intLvSpec = 280
    mConfigData.strVPGModel = "22294"
    mConfigData.strVPGTiming = "69"
    mConfigData.strVPG100IRE = "101"
    mConfigData.strVPG80IRE = "103"
    mConfigData.strVPG20IRE = "109"
    mConfigData.lngI2cClockRate = 50
    mConfigData.bolEnableCool2 = False
    mConfigData.bolEnableCool1 = True
    mConfigData.bolEnableNormal = True
    mConfigData.bolEnableWarm1 = True
    mConfigData.bolEnableWarm2 = False
    mConfigData.bolEnableChkColor = True
    mConfigData.bolEnableAdjOffset = False
End Sub

Public Sub LoadConfigData()
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean

    If Not CBool(PathFileExists(mstrConfigFilePath)) Then
        MsgBox "Cannot open " & mstrConfigFilePath & " file."
        End
    End If
    
    success = xmlDoc.Load(mstrConfigFilePath)
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        If xmlDoc.selectSingleNode("/config/communication").selectSingleNode("@mode").Text = "UART" Then
            mConfigData.CommMode = modeUART
        ElseIf xmlDoc.selectSingleNode("/config/communication").selectSingleNode("@mode").Text = "Network" Then
            mConfigData.CommMode = modeNetwork
        ElseIf xmlDoc.selectSingleNode("/config/communication").selectSingleNode("@mode").Text = "I2C" Then
            mConfigData.CommMode = modeI2c
        End If

        mConfigData.strComBaud = xmlDoc.selectSingleNode("/config/communication/common").selectSingleNode("@baud").Text
        mConfigData.intComID = val(xmlDoc.selectSingleNode("/config/communication/common").selectSingleNode("@id").Text)
        mConfigData.lngI2cClockRate = val(xmlDoc.selectSingleNode("/config/communication/i2c").selectSingleNode("@clockrate").Text)
        mConfigData.strInputSource = xmlDoc.selectSingleNode("/config/input_source").Text
        mConfigData.lngDelayMs = val(xmlDoc.selectSingleNode("/config/delayms").Text)
        mConfigData.intChannelNum = val(xmlDoc.selectSingleNode("/config/channel_number").Text)
        mConfigData.intBarCodeLen = val(xmlDoc.selectSingleNode("/config/length_bar_code").Text)
        mConfigData.intLvSpec = val(xmlDoc.selectSingleNode("/config/Lv_spec").Text)
        mConfigData.strVPGModel = xmlDoc.selectSingleNode("/config/VPG/model").Text
        mConfigData.strVPGTiming = xmlDoc.selectSingleNode("/config/VPG/timing").Text
        mConfigData.strVPG100IRE = xmlDoc.selectSingleNode("/config/VPG/IRE100").Text
        mConfigData.strVPG80IRE = xmlDoc.selectSingleNode("/config/VPG/IRE80").Text
        mConfigData.strVPG20IRE = xmlDoc.selectSingleNode("/config/VPG/IRE20").Text
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
        MsgBox xmlDoc.parseError.reason
    Else
        If mConfigData.CommMode = modeUART Then
            xmlDoc.selectSingleNode("/config/communication").selectSingleNode("@mode").Text = "UART"
        ElseIf mConfigData.CommMode = modeNetwork Then
            xmlDoc.selectSingleNode("/config/communication").selectSingleNode("@mode").Text = "Network"
        ElseIf mConfigData.CommMode = modeI2c Then
            xmlDoc.selectSingleNode("/config/communication").selectSingleNode("@mode").Text = "I2C"
        End If
        
        xmlDoc.selectSingleNode("/config/communication/common").selectSingleNode("@baud").Text = mConfigData.strComBaud
        xmlDoc.selectSingleNode("/config/communication/common").selectSingleNode("@id").Text = CStr(mConfigData.intComID)
        xmlDoc.selectSingleNode("/config/input_source").Text = mConfigData.strInputSource
        xmlDoc.selectSingleNode("/config/delayms").Text = CStr(mConfigData.lngDelayMs)
        xmlDoc.selectSingleNode("/config/channel_number").Text = CStr(mConfigData.intChannelNum)
        xmlDoc.selectSingleNode("/config/length_bar_code").Text = CStr(mConfigData.intBarCodeLen)
        xmlDoc.selectSingleNode("/config/Lv_spec").Text = CStr(mConfigData.intLvSpec)
        xmlDoc.selectSingleNode("/config/VPG/model").Text = mConfigData.strVPGModel
        xmlDoc.selectSingleNode("/config/VPG/timing").Text = mConfigData.strVPGTiming
        xmlDoc.selectSingleNode("/config/VPG/IRE100").Text = mConfigData.strVPG100IRE
        xmlDoc.selectSingleNode("/config/VPG/IRE80").Text = mConfigData.strVPG80IRE
        xmlDoc.selectSingleNode("/config/VPG/IRE20").Text = mConfigData.strVPG20IRE
        
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
        
        xmlDoc.Save mstrConfigFilePath
    End If
End Sub

Public Property Get CommMode() As CommunicationMode
    CommMode = mConfigData.CommMode
End Property

Public Property Let CommMode(enuCommMode As CommunicationMode)
    mConfigData.CommMode = enuCommMode
End Property

Public Property Get ComBaud() As String
    ComBaud = mConfigData.strComBaud
End Property

Public Property Let ComBaud(strComBaud As String)
    mConfigData.strComBaud = strComBaud
End Property

Public Property Get ComID() As Integer
    ComID = mConfigData.intComID
End Property

Public Property Let ComID(intComID As Integer)
    mConfigData.intComID = intComID
End Property

Public Property Get inputSource() As String
    inputSource = mConfigData.strInputSource
End Property

Public Property Let inputSource(strInputSource As String)
    mConfigData.strInputSource = strInputSource
End Property

Public Property Get DelayMS() As Long
    DelayMS = mConfigData.lngDelayMs
End Property

Public Property Let DelayMS(lngDelayMs As Long)
    mConfigData.lngDelayMs = lngDelayMs
End Property

Public Property Get ChannelNum() As Integer
    ChannelNum = mConfigData.intChannelNum
End Property

Public Property Let ChannelNum(intChannelNum As Integer)
    mConfigData.intChannelNum = intChannelNum
End Property

Public Property Get BarCodeLen() As Integer
    BarCodeLen = mConfigData.intBarCodeLen
End Property

Public Property Let BarCodeLen(intBarCodeLen As Integer)
    mConfigData.intBarCodeLen = intBarCodeLen
End Property

Public Property Get LvSpec() As Integer
    LvSpec = mConfigData.intLvSpec
End Property

Public Property Let LvSpec(intLvSpec As Integer)
    mConfigData.intLvSpec = intLvSpec
End Property

Public Property Get VPGModel() As String
    VPGModel = mConfigData.strVPGModel
End Property

Public Property Let VPGModel(strVPGModel As String)
    mConfigData.strVPGModel = strVPGModel
End Property

Public Property Get VPGTiming() As String
    VPGTiming = mConfigData.strVPGTiming
End Property

Public Property Let VPGTiming(strVPGTiming As String)
    mConfigData.strVPGTiming = strVPGTiming
End Property

Public Property Get VPG100IRE() As String
    VPG100IRE = mConfigData.strVPG100IRE
End Property

Public Property Let VPG100IRE(strVPG100IRE As String)
    mConfigData.strVPG100IRE = strVPG100IRE
End Property

Public Property Get VPG80IRE() As String
    VPG80IRE = mConfigData.strVPG80IRE
End Property

Public Property Let VPG80IRE(strVPG80IRE As String)
    mConfigData.strVPG80IRE = strVPG80IRE
End Property

Public Property Get VPG20IRE() As String
    VPG20IRE = mConfigData.strVPG20IRE
End Property

Public Property Let VPG20IRE(strVPG20IRE As String)
    mConfigData.strVPG20IRE = strVPG20IRE
End Property

Public Property Get I2cClockRate() As String
    I2cClockRate = mConfigData.lngI2cClockRate
End Property

Public Property Let I2cClockRate(lngI2cClockRate As String)
    mConfigData.lngI2cClockRate = lngI2cClockRate
End Property

Public Property Get EnableCool2() As Boolean
    EnableCool2 = mConfigData.bolEnableCool2
End Property

Public Property Let EnableCool2(bolEnableCool2 As Boolean)
    mConfigData.bolEnableCool2 = bolEnableCool2
End Property

Public Property Get EnableCool1() As Boolean
    EnableCool1 = mConfigData.bolEnableCool1
End Property

Public Property Let EnableCool1(bolEnableCool1 As Boolean)
    mConfigData.bolEnableCool1 = bolEnableCool1
End Property

Public Property Get EnableNormal() As Boolean
    EnableNormal = mConfigData.bolEnableNormal
End Property

Public Property Let EnableNormal(bolEnableNormal As Boolean)
    mConfigData.bolEnableNormal = bolEnableNormal
End Property

Public Property Get EnableWarm1() As Boolean
    EnableWarm1 = mConfigData.bolEnableWarm1
End Property

Public Property Let EnableWarm1(bolEnableWarm1 As Boolean)
    mConfigData.bolEnableWarm1 = bolEnableWarm1
End Property

Public Property Get EnableWarm2() As Boolean
    EnableWarm2 = mConfigData.bolEnableWarm2
End Property

Public Property Let EnableWarm2(bolEnableWarm2 As Boolean)
    mConfigData.bolEnableWarm2 = bolEnableWarm2
End Property

Public Property Get EnableChkColor() As Boolean
    EnableChkColor = mConfigData.bolEnableChkColor
End Property

Public Property Let EnableChkColor(bolEnableChkColor As Boolean)
    mConfigData.bolEnableChkColor = bolEnableChkColor
End Property

Public Property Get EnableAdjOffset() As Boolean
    EnableAdjOffset = mConfigData.bolEnableAdjOffset
End Property

Public Property Let EnableAdjOffset(bolEnableAdjOffset As Boolean)
    mConfigData.bolEnableAdjOffset = bolEnableAdjOffset
End Property
